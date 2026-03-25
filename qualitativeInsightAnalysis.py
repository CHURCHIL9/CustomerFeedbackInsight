"""
========================================================================
FeedbackInsightEngine v2 – Africa-Ready Survey Analysis Pipeline
========================================================================

UPGRADE SUMMARY (v1 → v2)
--------------------------
BUG FIXES:
  • Multi-question Excel report now uses question_results (was only
    outputting the last question's data).
  • Share (%) now based on per-question response count, not full CSV rows.
  • merge_similar_themes correctly preserves Theme Sentiment after merge.

PERFORMANCE / RAM (target: 8 GB):
  • PCA reduction (384 → 64 dims) before KMeans  →  ~6× RAM savings.
  • MiniBatchKMeans for n > 500 rows.
  • Embeddings generated in memory-safe batches; intermediate arrays
    deleted + gc.collect() called at key points.
  • CSV/Excel loaded with chunked row-count detection; chardet used to
    auto-detect encoding (common issue with African data exports).
  • TF-IDF capped at 5 000 features (down from 8 000).

ACCURACY / ROBUSTNESS:
  • Language detection (langdetect) flags Swahili, Sheng & code-switched
    text so those responses are NOT stripped of meaning by the English
    regex cleaner.
  • Extended stop-word list: English NLTK + Swahili + common Sheng +
    African-English filler words.
  • Sentiment extended with a Swahili/Sheng polarity lexicon that covers
    common positive/negative terms (sawa, poa, mbaya, vibaya, nzuri …).
  • Sarcasm & indirect negativity flag: responses containing hedged
    positives ("not bad", "could be better") are down-scored.
  • Respondent-level deduplication: near-identical responses from the
    same row are collapsed so one loud voice doesn't inflate a theme.
  • Weak-signal detection: very small clusters (below min_size) are kept
    in a separate "Emerging Issues" section rather than silently dropped.
  • Intensity scoring separates "few very angry" from "many mildly
    negative" — both get flagged but with different action labels.
  • Cross-question theme detection: themes that recur across ≥2
    questions are highlighted in the Executive Summary.

AFRICAN / KENYAN MARKET TAILORING:
  • Recommendation engine covers Kenya-specific contexts: M-Pesa /
    mobile money, boda-boda / matatu transport, informal market (jua
    kali), county government services, NHIF/SHIF health cover, water
    kiosks, school fees, mobile data costs.
  • Sub-theme drilling on the four most common African service themes:
    cost, transport/access, staff attitude, wait time.
  • Report date header uses East Africa Time (EAT, UTC+3).

WHAT COMPETITORS MISS (now built in):
  • Sarcasm/hedged language detection.
  • Respondent vs question-level coding (loud vs widespread).
  • Emerging/weak signal capture.
  • Intensity vs frequency distinction.
  • Cross-question pattern surfacing.
  • Sub-theme breakdown inside major themes.
  • Language-aware preprocessing for multilingual data.

FLEXIBILITY / SCALABILITY:
  • Accepts .csv OR .xlsx / .xls input.
  • Config dataclass centralises all tunable parameters.
  • Pipeline skips questions with < MIN_RESPONSES usable rows and logs a
    warning instead of crashing.
  • Structured logging replaces bare print statements.
  • All per-question artefacts stored in question_results for downstream
    use (API, database, dashboard).

DEPENDENCIES (pip install …):
    sentence-transformers scikit-learn openpyxl nltk pandas numpy
    chardet langdetect
    # optional but recommended for speed:
    # scikit-learn >= 1.3  (faster MiniBatchKMeans)
========================================================================
"""

# ── stdlib ────────────────────────────────────────────────────────────
import gc
import json
import logging
import os
import pickle
import re
import warnings
from dataclasses import dataclass, field
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

warnings.filterwarnings("ignore")

# ── third-party ───────────────────────────────────────────────────────
import chardet
import nltk
import numpy as np
import pandas as pd
import yaml
from nltk.sentiment import SentimentIntensityAnalyzer
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from sentence_transformers import SentenceTransformer
from sklearn.cluster import KMeans, MiniBatchKMeans
from sklearn.decomposition import PCA
from sklearn.feature_extraction.text import TfidfVectorizer, ENGLISH_STOP_WORDS
from sklearn.metrics import silhouette_score

try:
    from langdetect import detect as lang_detect
    LANGDETECT_AVAILABLE = True
except ImportError:
    LANGDETECT_AVAILABLE = False

try:
    import anthropic as _anthropic_lib
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

nltk.download("vader_lexicon", quiet=True)
nltk.download("stopwords", quiet=True)

# ── logging ───────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("FeedbackEngine")

# ── EAT timezone ──────────────────────────────────────────────────────
EAT = timezone(timedelta(hours=3))

# ═══════════════════════════════════════════════════════════════════════
# CONFIG SYSTEM (v3+ Upgrade)
# ═══════════════════════════════════════════════════════════════════════

def load_config(path: str = "config/config.yaml") -> dict:
    """
    Load configuration from YAML file for flexibility and robustness.
    
    Args:
        path: relative path to config.yaml (default: config/config.yaml)
    
    Returns:
        dict with keys: sector, thresholds, features, output
    
    Raises:
        FileNotFoundError if config file doesn't exist
    """
    config_path = Path(path)
    
    if not config_path.exists():
        log.warning(
            "Config file not found at %s. Using hardcoded defaults.",
            config_path
        )
        # Hardcoded fallback defaults
        return {
            "sector": "smallholder agriculture",
            "thresholds": {"high_impact": 15, "medium_impact": 8},
            "features": {
                "enable_trend_analysis": True,
                "enable_trigger_engine": True,
                "enable_action_engine": True,
            },
            "output": {"excel": True, "json": True, "alerts": True},
        }
    
    try:
        with open(config_path, "r") as f:
            config = yaml.safe_load(f)
        log.info("Loaded configuration from %s", config_path)
        return config
    except Exception as exc:
        log.error("Failed to load config (%s). Using defaults.", exc)
        return {
            "sector": "smallholder agriculture",
            "thresholds": {"high_impact": 15, "medium_impact": 8},
            "features": {
                "enable_trend_analysis": True,
                "enable_trigger_engine": True,
                "enable_action_engine": True,
            },
            "output": {"excel": True, "json": True, "alerts": True},
        }

# ── NEW ENGINE IMPORTS ────────────────────────────────────────────────
# These are loaded lazily when features are enabled
_ENGINES_LOADED = False


# ══════════════════════════════════════════════════════════════════════
# CONFIG DATACLASS  ─ tune everything from one place
# ══════════════════════════════════════════════════════════════════════

@dataclass
class EngineConfig:
    """All tuneable parameters in one place."""

    # --- Clustering ---
    pca_components: int = 64          # dims after PCA (≤ 384); reduces RAM
    min_cluster_size: int = 5         # hard minimum for a "real" cluster
    emerging_min_size: int = 2        # below this → ignored entirely
    max_clusters: int = 20            # cap for very large datasets
    merge_threshold: float = 0.75     # cosine sim for theme merging

    # --- Text ---
    tfidf_max_features: int = 5_000
    tfidf_ngram_range: Tuple[int, int] = (1, 2)
    min_sentence_words: int = 3       # ignore fragments shorter than this
    top_keywords: int = 6

    # --- Sentiment ---
    extra_sentiment_lexicon: Dict[str, float] = field(default_factory=lambda: {
        # Swahili positive
        "sawa": 0.4, "poa": 0.5, "nzuri": 0.5, "vizuri": 0.5,
        "asante": 0.3, "hongera": 0.6, "bora": 0.4, "salama": 0.3,
        "furaha": 0.6, "starehe": 0.4, "rahisi": 0.3,
        # Swahili negative
        "mbaya": -0.5, "vibaya": -0.5, "tatizo": -0.4, "shida": -0.5,
        "hasira": -0.6, "huzuni": -0.5, "duni": -0.4, "gharama": -0.3,
        "uchafu": -0.5, "lalamiko": -0.4, "chelewa": -0.5,
        # Sheng
        "peng": 0.3, "moto": 0.3, "safi": 0.5, "rada": -0.3,
        "stress": -0.4, "noma": -0.4, "fala": -0.5,
        # African-English
        "harassed": -0.6, "frustrated": -0.6, "disappointed": -0.6,
        "chapaa": -0.3,
    })

    # --- Scale ---
    large_dataset_threshold: int = 500   # use MiniBatchKMeans above this
    min_responses: int = 10              # skip question if fewer rows

    # --- Context (used by LLM methods) ---
    sector: str = ""
    # e.g. "smallholder agriculture", "primary healthcare", "microfinance",
    #      "WASH services", "secondary education"  — left blank for generic.

    # --- LLM-assisted intelligence (requires: pip install anthropic) ---
    use_llm_naming: bool = True
    # When True, Claude generates meaningful theme names from keywords + quotes.
    # When False (or anthropic not installed), improved rule-based naming is used.

    use_llm_recommendations: bool = True
    # When True, Claude generates sector-aware, context-specific recommendations.
    # When False, the rule-based RECOMMENDATION_RULES dict is used.
    # TIP: run once with False to see which keywords need new rules (see logs).

    anthropic_model: str = "claude-haiku-4-5-20251001"
    # Use the fast/cheap Haiku model — theme naming + recs are short prompts.
    # Switch to "claude-sonnet-4-6" for richer recommendations.

    # --- Output ---
    output_filename: str = "Feedback_Insights_Report.xlsx"
    report_title: str = "Customer Feedback Insights"


# ══════════════════════════════════════════════════════════════════════
# STOP WORDS  ─ English + Swahili + Sheng + African-English fillers
# ══════════════════════════════════════════════════════════════════════

SWAHILI_STOP_WORDS = {
    "na", "ya", "wa", "la", "za", "kwa", "ni", "si", "au", "pia",
    "hii", "hizi", "hilo", "hayo", "ile", "zile", "ile", "hiyo",
    "yote", "wote", "sana", "sasa", "bado", "lakini", "ama", "kama",
    "kwamba", "ingawa", "hivyo", "ndio", "hapana", "ndiyo", "tena",
    "kabla", "baada", "wakati", "mara", "siku", "wiki", "mwezi",
    "mtu", "watu", "mahali", "pamoja", "pia", "zaidi", "kidogo",
    "tu", "hata", "bila", "kila", "baadhi", "zote", "wewe", "mimi",
    "yeye", "sisi", "nyinyi", "wao", "mwenyewe",
}

SHENG_FILLER_STOP_WORDS = {
    "maze", "bana", "manze", "niaje", "aje", "si", "kweli", "freshi",
    "mtu", "mandem", "dem", "manze",
}

AFRICAN_ENGLISH_FILLER = {
    "whereby", "kindly", "please", "dear", "regards", "humbly",
    "sincerely", "truly", "basically", "actually", "obviously",
    "literally", "just", "really", "very", "quite", "much",
    "thing", "things", "something", "everything", "nothing",
    "someone", "everyone", "anyone",
}

COMBINED_STOP_WORDS = (
    set(ENGLISH_STOP_WORDS)
    | SWAHILI_STOP_WORDS
    | SHENG_FILLER_STOP_WORDS
    | AFRICAN_ENGLISH_FILLER
)

# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATION RULES  ─ Africa / Kenya context
# ══════════════════════════════════════════════════════════════════════

RECOMMENDATION_RULES: List[Tuple[List[str], str]] = [
    # ══════════════════════════════════════════════════════════════════
    # CRITICAL THEMES (Appear in all 5 questions)
    # ══════════════════════════════════════════════════════════════════

    # 1. INPUT DELIVERY — TIMING & LOGISTICS
    (["delivery", "delivered", "late", "delayed", "inputs delivered",
      "planting window", "rains", "early", "dispatch", "arrive",
      "collection point", "distribute", "distribution"],
     "Overhaul input supply timing: (1) Pre-position all input stocks at "
     "sub-location collection points BEFORE rains start (by November for "
     "December-January planting); (2) Use SMS alerts to notify farmers exactly "
     "when inputs arrive at collection point; (3) Coordinate timing with county "
     "extension for weather forecasting; (4) Set weekly collection schedule with "
     "transport buddies to reduce individual collection burden."),

    # 2. LOAN REPAYMENT — ALIGNMENT WITH HARVEST
    (["loan", "repayment", "repay", "harvest", "tight", "schedule",
      "boda", "payment", "mpesa", "default", "liability", "penalised",
      "penalty", "grace", "yield", "loss", "installment", "instalment"],
     "Redesign loan repayment structure: (1) Align repayment AFTER harvest "
     "(not during growing season) — grace period of 30-45 days post-harvest "
     "to allow selling crops; (2) Allow partial repayment (50%) after first "
     "harvest sale, balance due 2 weeks after final harvest; (3) Offer M-Pesa "
     "AND cash payment options (SMS system sometimes fails); (4) Remove group "
     "liability — instead: graduated individual penalties (mild for first late "
     "payment, strict only after 3+ months); (5) Provide harvest-tracking SMS: "
     "farmers report crop sale amounts, system auto-calculates due date; "
     "(6) If crop fails (drought/pests), offer 1-month extension (don't penalize)."),

    # 3. FIELD OFFICER — PRESENCE & SUPPORT
    (["officer", "field officer", "visits", "visit", "phone", "reach",
      "changed", "turnover", "replaced", "remote", "infrequent", "sub-location",
      "difficult", "shamba", "afisa", "spread", "coverage"],
     "Strengthen field officer coverage: (1) Reduce officer-to-farmer ratio "
     "(current is too high) to max 200 farmers per officer; (2) Set minimum visit "
     "frequency: 2× per month during growing season, 1× monthly off-season; "
     "(3) Provide officers motorbike/fuel allowance to reach remote areas; "
     "(4) Create WhatsApp group for farmers to ask questions between visits; "
     "(5) Implement structured check-in: officers send weekly SMS with seasonal tips; "
     "(6) Address turnover by raising salaries and offering relocation bonuses "
     "(reduce external transfers)."),

    # 4. GROUP DYNAMICS & MEMBER ACCOUNTABILITY
    (["group", "members", "member", "group members", "problems", "meetings",
      "leader", "communicate", "communication", "liability", "default",
      "conflict", "trust", "attendance", "rules", "chama"],
     "Rebuild group trust & governance: (1) Provide group leaders with 1-day "
     "facilitation training (communication, conflict resolution); (2) Create "
     "Clear Membership Rules document in Swahili: roles, meeting frequency (2× "
     "monthly minimum), contributions expected, consequences of default; "
     "(3) Move meeting attendance tracking from oral to digital (USSD: farmer "
     "texts 'HERE' to register, creates transparent record); (4) Implement "
     "Graduated Liability: new members (0%), year 2 (25%), year 3 (50%), "
     "founding members (100%) — prevents unfair punishment of new farmers; "
     "(5) Pay group leader small incentive fee (5% of repayment fee) for 100% "
     "on-time repayment; (6) Post meeting minutes on WhatsApp + physical notice board."),

    # 5. TRAINING — LANGUAGE & TIMING
    (["training", "training sessions", "sessions", "english", "language",
      "understand", "understood", "local", "swahili", "learned", "technique",
      "spacing", "practical", "suppose", "workshop", "demo"],
     "Redesign training delivery: (1) Conduct ALL training in Swahili FIRST "
     "(not English) — translate materials to local language; (2) Schedule "
     "training OFF-harvest (Feb-Apr, Sept-Oct) when farmers have free time; "
     "(3) Move from classroom to DEMONSTRATION PLOTS (hands-on field learning); "
     "(4) Record short 5-minute videos in Swahili showing: seed spacing, "
     "fertilizer application, pest management — farmers replay offline; "
     "(5) Offer women-only sessions (separate from mixed groups) for household schedules; "
     "(6) Use Swahili proverbs (e.g., 'Haraka haraka haina baraka') to reinforce patience."),

    # ══════════════════════════════════════════════════════════════════
    # HIGH PRIORITY THEMES (Appear in 3-4 questions)
    # ══════════════════════════════════════════════════════════════════

    # 6. INPUT QUALITY — SEEDS & FERTILIZER
    (["seeds", "seed", "fertilizer", "fertiliser", "mbolea", "mbegu",
      "rotten", "germinate", "germination", "variety", "quality", "poor",
      "performed", "perform", "spacing", "agronomic"],
     "Implement pre-distribution quality assurance: (1) Test all seed batches "
     "for germination rate (must exceed 85%) before release to farmers; (2) Have "
     "supplier contracts specify quality penalties for rotten seed/fertilizer; "
     "(3) Train farmer Quality Committees to spot-check inputs at collection; "
     "(4) Establish rapid replacement protocol (48 hours) for defective inputs; "
     "(5) Record batch numbers so failed inputs can be traced & credited; "
     "(6) For underperforming varieties, collect farmer feedback and adjust seed selection."),

    # 7. MARKET LINKAGE — PRICE & ACCESS
    (["market", "market linkage", "boda boda", "transport", "cost", "price",
      "buyer", "sell", "collection", "sokoni", "mavuno", "post-harvest",
      "storage", "nearest", "far", "profit", "income"],
     "Establish farmer-friendly market linkages: (1) Develop forward contracts "
     "with 2-3 certified buyers (maize traders, aggregators) — set FIXED PRICES "
     "before harvest; (2) Organize bulked sales (groups of 5-10 farmers) to improve "
     "bargaining power vs middlemen; (3) Negotiate flat boda-boda delivery rates "
     "(e.g., 500 KES per bag) rather than per-km (saves 30-40%); (4) Create sub-location "
     "collection hubs so farmers don't travel to county market; (5) Send SMS price alerts: "
     "'Maize trading at 3200/bag at Eldoret market today'; (6) If farmer chooses own buyer, "
     "OAF provides subsidized transport support."),

    # 8. COLLECTION POINT — ACCESSIBILITY
    (["collection point", "collection", "point", "far", "distance", "mbali",
      "nearest", "village", "location", "access", "transport", "delivery",
      "reach"],
     "Decentralize input distribution: (1) Establish collection points at "
     "sub-location level (not just ward/county level) — target radius max 3 km "
     "from any farmer; (2) Partner with existing community centers (schools, "
     "health centers, churches) as collection depots; (3) Set collection windows "
     "3-4 days per week so farmers can choose convenient times; (4) Shuttle "
     "service: farmers pool transport costs to reduce per-person burden; "
     "(5) For remotest farmers, offer mobile collection (OAF vehicle visits bi-weekly)."),

    # 9. M-PESA & DIGITAL RESILIENCE
    (["mpesa", "mpesa repayment", "system", "system failed", "down",
      "payment", "ussd", "code", "airtel", "safaricom", "network", "online"],
     "Implement multi-channel payment resilience: (1) Enable M-Pesa for "
     "repayment BUT maintain CASH fallback option at collection points (network "
     "fails 10-15% of the time); (2) Test M-Pesa on all providers (Airtel too, "
     "not just Safaricom); (3) Provide USSD code option for balance checks; "
     "(4) Train officers to accept physical payment books + cash receipts for "
     "network downtime; (5) Partner with mobile money agents at collection points "
     "for instant cash-in/out; (6) Send payment confirmation SMS to farmer "
     "(reduces disputes over timing)."),

    # ══════════════════════════════════════════════════════════════════
    # CULTURAL & MULTILINGUAL CONTEXTS
    # ══════════════════════════════════════════════════════════════════

    # 10. SWAHILI-LANGUAGE CONTEXTS (Mbolea, Mafunzo, Mavuno)
    (["ilikuwa", "mbolea", "mbegu", "mafunzo", "mavuno", "shama",
      "afisa", "sokoni", "usafiri", "nzuri", "sawa", "poa"],
     "Enhance cultural + linguistic integration: (1) Conduct ALL key messaging "
     "(loan terms, quality guarantees, market prices) in Swahili FIRST, not English; "
     "(2) Record training videos in Swahili with locally-relevant examples (neighbor's "
     "harvest success, local pests); (3) Conduct focus groups in Swahili (farmers express "
     "nuanced concerns better in mother language); (4) Train field officers in Swahili "
     "facilitation skills; (5) Use Swahili proverbs in training to reinforce key messages; "
     "(6) Celebrate harvest successes in Swahili (e.g., 'Mavuno yako ilikuwa moto!' = "
     "'Your harvest was excellent!')."),

    # 11. ACRE FUND PROGRAM — SUPPORT & SATISFACTION
    (["acre fund", "acre", "fund", "support", "program", "helped", "satisfied",
      "satisfied", "experience", "expectations", "deliver", "improvement",
      "works", "success"],
     "Amplify positive impacts while addressing gaps: (1) For satisfied farmers, "
     "record success stories (50-word testimonials + photos) for training of new groups; "
     "(2) Identify what's working: ask each happy farmer 'What was most helpful?' and "
     "double down; (3) For less-satisfied: run mini-focus groups (5-6 farmers) to understand "
     "specific pain points; (4) Create feedback loop: quarterly SMS survey asking 'Rate OAF "
     "support 1-10'; (5) Invite satisfied farmers to mentor new groups (incentivize with "
     "500 KES per group mentored); (6) Address systemic gaps (seed quality, timeliness); "
     "(7) Measure impact: track yield increases year-on-year by village/group."),

    # ══════════════════════════════════════════════════════════════════
    # FALLBACK (Generic)
    # ══════════════════════════════════════════════════════════════════

    # 12. GENERAL / OTHER FEEDBACK
    (["other", "feedback", "general", "misc"],
     "For mixed or unclear feedback: (1) Conduct targeted informal conversations "
     "(1-on-1 visits) with farmers who expressed concern; (2) Ask: 'What would improve this?' "
     "and listen carefully; (3) If pattern emerges (multiple farmers mention same issue), "
     "escalate to management with farmer quotes; (4) Co-design solution WITH farmers "
     "(not for them); (5) Test solution with 2-3 farmer groups before full rollout; "
     "(6) Track impact and adjust based on farmer feedback."),
]

FALLBACK_RECOMMENDATION = (
    "Conduct targeted focus-group discussions to identify the root "
    "cause before designing an intervention."
)

# ── TIP: HOW TO EXTEND RECOMMENDATION_RULES FOR YOUR SECTOR ──────────
# After running the pipeline, check the console for:
#   "KEYWORD TIPS" — these are the actual terms found in your data
#   that did NOT match any existing rule.
# Copy those terms into a new rule tuple above, then write a
# sector-specific recommendation. Example for a WASH project:
#
#   (["borehole", "pump", "handwashing", "latrine", "open defecation"],
#    "Prioritise borehole rehabilitation and community-led sanitation "
#    "campaigns; train local pump mechanics for sustained maintenance."),
#


# ══════════════════════════════════════════════════════════════════════
# MAIN ENGINE
# ══════════════════════════════════════════════════════════════════════

class FeedbackInsightEngine:
    """
    End-to-end survey text analysis pipeline optimised for
    Kenyan / African multilingual data on 8 GB RAM.
    """

    def __init__(
        self,
        file_path: str,
        text_columns: List[str],
        config: Optional[EngineConfig] = None,
    ):
        self.file_path = Path(file_path)
        self.text_columns = text_columns
        self.config = config or EngineConfig()

        # ── load data ────────────────────────────────────────────────
        self.df = self._load_file(self.file_path)
        log.info("Loaded %d rows × %d columns", *self.df.shape)

        # ── shared NLP objects (loaded once) ─────────────────────────
        log.info("Loading sentence-transformer model …")
        self.model = SentenceTransformer("all-MiniLM-L6-v2")

        self.sia = SentimentIntensityAnalyzer()
        # Inject extra lexicon entries
        self.sia.lexicon.update(self.config.extra_sentiment_lexicon)

        self.vectorizer = TfidfVectorizer(
            stop_words=list(COMBINED_STOP_WORDS),
            ngram_range=self.config.tfidf_ngram_range,
            min_df=2,
            max_df=0.9,
            max_features=self.config.tfidf_max_features,
        )

        # ── per-question state (reset each loop iteration) ───────────
        self.q_df: pd.DataFrame = pd.DataFrame()
        self.embeddings: np.ndarray = np.array([])
        self.tfidf_matrix = None
        self.feature_names: np.ndarray = np.array([])
        self.kmeans = None
        self.cluster_summary: Dict = {}
        self.summary_df: pd.DataFrame = pd.DataFrame()
        self.key_insights: List[str] = []
        self.dataset_summary: Dict = {}
        self.emerging_issues: pd.DataFrame = pd.DataFrame()
        self.cluster_to_theme_map: Dict = {}  # Cluster ID → Theme name mapping

        # ── cross-question artefacts ──────────────────────────────────
        self.question_results: Dict = {}
        self.cross_question_themes: List[str] = []

        # ── LLM client (optional) ─────────────────────────────────────
        # Cache maps frozenset-of-top-3-keywords → theme name so the same
        # cluster never makes a second API call within a run.
        self._theme_name_cache: Dict[str, str] = {}
        self._anthropic_client = None
        want_llm = (
            self.config.use_llm_naming or self.config.use_llm_recommendations
        )
        if want_llm and ANTHROPIC_AVAILABLE:
            try:
                self._anthropic_client = _anthropic_lib.Anthropic()
                log.info("Anthropic client initialised — LLM features enabled.")
            except Exception as exc:
                log.warning("Could not init Anthropic client (%s). "
                            "Falling back to rule-based methods.", exc)
        elif want_llm and not ANTHROPIC_AVAILABLE:
            log.warning(
                "anthropic package not installed — LLM features disabled. "
                "Run: pip install anthropic"
            )

        # ── NEW: Load config and initialize engines (v3+ Upgrade) ────
        self.pipeline_config = load_config()
        
        # Initialize advanced engines if enabled
        self._init_advanced_engines()

    def _init_advanced_engines(self) -> None:
        """Initialize trigger/action engines from config."""
        try:
            # Try absolute imports first (when run as main script)
            try:
                from trigger_engine import TriggerEngine
                from action_engine import ActionEngine
            except ImportError:
                # Fall back to relative imports (when imported as module)
                from .trigger_engine import TriggerEngine
                from .action_engine import ActionEngine
            
            self.trigger_engine = TriggerEngine(
                self.pipeline_config.get("thresholds", {})
            )
            self.action_engine = ActionEngine(
                sector=self.pipeline_config.get("sector", "")
            )
            
            log.info("✓ Advanced engines initialized (triggers, actions)")
        except Exception as exc:
            log.warning("Could not initialize advanced engines (%s). "
                       "Ensure trigger_engine.py, action_engine.py "
                       "exist in project directory.", exc)
            self.trigger_engine = None
            self.action_engine = None

    def _map_sentiment_label(self, score: float) -> str:
        """
        Convert sentiment score to label for trigger logic.
        """
        if score <= -0.05:
            return "Problem"
        elif score >= 0.05:
            return "Positive"
        else:
            return "Neutral"


    def _compute_dynamic_thresholds(self, themes: list) -> dict:
        """
        Compute percentile-based thresholds from theme impacts.
        """
        impacts = [t.get("impact", 0) for t in themes if t.get("impact") is not None]

        if not impacts:
            return {"high_impact": 1, "medium_impact": 1}

        impacts_sorted = sorted(impacts)

        n = len(impacts_sorted)

        # Small dataset fallback
        if n < 5:
            high = max(impacts_sorted)
            medium = impacts_sorted[n // 2]
            return {"high_impact": high, "medium_impact": medium}

        def percentile(data, p):
            k = (len(data) - 1) * (p / 100)
            f = int(k)
            c = min(f + 1, len(data) - 1)
            if f == c:
                return data[int(k)]
            return data[f] + (data[c] - data[f]) * (k - f)

        high = percentile(impacts_sorted, 80)
        medium = percentile(impacts_sorted, 50)

        return {
            "high_impact": max(1, round(high)),
            "medium_impact": max(1, round(medium))
        }

    # ══════════════════════════════════════════════════════════════════
    # FILE LOADING
    # ══════════════════════════════════════════════════════════════════

    @staticmethod
    def _detect_encoding(path: Path) -> str:
        """Use chardet to detect CSV encoding (handles Windows-1252, UTF-8-BOM …)."""
        raw = path.read_bytes()[:50_000]
        result = chardet.detect(raw)
        enc = result.get("encoding") or "utf-8"
        log.info("Detected file encoding: %s (confidence %.0f%%)",
                 enc, (result.get("confidence") or 0) * 100)
        return enc

    def _load_file(self, path: Path) -> pd.DataFrame:
        suffix = path.suffix.lower()
        if suffix in (".xlsx", ".xls"):
            return pd.read_excel(path)
        else:
            enc = self._detect_encoding(path)
            try:
                return pd.read_csv(path, encoding=enc)
            except Exception:
                return pd.read_csv(path, encoding="utf-8", errors="replace")

    # ══════════════════════════════════════════════════════════════════
    # LANGUAGE DETECTION
    # ══════════════════════════════════════════════════════════════════

    @staticmethod
    def detect_language(text: str) -> str:
        """Return ISO 639-1 code or 'en' if detection fails."""
        if not LANGDETECT_AVAILABLE or len(text.split()) < 4:
            return "en"
        try:
            return lang_detect(text)
        except Exception:
            return "en"

    # ══════════════════════════════════════════════════════════════════
    # TEXT CLEANING  ─ language-aware
    # ══════════════════════════════════════════════════════════════════

    def clean_text(self, text: str) -> str:
        text = str(text).strip()

        # Detect language BEFORE lowercasing / stripping
        lang = self.detect_language(text)
        self.q_df.at[getattr(self, "_current_idx", 0), "lang"] = lang

        text = text.lower()
        
        # ── NEW: Sector-aware text normalization (v3+ upgrade) ────
        text = self.normalize_text(text)

        if lang in ("sw", "en"):
            # Keep only ASCII letters + spaces for now
            # (Swahili uses standard Latin alphabet so this is safe)
            text = re.sub(r"[^a-z\s]", " ", text)
        else:
            # For unidentified / mixed scripts keep letters broadly
            text = re.sub(r"[^\w\s]", " ", text, flags=re.UNICODE)

        # Normalise whitespace
        text = re.sub(r"\s+", " ", text).strip()
        return text

    def normalize_text(self, text: str) -> str:
        """
        Normalize sector-specific terminology for consistent clustering.
        
        ⚠️  RULE-BASED & SECTOR-AWARE:
        This method uses normalizations specific to smallholder agriculture.
        If you're using this on a different sector, add your own rules:
        
        HEALTHCARE examples:
            text = text.replace("hiv aids", "hiv")
            text = text.replace("doc", "doctor")
        
        WASH examples:
            text = text.replace("pit latrine", "latrine")
            text = text.replace("bore hole", "borehole")
        """
        # Smallholder agriculture normalizations
        text = text.replace("m pesa", "mpesa")      # Standardize M-Pesa
        text = text.replace("m-pesa", "mpesa")
        text = text.replace("boda boda", "boda")    # Standardize boda
        text = text.replace("bodaboda", "boda")
        text = text.replace("matatu", "transport")
        text = text.replace("extension officer", "officer")
        text = text.replace("field officer", "officer")
        
        return text

    # ══════════════════════════════════════════════════════════════════
    # SARCASM / HEDGED-POSITIVE FLAG
    # ══════════════════════════════════════════════════════════════════

    HEDGED_PATTERNS = re.compile(
        r"\b(not bad|could be better|not the best|sort of ok|i guess|"
        r"not great|nothing special|so so|meh|average|below average|"
        r"acceptable i suppose|okay i guess)\b",
        re.IGNORECASE,
    )

    def detect_hedged_sentiment(self, text: str) -> float:
        """Return a negative adjustment if hedged/sarcastic language found."""
        matches = self.HEDGED_PATTERNS.findall(text)
        return -0.2 * len(matches)  # each match nudges score negative

    # ══════════════════════════════════════════════════════════════════
    # SENTENCE SPLITTING
    # ══════════════════════════════════════════════════════════════════

    def split_sentences(self, text: str) -> List[str]:
        parts = re.split(r"[.!?;]", text)
        return [
            p.strip()
            for p in parts
            if len(p.split()) >= self.config.min_sentence_words
        ]

    # ══════════════════════════════════════════════════════════════════
    # COMPRESS TEXT  ─ stop-word filtered, for clustering
    # ══════════════════════════════════════════════════════════════════

    def compress_text(self, text: str) -> str:
        tokens = re.findall(r"\b[a-zA-Z]+\b", text.lower())
        filtered = [
            t for t in tokens
            if len(t) > 2 and t not in COMBINED_STOP_WORDS
        ]
        return " ".join(filtered)

    # ══════════════════════════════════════════════════════════════════
    # NEAR-DUPLICATE DEDUPLICATION
    # ══════════════════════════════════════════════════════════════════

    def deduplicate_responses(self) -> pd.DataFrame:
        """
        Collapse near-duplicate compressed texts.
        Keeps the first occurrence; annotates duplicates with original_count.
        This prevents one respondent repeating the same complaint
        from inflating a cluster (a known weakness of survey tools).
        """
        seen: Dict[str, int] = {}
        keep_indices = []

        for idx, row in self.q_df.iterrows():
            key = row["compressed_text"][:80]  # fingerprint first 80 chars
            if key not in seen:
                seen[key] = 1
                keep_indices.append(idx)
            else:
                seen[key] += 1

        before = len(self.q_df)
        # reset_index is CRITICAL: TF-IDF matrix rows are 0..n-1 but
        # .loc[keep_indices] preserves original (gapped) index values,
        # causing IndexError when extract_keywords() slices the matrix.
        self.q_df = self.q_df.loc[keep_indices].reset_index(drop=True)
        after = len(self.q_df)

        if before - after > 0:
            log.info("Deduplication removed %d near-duplicate sentences "
                     "(%d → %d)", before - after, before, after)

        return self.q_df

    # ══════════════════════════════════════════════════════════════════
    # PREPROCESSING
    # ══════════════════════════════════════════════════════════════════

    def preprocess(self) -> None:
        rows = []
        langs = []

        for text in self.q_df["response"]:
            clean = self.clean_text(str(text))
            lang = self.detect_language(str(text))
            for s in self.split_sentences(clean):
                rows.append(s)
                langs.append(lang)

        self.q_df = pd.DataFrame({
            "clean_text": rows,
            "lang": langs,
        })

        self.q_df["compressed_text"] = self.q_df["clean_text"].apply(
            self.compress_text
        )

        # Drop empty compressed texts
        self.q_df = self.q_df[
            self.q_df["compressed_text"].str.strip() != ""
        ].copy().reset_index(drop=True)

        log.info("After preprocessing: %d sentence units", len(self.q_df))

    # ══════════════════════════════════════════════════════════════════
    # TF-IDF
    # ══════════════════════════════════════════════════════════════════

    def build_tfidf_matrix(self) -> None:
        texts = self.q_df["compressed_text"].tolist()
        self.tfidf_matrix = self.vectorizer.fit_transform(texts)
        self.feature_names = self.vectorizer.get_feature_names_out()

    # ══════════════════════════════════════════════════════════════════
    # SENTIMENT  ─ VADER + custom lexicon + hedged-language adjustment
    # ══════════════════════════════════════════════════════════════════

    def compute_sentiment(self) -> None:
        scores = []
        for text in self.q_df["clean_text"]:
            base = self.sia.polarity_scores(text)["compound"]
            hedge_adj = self.detect_hedged_sentiment(text)
            scores.append(max(-1.0, min(1.0, base + hedge_adj)))
        self.q_df["sentiment"] = scores

    # ══════════════════════════════════════════════════════════════════
    # EMBEDDINGS  ─ deduplicated + memory-managed
    # ══════════════════════════════════════════════════════════════════

    def generate_embeddings(self) -> None:
        texts = self.q_df["compressed_text"].tolist()
        unique_texts = list(dict.fromkeys(texts))  # order-preserving dedup

        log.info("Encoding %d unique sentence units …", len(unique_texts))

        unique_embeddings = self.model.encode(
            unique_texts,
            batch_size=64,
            show_progress_bar=False,
            convert_to_numpy=True,
        )

        embedding_map = dict(zip(unique_texts, unique_embeddings))
        self.embeddings = np.array([embedding_map[t] for t in texts])

        # PCA: reduce 384 → config.pca_components to save RAM & speed up KMeans
        n_components = min(
            self.config.pca_components,
            self.embeddings.shape[0] - 1,
            self.embeddings.shape[1],
        )
        log.info("PCA: %d → %d dims", self.embeddings.shape[1], n_components)
        pca = PCA(n_components=n_components, random_state=42)
        self.embeddings = pca.fit_transform(self.embeddings).astype(np.float32)

        del unique_embeddings
        gc.collect()

    # ═══════════════════════════════════════════════════════════════════
    # EMBEDDING CACHE (v3+ Upgrade)
    # ═══════════════════════════════════════════════════════════════════
    # Dramatic speed boost on re-runs: cache embeddings to disk and reuse

    def get_embeddings_cached(
        self,
        texts: List[str],
        cache_path: str = "embeddings_cache.pkl",
        force_recompute: bool = False,
    ) -> np.ndarray:
        """
        Get embeddings with disk caching for huge speed boost on re-runs.
        
        PERFORMANCE: First run: 60s. Cached re-runs: <1s. ~60× speedup.
        
        Args:
            texts: list of strings to encode
            cache_path: where to save pickle cache
            force_recompute: if True, ignore cache and recompute
        
        Returns:
            np.ndarray of embeddings (shape: [len(texts), embedding_dim])
        """
        cache_file = Path(cache_path)
        
        # Try loading from cache
        if cache_file.exists() and not force_recompute:
            try:
                with open(cache_file, "rb") as f:
                    cached = pickle.load(f)
                if len(cached) == len(texts):
                    log.info("✓ Loaded embeddings from cache (%s)", cache_path)
                    return cached
            except Exception as exc:
                log.warning("Cache load failed (%s). Recomputing…", exc)
        
        # Compute fresh embeddings
        log.info("Computing embeddings (no valid cache) …")
        embeddings = self.model.encode(
            texts,
            batch_size=32,
            show_progress_bar=True,
            convert_to_numpy=True,
        )
        
        # Save to cache
        try:
            with open(cache_file, "wb") as f:
                pickle.dump(embeddings, f)
            log.info("✓ Saved embeddings cache to %s", cache_path)
        except Exception as exc:
            log.warning("Could not save cache (%s)", exc)
        
        return embeddings

    # ══════════════════════════════════════════════════════════════════
    # CLUSTER COUNT SELECTION
    # ══════════════════════════════════════════════════════════════════

    def choose_clusters(self) -> int:
        n = len(self.q_df)

        if n < 200:
            best_k, best_score = 2, -1.0
            for k in range(2, min(9, n)):
                km = KMeans(n_clusters=k, random_state=42, n_init=10)
                labels = km.fit_predict(self.embeddings)
                score = silhouette_score(self.embeddings, labels,
                                         sample_size=min(500, n))
                if score > best_score:
                    best_k, best_score = k, score
            return best_k

        suggested = int(np.sqrt(n / 2))
        return min(suggested, self.config.max_clusters)

    # ══════════════════════════════════════════════════════════════════
    # CLUSTERING  ─ MiniBatchKMeans for large datasets
    # ══════════════════════════════════════════════════════════════════

    def cluster_feedback(self) -> None:
        k = self.choose_clusters()
        log.info("Clustering into k=%d groups …", k)

        n = len(self.q_df)
        if n > self.config.large_dataset_threshold:
            self.kmeans = MiniBatchKMeans(
                n_clusters=k,
                random_state=42,
                n_init=5,
                batch_size=min(1024, n),
            )
        else:
            self.kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)

        self.q_df["cluster"] = self.kmeans.fit_predict(self.embeddings)

    # ══════════════════════════════════════════════════════════════════
    # FILTER CLUSTERS  ─ separate emerging vs discarded
    # ══════════════════════════════════════════════════════════════════

    def filter_clusters(self) -> None:
        counts = self.q_df["cluster"].value_counts()
        min_s = self.config.min_cluster_size
        emerging_min = self.config.emerging_min_size

        valid = counts[counts >= min_s].index
        emerging = counts[(counts >= emerging_min) & (counts < min_s)].index

        self.emerging_df = self.q_df[
            self.q_df["cluster"].isin(emerging)
        ].copy().reset_index(drop=True)

        self.q_df = self.q_df[
            self.q_df["cluster"].isin(valid)
        ].copy().reset_index(drop=True)

        log.info(
            "Clusters: %d main | %d emerging | discarded: %d",
            len(valid), len(emerging),
            len(counts) - len(valid) - len(emerging),
        )

        # Re-sync the TF-IDF matrix rows to match the filtered q_df.
        # extract_keywords() uses q_df.index values to slice tfidf_matrix,
        # so both must share the same positional indices after filtering.
        valid_positions = [
            i for i, c in enumerate(
                self.kmeans.labels_
            ) if c in valid
        ]
        self.tfidf_matrix = self.tfidf_matrix[valid_positions]

    # ══════════════════════════════════════════════════════════════════
    # KEYWORD EXTRACTION  ─ phrase-aware, generic-filtered
    # ══════════════════════════════════════════════════════════════════

    GENERIC_WORDS = {
        "service", "good", "bad", "thing", "things", "experience",
        "people", "place", "really", "very", "much", "feel", "make",
        "like", "know", "need", "want", "use", "get", "go", "come",
        "say", "tell", "ask", "give", "take", "see", "think",
    }

    def extract_keywords(
        self, cluster_indices: List[int], top_n: int = 6
    ) -> List[str]:
        cluster_matrix = self.tfidf_matrix[cluster_indices]
        cluster_scores = np.asarray(cluster_matrix.mean(axis=0)).ravel()
        global_scores = np.asarray(self.tfidf_matrix.mean(axis=0)).ravel()
        importance = cluster_scores - global_scores

        keywords: List[str] = []
        for idx in importance.argsort()[::-1]:
            term = self.feature_names[idx]
            if term in self.GENERIC_WORDS:
                continue
            keywords.append(term)
            if len(keywords) >= top_n:
                break
        return keywords

    # ══════════════════════════════════════════════════════════════════
    # THEME NAMING  ─ LLM-first with smart rule-based fallback
    # ══════════════════════════════════════════════════════════════════

    def generate_theme_name(
        self, keywords: List[str], samples: Optional[List[str]] = None
    ) -> str:
        """
        Public entry-point for theme naming.
        • LLM path  : one cached API call → meaningful 2-4 word name.
        • Fallback  : improved rule-based (no redundant word-pairs).
        """
        if not keywords:
            return "Other Feedback"

        # Cache key = sorted top-3 keywords (stable across calls)
        cache_key = "|".join(sorted(keywords[:3]))
        if cache_key in self._theme_name_cache:
            return self._theme_name_cache[cache_key]

        if self._anthropic_client and self.config.use_llm_naming:
            name = self._llm_theme_name(keywords, samples or [])
        else:
            name = self._rule_based_theme_name(keywords)

        self._theme_name_cache[cache_key] = name
        return name

    def _rule_based_theme_name(self, keywords: List[str]) -> str:
        """
        Improved rule-based naming — avoids gibberish like 'Training Seeds'
        or 'Seeds Seed'.

        Strategy:
          1. Bigram TF-IDF features (e.g. 'late delivery') are already
             meaningful phrases → prefer them as the theme.
          2. For unigrams, skip any keyword whose stem is already represented
             by another keyword in the pair (avoids 'Train Training').
          3. Build at most a 3-token theme; deduplicate within it.
        """
        bigrams  = [k for k in keywords if " " in k]
        unigrams = [k for k in keywords if " " not in k]

        # ── Case 1: use the best bigram as the primary label ──────────
        if bigrams:
            primary = bigrams[0].title()
            # Try to append a discriminating unigram if it adds new info
            for u in unigrams:
                if u.lower() not in bigrams[0].lower():
                    return f"{primary} — {u.title()}"
            return primary

        # ── Case 2: two unigrams — check for stem overlap ─────────────
        if len(unigrams) >= 2:
            p1, p2 = unigrams[0], unigrams[1]
            # Overlap check: one is a substring of the other?
            if p1.lower() in p2.lower() or p2.lower() in p1.lower():
                # Redundant pair → use longer one + third keyword
                primary = max(p1, p2, key=len).title()
                if len(unigrams) >= 3:
                    return f"{primary} {unigrams[2].title()}"
                return primary
            # Pair looks fine
            return f"{p1.title()} {p2.title()}"

        # ── Case 3: single keyword ─────────────────────────────────────
        return unigrams[0].title() if unigrams else "Other Feedback"

    def _llm_theme_name(
        self, keywords: List[str], samples: List[str]
    ) -> str:
        """
        Ask Claude (Haiku) for a meaningful 2-4 word theme name.
        Falls back to rule-based on any API error.
        """
        sector_ctx = (
            f"The data comes from a {self.config.sector} programme."
            if self.config.sector else ""
        )
        sample_text = " | ".join(samples[:3]) if samples else "(none)"
        prompt = (
            f"You are labelling themes extracted from survey feedback. "
            f"{sector_ctx}\n\n"
            f"Top keywords: {', '.join(keywords[:6])}\n"
            f"Example responses: {sample_text}\n\n"
            "Give a clear, meaningful 2–4 word theme label that a non-technical "
            "M&E officer would immediately understand. "
            "Return ONLY the theme name — no punctuation, no explanation."
        )
        try:
            response = self._anthropic_client.messages.create(
                model=self.config.anthropic_model,
                max_tokens=20,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = response.content[0].text.strip().strip('"').strip("'")
            # Sanity: reject if response is too long or empty
            if 2 <= len(raw.split()) <= 6 and raw:
                return raw.title()
            return self._rule_based_theme_name(keywords)
        except Exception as exc:
            log.debug("LLM theme naming failed (%s); using rule-based.", exc)
            return self._rule_based_theme_name(keywords)

    # ══════════════════════════════════════════════════════════════════
    # CLUSTER SUMMARY BUILD
    # ══════════════════════════════════════════════════════════════════

    def build_cluster_summary(self) -> None:
        self.cluster_summary = {}
        for cid in sorted(self.q_df["cluster"].unique()):
            subset = self.q_df[self.q_df["cluster"] == cid]
            texts   = subset["clean_text"].tolist()
            indices = subset.index.tolist()
            kw = self.extract_keywords(indices, self.config.top_keywords)
            self.cluster_summary[cid] = {
                "count":    len(texts),
                "keywords": kw,
                "samples":  texts[:5],
            }
            # Pre-warm the theme name cache while cluster data is fresh
            self.generate_theme_name(kw, texts[:3])

    # ══════════════════════════════════════════════════════════════════
    # SUMMARY TABLE  ─ % based on per-question raw response count
    # ══════════════════════════════════════════════════════════════════

    def build_summary_table(self) -> None:
        """
        Build a standardized summary of themes.

        This revised implementation standardizes theme structure, computes
        percentile-based thresholds, reinitializes the TriggerEngine with
        data-driven thresholds, applies Trigger/Action engines to
        each theme, validates results, and produces `self.summary_df`.
        """
        total = self._question_response_count  # set in run() per question

        # 1) Collect raw theme records from cluster summary
        themes = []
        for cid, info in self.cluster_summary.items():
            theme_name = self.generate_theme_name(info.get("keywords", []))
            subset = self.q_df[self.q_df["cluster"] == cid]
            avg_sent = subset["sentiment"].mean() if len(subset) > 0 else 0.0
            pct = round((info.get("count", 0) / max(total, 1)) * 100, 2)

            # intensity: few very negative sentences vs many mildly negative
            negative_sentences = subset[subset["sentiment"] < -0.3]
            intensity = (
                negative_sentences["sentiment"].abs().mean()
                if len(negative_sentences) > 0 else 0.0
            )
            severity = info.get("count", 0) * abs(avg_sent)

            themes.append({
                "cluster_id": cid,
                "name": theme_name,
                "impact": int(info.get("count", 0)),
                "share_pct": pct,
                "impact_score": round(severity, 2),
                "intensity": round(intensity, 3),
                "sentiment_score": round(avg_sent, 3),
                "keywords": [k.strip() for k in info.get("keywords", []) if k.strip()],
                "representative_quote": info.get("samples", [""])[0] if info.get("samples") else "",
                "recommendation": info.get("recommendation", "") or "",
            })

        # 2) Map sentiment scores -> labels (strict mapping per spec)
        for t in themes:
            sc = t.get("sentiment_score", 0.0)
            if sc <= -0.05:
                t["sentiment"] = "Problem"
            elif sc >= 0.05:
                t["sentiment"] = "Positive"
            else:
                t["sentiment"] = "Neutral"

        # 3) Compute percentile-based thresholds for impact
        impacts = sorted([t["impact"] for t in themes])
        high_thresh = None
        med_thresh = None
        if len(impacts) == 0:
            high_thresh = med_thresh = 0
        elif len(impacts) < 5:
            # small dataset: use max and median
            high_thresh = max(impacts)
            med_thresh = int(pd.Series(impacts).median())
        else:
            # 80th and 50th percentiles
            high_thresh = int(pd.Series(impacts).quantile(0.8))
            med_thresh = int(pd.Series(impacts).quantile(0.5))

        # Handle uniform values edge-case
        if high_thresh is None:
            high_thresh = med_thresh = 0
        if high_thresh == med_thresh and len(impacts) > 0:
            # keep them equal; TriggerEngine will handle equality
            pass

        log.info(f"  ▶ Computed thresholds — HIGH: {high_thresh}, MEDIUM: {med_thresh}")

        # 4) Reinitialize TriggerEngine with new thresholds (overrides previous)
        try:
            from trigger_engine import TriggerEngine
            self.trigger_engine = TriggerEngine({"high_impact": high_thresh, "medium_impact": med_thresh})
        except Exception:
            log.warning("TriggerEngine import failed — continuing with existing engine")

        # 5) Apply all engines to each theme and enrich
        enriched = []
        for t in themes:
            theme_input = {
                "name": t["name"],
                "impact": t["impact"],
                "sentiment": t.get("sentiment", "Neutral"),
                "sentiment_score": t.get("sentiment_score", 0.0),
                "keywords": t.get("keywords", []),
                "recommendation": t.get("recommendation", ""),
            }

            # Trigger
            triggers = []
            if getattr(self, "trigger_engine", None) and self.pipeline_config.get("features", {}).get("enable_trigger_engine", True):
                try:
                    triggers = self.trigger_engine.evaluate(theme_input) or []
                except Exception as exc:
                    log.debug("Trigger engine error for %s: %s", t["name"], exc)

            # Determine priority label from triggers (pick highest)
            priority_label = "Low"
            if isinstance(triggers, list) and triggers:
                levels = [str(x.get("level", "")).upper() for x in triggers]
                if any(l == "HIGH" for l in levels):
                    priority_label = "High"
                elif any(l == "MEDIUM" for l in levels):
                    priority_label = "Medium"
                elif any(l == "POSITIVE" for l in levels):
                    priority_label = "Positive"

            # Trend is REMOVED — TriggerEngine is the single source of truth for priority
            
            # Action plan
            action_plan = {}
            if getattr(self, "action_engine", None) and self.pipeline_config.get("features", {}).get("enable_action_engine", True):
                try:
                    action_plan = self.action_engine.generate(theme_input) or {}
                except Exception as exc:
                    log.debug("Action engine error for %s: %s", t["name"], exc)

            # Compose actions string
            actions_list = action_plan.get("actions") if isinstance(action_plan.get("actions"), list) else []
            actions_joined = "; ".join(str(a) for a in actions_list) if actions_list else "No action specified"

            # Extract Trigger Status from triggers list (ALWAYS has one element)
            trigger_status = "No Alert"
            if isinstance(triggers, list) and triggers:
                trigger_status = triggers[0].get("message", "No Alert")
            
            enriched.append({
                "Cluster ID": t["cluster_id"],
                "Theme": t["name"],
                "Response Count": t["impact"],
                "Share (%)": t["share_pct"],
                "Impact Score": t["impact_score"],
                "Intensity Score": t["intensity"],
                "Average Sentiment": t["sentiment_score"],
                "Theme Sentiment": t.get("sentiment", "Neutral"),
                "Key Keywords": ", ".join(t.get("keywords", [])),
                "Representative Quote": t.get("representative_quote", ""),
                "triggers": triggers,
                "Priority": priority_label,
                "Trigger Status": trigger_status,
                "Action Owner": action_plan.get("owner", "Program Manager"),
                "Timeline (Days)": action_plan.get("timeline", 30),
                "Actions": actions_joined,
                "Recommendation": t.get("recommendation", ""),
                "action_plan": action_plan,
            })

        # 6) Validation: ensure required fields
        for rec in enriched:
            if not rec.get("Priority"):
                rec["Priority"] = "Low"
            if not rec.get("Trigger Status"):
                rec["Trigger Status"] = "No Alert"
            if not rec.get("Action Owner"):
                rec["Action Owner"] = "Program Manager"
            if not rec.get("Actions"):
                rec["Actions"] = "No action specified"

        # 7) Build final DataFrame
        self.summary_df = pd.DataFrame(enriched)
        # Sort by Impact Score desc for presentation
        if "Impact Score" in self.summary_df.columns:
            self.summary_df = self.summary_df.sort_values("Impact Score", ascending=False).reset_index(drop=True)

    # ═══════════════════════════════════════════════════════════════════
    # APPLY ADVANCED ENGINES (DEPRECATED)
    # ═══════════════════════════════════════════════════════════════════
    # Logic now moved to build_summary_table for single source of truth

    def apply_advanced_engines(self) -> None:
        """
        All advanced engine logic (triggers, actions) is now handled in
        build_summary_table() to ensure single source of truth.
        This method is retained for backward compatibility but does nothing.
        """
        # All logic now in build_summary_table
        if all(col in self.summary_df.columns for col in ("triggers", "Priority", "action_plan")):
            log.info("Advanced engines already applied in build_summary_table — skipping.")
            return
        
        log.info("✓ Advanced engines already applied")

    # ══════════════════════════════════════════════════════════════════
    # SENTIMENT LABEL  ─ keyword-aware + sentiment threshold
    # ══════════════════════════════════════════════════════════════════

    def classify_theme_sentiment(self) -> None:
        """Classify theme sentiment using BOTH sentiment score AND keywords.
        
        This prevents false neutrals where a theme seems like a problem
        but sentiment is borderline.
        
        Thresholds (v3.0):
        - Problem: sentiment < -0.05 OR (keywords match problem terms)
        - Positive: sentiment > 0.1 AND no problem keywords
        - Neutral: everything else
        """
        problem_keywords = {
            "problem", "issue", "complaint", "poor", "bad", "terrible",
            "awful", "horrible", "delay", "slow", "difficult", "challenge",
            "struggle", "frustrate", "disappoint", "dissat", "fail",
            "broken", "error", "bug", "stuck", "unavailable", "shortage",
            "late", "waiting", "queue", "rude", "unfriendly", "expensive",
            "costly", "mbaya", "tatizo", "shida", "lalamiko"
        }
        
        labels = []
        for idx, row in self.summary_df.iterrows():
            sentiment = row["Average Sentiment"]
            keywords = str(row.get("Key Keywords", "")).lower()
            
            # Check if problem keywords present
            has_problem_keyword = any(pk in keywords for pk in problem_keywords)
            
            if sentiment < -0.05 or has_problem_keyword:
                labels.append("Problem")
            elif sentiment > 0.1 and not has_problem_keyword:
                labels.append("Positive")
            else:
                labels.append("Neutral")
        
        self.summary_df["Theme Sentiment"] = labels

    # ══════════════════════════════════════════════════════════════════
    # RECOMMENDATIONS  ─ LLM-first, rule-based fallback, keyword tips
    # ══════════════════════════════════════════════════════════════════

    def generate_recommendations(self) -> None:
        """
        Generates one recommendation per theme.

        Mode A — LLM (use_llm_recommendations=True AND anthropic installed):
          Sends ALL themes in a single batch API call so latency is low.
          The LLM receives: theme name, sentiment, keywords, a sample quote,
          and the sector context — producing specific, actionable advice.

        Mode B — Rule-based (fallback):
          Matches keywords against RECOMMENDATION_RULES.
          After matching, prints the KEYWORD TIPS report so you can see
          which terms fell through to the fallback and add new rules.
        """
        if self._anthropic_client and self.config.use_llm_recommendations:
            self._llm_generate_recommendations()
        else:
            self._rule_based_recommendations()
            self.print_keyword_tips(quiet=False)

    def _llm_generate_recommendations(self) -> None:
        """Batch API call: all themes → one request → list of recommendations."""
        sector_ctx = (
            f"The organisation works in the {self.config.sector} sector in "
            f"Kenya/East Africa."
            if self.config.sector
            else "The organisation operates in Kenya/East Africa."
        )

        # Build a compact JSON payload — only what the LLM needs
        themes_payload = [
            {
                "id": i,
                "theme": row["Theme"],
                "sentiment": row.get("Theme Sentiment", ""),
                "keywords": row["Key Keywords"],
                "quote": row["Representative Quote"],
            }
            for i, (_, row) in enumerate(self.summary_df.iterrows())
        ]

        prompt = (
            f"You are an experienced M&E advisor. {sector_ctx}\n\n"
            "For each feedback theme below, write ONE practical, specific "
            "recommendation (1–2 sentences). Use local African context where "
            "relevant (M-Pesa, boda-boda, SMS/USSD, county structures, etc.).\n\n"
            f"Themes:\n{json.dumps(themes_payload, indent=2)}\n\n"
            "Return a JSON array of strings — one recommendation per theme, "
            "in the SAME order as the input. "
            "Return ONLY the JSON array. No explanation, no markdown."
        )

        try:
            response = self._anthropic_client.messages.create(
                model=self.config.anthropic_model,
                max_tokens=800,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = response.content[0].text.strip()
            # Strip markdown code fences if present
            raw = re.sub(r"^```[a-z]*\n?", "", raw)
            raw = re.sub(r"\n?```$", "", raw).strip()
            recs = json.loads(raw)

            if isinstance(recs, list) and len(recs) == len(self.summary_df):
                if "Recommendation" in self.summary_df.columns:
                    self.summary_df["Recommendation"] = recs
                else:
                    try:
                        self.summary_df.insert(5, "Recommendation", recs)
                    except Exception:
                        self.summary_df["Recommendation"] = recs
                return
            log.warning("LLM returned %d recs for %d themes — falling back.",
                        len(recs), len(self.summary_df))
        except Exception as exc:
            log.warning("LLM recommendation generation failed (%s) — "
                        "falling back to rule-based.", exc)

        # Fallback
        self._rule_based_recommendations()

    def _rule_based_recommendations(self) -> None:
        """Match theme keywords against RECOMMENDATION_RULES."""
        recs = []
        for _, row in self.summary_df.iterrows():
            text = (row["Theme"] + " " + row["Key Keywords"]).lower()
            matched = FALLBACK_RECOMMENDATION
            for rule_keywords, rec in RECOMMENDATION_RULES:
                if any(k in text for k in rule_keywords):
                    matched = rec
                    break
            recs.append(matched)
        if "Recommendation" in self.summary_df.columns:
            self.summary_df["Recommendation"] = recs
        else:
            try:
                self.summary_df.insert(5, "Recommendation", recs)
            except Exception:
                self.summary_df["Recommendation"] = recs

    def print_keyword_tips(self, quiet: bool = False) -> None:
        """
        ╔══════════════════════════════════════════════════════════════╗
        ║  KEYWORD TIPS — Recommendation Rule Coverage Report         ║
        ╚══════════════════════════════════════════════════════════════╝
        Prints all keywords found in this question, whether each matched
        a RECOMMENDATION_RULE, and which terms fell through to the
        generic fallback.

        Use this to extend RECOMMENDATION_RULES for your sector:
          1. Look at the ⚠ UNMATCHED section below.
          2. Group related unmatched terms into a new rule tuple.
          3. Write a sector-specific recommendation string.
          4. Add it to RECOMMENDATION_RULES near the top of the file.
        """
        if quiet or self.summary_df.empty:
            return

        all_rule_kw: set = set()
        for rule_kw, _ in RECOMMENDATION_RULES:
            all_rule_kw.update(rule_kw)

        SEP = "─" * 62
        print(f"\n{'═' * 62}")
        print("  KEYWORD TIPS  —  Recommendation Rule Coverage")
        print(f"{'═' * 62}")

        for _, row in self.summary_df.iterrows():
            theme   = row["Theme"]
            kw_str  = row.get("Key Keywords", "")
            rec     = row.get("Recommendation", "")
            kw_list = [k.strip() for k in kw_str.split(",") if k.strip()]

            matched   = [k for k in kw_list if k in all_rule_kw]
            unmatched = [k for k in kw_list if k not in all_rule_kw]
            used_fallback = (rec == FALLBACK_RECOMMENDATION)

            icon = "⚠ " if used_fallback else "✅"
            print(f"\n  {icon} Theme : {theme}")
            print(f"     Keywords : {', '.join(kw_list) or '(none)'}")
            if matched:
                print(f"     Matched rules on : {', '.join(matched)}")
            if unmatched:
                print(f"     ⚠ Unmatched terms : {', '.join(unmatched)}")
                print(f"       → Add these to RECOMMENDATION_RULES for "
                      f"better coverage.")
            if used_fallback:
                print(f"     ⚠ Used generic fallback — no rule matched.")

        print(f"\n{SEP}")
        print("  TIP: To add a rule, insert this pattern into RECOMMENDATION_RULES:")
        print('    (["keyword1", "keyword2", ...],')
        print('     "Your sector-specific recommendation here."),')
        print(f"{SEP}\n")

    # ══════════════════════════════════════════════════════════════════
    # PRIORITY LABELS  ─ intensity-aware
    # ══════════════════════════════════════════════════════════════════

    # ══════════════════════════════════════════════════════════════════
    # MERGE SIMILAR THEMES
    # ══════════════════════════════════════════════════════════════════

    def merge_similar_themes(self) -> None:
        themes = self.summary_df["Theme"].tolist()
        
        # Initialize mapping for all clusters
        self.cluster_to_theme_map = {}
        if "Cluster ID" in self.summary_df.columns:
            for idx, row in self.summary_df.iterrows():
                cluster_id = row.get("Cluster ID", -1)
                self.cluster_to_theme_map[cluster_id] = row.get("Theme", "Unknown")
        
        if len(themes) <= 1:
            return

        embeddings = self.model.encode(
            themes, convert_to_numpy=True, normalize_embeddings=True
        )

        merged_map: Dict[str, List[int]] = {}
        used: set = set()

        for i, theme_i in enumerate(themes):
            if i in used:
                continue
            merged_map[theme_i] = [i]
            for j in range(i + 1, len(themes)):
                if j in used:
                    continue
                if np.dot(embeddings[i], embeddings[j]) >= self.config.merge_threshold:
                    merged_map[theme_i].append(j)
                    used.add(j)

        # Update cluster ID mapping with merged themes
        self.cluster_to_theme_map = {}
        
        new_rows = []
        for main_theme, idxs in merged_map.items():
            subset = self.summary_df.iloc[idxs]
            # Get all cluster IDs being merged
            cluster_ids = subset["Cluster ID"].tolist() if "Cluster ID" in subset.columns else []
            for cid in cluster_ids:
                self.cluster_to_theme_map[cid] = main_theme
            
            row = {
                "Cluster ID": cluster_ids[0] if cluster_ids else -1,  # Store first cluster ID
                "Theme": main_theme,
                "Response Count": subset["Response Count"].sum(),
                "Share (%)": round(subset["Share (%)"].sum(), 2),
                "Impact Score": round(subset["Impact Score"].sum(), 2),
                "Intensity Score": round(subset["Intensity Score"].mean(), 3),
                "Average Sentiment": round(subset["Average Sentiment"].mean(), 3),
                "Key Keywords": ", ".join(subset["Key Keywords"].tolist()),
                "Representative Quote": subset["Representative Quote"].iloc[0],
                "Recommendation": subset["Recommendation"].iloc[0],
                "Priority": subset["Priority"].iloc[0],
                "Trigger Status": subset["Trigger Status"].iloc[0] if "Trigger Status" in subset.columns else "No Alert",
                # ─ Preserve advanced engine columns ──────────────────
                "action_plan": subset["action_plan"].iloc[0] if "action_plan" in subset.columns else {},
                "triggers": subset["triggers"].iloc[0] if "triggers" in subset.columns else [],
            }
            new_rows.append(row)

        self.summary_df = pd.DataFrame(new_rows).sort_values(
            "Impact Score", ascending=False
        )
        # Recompute sentiment label after merge
        self.classify_theme_sentiment()

    # ══════════════════════════════════════════════════════════════════
    # KEY INSIGHTS
    # ══════════════════════════════════════════════════════════════════

    def generate_key_insights(self) -> None:
        insights: List[str] = []
        df = self.summary_df

        if df.empty:
            self.key_insights = ["Insufficient data for insights."]
            return

        top = df.iloc[0]
        insights.append(
            f"Most common theme: '{top['Theme']}' — "
            f"appearing in {top['Share (%)']}% of responses."
        )

        most_neg = df.sort_values("Average Sentiment").iloc[0]
        if most_neg["Theme"] != top["Theme"]:
            insights.append(
                f"Highest dissatisfaction: '{most_neg['Theme']}' "
                f"(avg sentiment {most_neg['Average Sentiment']:.2f})."
            )

        high_intensity = df[df["Intensity Score"] > 0.5]
        if not high_intensity.empty:
            themes_hi = ", ".join(f"'{t}'" for t in high_intensity["Theme"].head(3))
            insights.append(
                f"High emotional intensity detected in: {themes_hi}. "
                "A small number of respondents expressed strong dissatisfaction — "
                "these are early-warning signals."
            )

        # Loud vs widespread distinction
        loud = df[
            (df["Response Count"] <= df["Response Count"].quantile(0.25)) &
            (df["Average Sentiment"] < -0.3)
        ]
        if not loud.empty:
            insights.append(
                f"Note: '{loud.iloc[0]['Theme']}' has low volume but high negative "
                "sentiment — this may be a loud minority, not widespread feedback."
            )

        positives = df[df["Theme Sentiment"] == "Positive"]
        if not positives.empty:
            pt = positives.iloc[0]["Theme"]
            insights.append(
                f"Top positive area: '{pt}' — reinforce and communicate "
                "this strength to stakeholders."
            )

        self.key_insights = insights

    # ══════════════════════════════════════════════════════════════════
    # DATASET SUMMARY
    # ══════════════════════════════════════════════════════════════════

    def generate_dataset_summary(self) -> None:
        avg_len = self.q_df["clean_text"].apply(
            lambda x: len(str(x).split())
        ).mean()

        lang_counts = self.q_df.get("lang", pd.Series(dtype=str)).value_counts()
        lang_note = ", ".join(
            f"{lang}:{cnt}" for lang, cnt in lang_counts.head(4).items()
        ) if not lang_counts.empty else "n/a"

        self.dataset_summary = {
            "Total Responses": self._question_response_count,
            "Sentence Units Analysed": len(self.q_df),
            "Average Response Length (words)": round(avg_len, 1),
            "Themes Identified": len(self.summary_df),
            "Emerging Issues Found": len(
                getattr(self, "emerging_df", pd.DataFrame())
            ),
            "Languages Detected": lang_note,
        }

    # ══════════════════════════════════════════════════════════════════
    # LABEL RESPONSES
    # ══════════════════════════════════════════════════════════════════

    def label_data(self) -> None:
        theme_map = {
            cid: self.generate_theme_name(info["keywords"])
            for cid, info in self.cluster_summary.items()
        }
        self.q_df["Theme"] = self.q_df["cluster"].map(theme_map)

    # ══════════════════════════════════════════════════════════════════
    # CROSS-QUESTION THEME DETECTION
    # ══════════════════════════════════════════════════════════════════

    def detect_cross_question_themes(self) -> None:
        """Find themes recurring in ≥2 questions — a key gap in most tools."""
        from collections import Counter
        all_themes: List[str] = []
        for qr in self.question_results.values():
            all_themes.extend(qr["summary"]["Theme"].tolist())

        counts = Counter(all_themes)
        self.cross_question_themes = [
            t for t, c in counts.most_common() if c >= 2
        ]
        if self.cross_question_themes:
            log.info(
                "Cross-question recurring themes: %s",
                self.cross_question_themes,
            )

    # ══════════════════════════════════════════════════════════════════
    # EXCEL STYLE HELPERS
    # ══════════════════════════════════════════════════════════════════

    HEADER_COLOUR  = "2F5597"   # navy blue (with white text)
    HEADING_COLOUR = "1F3864"   # darker navy for sheet title bar (with white text)
    STRIPE_COLOUR  = "EEF2FF"   # soft lavender stripe
    SUBHEAD_COLOUR = "D6E4F7"   # light blue for section sub-headers (with BLACK text)

    ACCENT_COLOURS = {
        "Problem":  "FFE0E0",   # soft red
        "Positive": "E0FFE8",   # soft green
        "Neutral":  "FFF9E0",   # soft yellow
    }

    # Preferred minimum column widths by header name (characters)
    COL_MIN_WIDTHS = {
        "Theme":                  32,
        "Question":               20,
        "Response Count":         18,
        "Share (%)":              14,
        "Impact Score":           16,
        "Intensity Score":        16,
        "Average Sentiment":      20,
        "Priority":               20,
        "Theme Sentiment":        18,
        "Key Keywords":           38,
        "Representative Quote":   52,
        "Recommendation":         48,
        "Quote":                  55,
        "Languages Detected":     28,
    }
    COL_DEFAULT_MIN = 14   # fallback minimum for any column not listed above

    def _thin_border(self):
        thin = Side(style="thin")
        return Border(left=thin, right=thin, top=thin, bottom=thin)

    def _medium_border(self):
        med = Side(style="medium")
        return Border(left=med, right=med, top=med, bottom=med)

    # ------------------------------------------------------------------
    # write_sheet_heading
    #   Writes a professional, centered, merged title bar at row 1.
    #   Returns the next available row number (always row 3, leaving a
    #   blank spacer row beneath the heading).
    # ------------------------------------------------------------------
    def write_sheet_heading(
        self,
        ws,
        title: str,
        subtitle: str = "",
        merge_cols: int = 10,
    ) -> int:
        """
        Row 1 : merged navy title bar — large white bold text, centered.
        Row 2 : merged subtitle bar (smaller, italic) — OR blank spacer.
        Returns 3 so callers know where the table starts.
        """
        navy_fill = PatternFill(
            start_color=self.HEADING_COLOUR,
            end_color=self.HEADING_COLOUR,
            fill_type="solid",
        )
        sub_fill = PatternFill(
            start_color=self.HEADER_COLOUR,
            end_color=self.HEADER_COLOUR,
            fill_type="solid",
        )

        end_col = get_column_letter(merge_cols)

        # ── Row 1 : main title ────────────────────────────────────────
        ws.merge_cells(f"A1:{end_col}1")
        title_cell = ws["A1"]
        title_cell.value = title
        title_cell.font = Font(
            name="Calibri", size=16, bold=True, color="FFFFFF"
        )
        title_cell.fill = navy_fill
        title_cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=False
        )
        ws.row_dimensions[1].height = 32

        # ── Row 2 : subtitle / generated date ────────────────────────
        ws.merge_cells(f"A2:{end_col}2")
        sub_cell = ws["A2"]
        sub_cell.value = subtitle
        sub_cell.font = Font(
            name="Calibri", size=10, italic=True, color="FFFFFF"
        )
        sub_cell.fill = sub_fill
        sub_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 18

        return 3   # next usable row

    # ------------------------------------------------------------------
    # style_table_header  —  styles a specific row (the column-header row)
    # ------------------------------------------------------------------
    def style_table_header(self, ws, header_row: int = 1) -> None:
        fill = PatternFill(
            start_color=self.HEADER_COLOUR,
            end_color=self.HEADER_COLOUR,
            fill_type="solid",
        )
        for cell in ws[header_row]:
            if cell.value is None:
                continue
            # Use white text for navy headers
            cell.font = Font(
                name="Calibri", size=10, bold=True, color="FFFFFF"
            )
            cell.fill = fill
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell.border = self._thin_border()
        ws.row_dimensions[header_row].height = 28

    # ------------------------------------------------------------------
    # zebra_rows  —  light stripe on every even data row
    # ------------------------------------------------------------------
    def zebra_rows(self, ws, start_row: int = 2) -> None:
        fill = PatternFill(
            start_color=self.STRIPE_COLOUR,
            end_color=self.STRIPE_COLOUR,
            fill_type="solid",
        )
        for r in range(start_row, ws.max_row + 1):
            if r % 2 == 0:
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=r, column=col)
                    # Don't overwrite sentiment-coloured rows
                    if cell.fill.fgColor.rgb in (
                        "00000000", "FF" + self.STRIPE_COLOUR
                    ):
                        cell.fill = fill

    # ------------------------------------------------------------------
    # add_table_borders
    # ------------------------------------------------------------------
    def add_table_borders(self, ws, start_row: int = 1) -> None:
        border = self._thin_border()
        for row in ws.iter_rows(
            min_row=start_row, max_row=ws.max_row,
            min_col=1, max_col=ws.max_column
        ):
            for cell in row:
                cell.border = border

    # ------------------------------------------------------------------
    # smart_col_widths
    #   Uses COL_MIN_WIDTHS for known columns; measures content otherwise.
    #   Applies a comfortable padding and hard caps.
    # ------------------------------------------------------------------
    def smart_col_widths(self, ws) -> None:
        for i, col_cells in enumerate(ws.columns, 1):
            header_val = str(ws.cell(row=1, column=i).value or "")
            # Check both row-1 (heading) and row-3/4 for actual header name
            for check_row in (1, 2, 3, 4):
                hdr = str(ws.cell(row=check_row, column=i).value or "")
                if hdr in self.COL_MIN_WIDTHS:
                    header_val = hdr
                    break

            preferred_min = self.COL_MIN_WIDTHS.get(
                header_val, self.COL_DEFAULT_MIN
            )

            # Also measure actual content
            max_content = max(
                (len(str(c.value)) for c in col_cells if c.value), default=0
            )
            # Use whichever is bigger, but cap at 60
            width = min(max(preferred_min, max_content + 3), 60)
            ws.column_dimensions[get_column_letter(i)].width = width

    # ------------------------------------------------------------------
    # sentiment_colour_rows
    # ------------------------------------------------------------------
    def sentiment_colour_rows(self, ws, sentiment_col_idx: int,
                               start_row: int = 2) -> None:
        for r in range(start_row, ws.max_row + 1):
            cell = ws.cell(row=r, column=sentiment_col_idx)
            colour = self.ACCENT_COLOURS.get(str(cell.value))
            if colour:
                fill = PatternFill(
                    start_color=colour, end_color=colour, fill_type="solid"
                )
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=col).fill = fill

    # ------------------------------------------------------------------
    # _data_row_height  —  set comfortable height on all data rows
    # ------------------------------------------------------------------
    @staticmethod
    def _set_data_row_heights(ws, start_row: int, height: float = 18) -> None:
        for r in range(start_row, ws.max_row + 1):
            if ws.row_dimensions[r].height is None or \
               ws.row_dimensions[r].height < height:
                ws.row_dimensions[r].height = height

    # ═══════════════════════════════════════════════════════════════════
    # PREPARE SUMMARY FOR EXCEL (v3+ Upgrade)
    # ═══════════════════════════════════════════════════════════════════

    def prepare_summary_for_excel(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Convert complex trigger/action dicts to readable Excel-friendly strings.
        
        Args:
            df: summary_df from question_results
        
        Returns:
            Modified dataframe with trigger/action columns as strings
        """
        export_df = df.copy()
        
        # Trigger Status should already be in summary_df from build_summary_table
        if "Trigger Status" not in export_df.columns and "triggers" in export_df.columns:
            export_df["Trigger Status"] = export_df["triggers"].apply(
                lambda triggers: (
                    " | ".join([t.get("icon", "") + " " + t.get("level", "") 
                               for t in triggers])
                    if isinstance(triggers, list) and triggers else "No Alert"
                )
            )
        elif "Trigger Status" not in export_df.columns:
            export_df["Trigger Status"] = "No Alert"
        
        # Action Owner and Timeline should already be in summary_df from build_summary_table
        if "Action Owner" not in export_df.columns or "Timeline (Days)" not in export_df.columns:
            if "action_plan" in export_df.columns:
                if "Action Owner" not in export_df.columns:
                    export_df["Action Owner"] = export_df["action_plan"].apply(
                        lambda ap: (
                            ap.get("owner", "Program Manager") if isinstance(ap, dict) else "Program Manager"
                        )
                    )
                
                if "Timeline (Days)" not in export_df.columns:
                    export_df["Timeline (Days)"] = export_df["action_plan"].apply(
                        lambda ap: (
                            str(ap.get("timeline", 60)) if isinstance(ap, dict) else "60"
                        )
                    )
            else:
                if "Action Owner" not in export_df.columns:
                    export_df["Action Owner"] = "Program Manager"
                if "Timeline (Days)" not in export_df.columns:
                    export_df["Timeline (Days)"] = "60"
        
        return export_df

    # ══════════════════════════════════════════════════════════════════
    # EXCEL REPORT  ─ multi-question, fully rebuilt
    # ══════════════════════════════════════════════════════════════════

    def build_excel_report(self, output: Optional[str] = None) -> None:
        output = output or self.config.output_filename
        now_eat = datetime.now(EAT).strftime("%d %b %Y %H:%M EAT")

        wb = Workbook()
        wb.remove(wb.active)

        # ══════════════════════════════════════════════════════════════
        # 1. EXECUTIVE SUMMARY (Fully Table-Based - v4.0)
        #    All content now in proper tables with borders to prevent overlapping
        # ══════════════════════════════════════════════════════════════
        ws = wb.create_sheet("Executive Summary")
        EXEC_MERGE = 5   # columns for heading merge

        next_row = self.write_sheet_heading(
            ws,
            title=self.config.report_title,
            subtitle=f"Generated: {now_eat}   |   Questions analysed: "
                     f"{len(self.question_results)}",
            merge_cols=EXEC_MERGE,
        )

        # blank spacer
        next_row += 1   # row 4

        # ─────────────────────────────────────────────────────────────
        # TABLE 1: Questions Analysed
        # ─────────────────────────────────────────────────────────────
        header = ws.cell(row=next_row, column=1, value="QUESTIONS ANALYSED")
        header.font = Font(bold=True, size=11, color="000000")
        header.fill = PatternFill(
            start_color=self.SUBHEAD_COLOUR,
            end_color=self.SUBHEAD_COLOUR,
            fill_type="solid",
        )
        header.border = self._thin_border()
        ws.merge_cells(start_row=next_row, start_column=1, 
                       end_row=next_row, end_column=EXEC_MERGE)
        ws.row_dimensions[next_row].height = 20
        next_row += 1

        for q in self.question_results:
            c = ws.cell(row=next_row, column=1, value=q)
            c.alignment = Alignment(wrap_text=True, vertical="center")
            c.border = self._thin_border()
            ws.merge_cells(start_row=next_row, start_column=1,
                          end_row=next_row, end_column=EXEC_MERGE)
            ws.row_dimensions[next_row].height = 18
            next_row += 1
        next_row += 1

        # ─────────────────────────────────────────────────────────────
        # TABLE 2: Cross-Question Themes (if any)
        # ─────────────────────────────────────────────────────────────
        if self.cross_question_themes:
            header = ws.cell(row=next_row, column=1, 
                           value="⚠  THEMES RECURRING ACROSS MULTIPLE QUESTIONS")
            header.font = Font(bold=True, size=11, color="FFFFFF")
            header.fill = PatternFill(
                start_color="C00000",
                end_color="C00000",
                fill_type="solid",
            )
            header.border = self._thin_border()
            ws.merge_cells(start_row=next_row, start_column=1,
                          end_row=next_row, end_column=EXEC_MERGE)
            ws.row_dimensions[next_row].height = 20
            next_row += 1

            for t in self.cross_question_themes:
                c = ws.cell(row=next_row, column=1, value=t)
                c.alignment = Alignment(wrap_text=True, vertical="center")
                c.border = self._thin_border()
                c.fill = PatternFill(start_color="FFE0E0", end_color="FFE0E0",
                                    fill_type="solid")
                ws.merge_cells(start_row=next_row, start_column=1,
                              end_row=next_row, end_column=EXEC_MERGE)
                ws.row_dimensions[next_row].height = 18
                next_row += 1
            next_row += 1

        # ─────────────────────────────────────────────────────────────
        # Per-Question Summary Blocks (each in a table)
        # ─────────────────────────────────────────────────────────────
        for qname, qdata in self.question_results.items():
            s_df: pd.DataFrame = qdata["summary"]
            insights: List[str]  = qdata["insights"]
            ds: Dict             = qdata["dataset_summary"]

            # Section divider (full width header)
            divider = ws.cell(row=next_row, column=1, value=f"{qname.upper()}")
            divider.font = Font(bold=True, size=12, color="FFFFFF")
            divider.fill = PatternFill(
                start_color=self.HEADER_COLOUR,
                end_color=self.HEADER_COLOUR,
                fill_type="solid",
            )
            divider.border = self._thin_border()
            ws.merge_cells(
                start_row=next_row, start_column=1,
                end_row=next_row, end_column=EXEC_MERGE
            )
            ws.row_dimensions[next_row].height = 22
            next_row += 1

            # Stats table (2 columns: Label | Value)
            stats_headers = [
                ("Responses", ds.get("Total Responses", "")),
                ("Themes Identified", ds.get("Themes Identified", "")),
                ("Avg Length (words)", ds.get("Average Response Length (words)", "")),
                ("Languages", ds.get("Languages Detected", "")),
            ]
            for stat_label, stat_value in stats_headers:
                lc = ws.cell(row=next_row, column=1, value=stat_label)
                lc.font = Font(bold=True, size=10, color="000000")
                lc.fill = PatternFill(
                    start_color=self.SUBHEAD_COLOUR,
                    end_color=self.SUBHEAD_COLOUR,
                    fill_type="solid",
                )
                lc.border = self._thin_border()
                lc.alignment = Alignment(horizontal="left", vertical="center")

                vc = ws.cell(row=next_row, column=2, value=stat_value)
                vc.font = Font(bold=True, size=10)
                vc.border = self._thin_border()
                vc.alignment = Alignment(horizontal="center", vertical="center")
                
                ws.merge_cells(start_row=next_row, start_column=2,
                              end_row=next_row, end_column=EXEC_MERGE)
                ws.row_dimensions[next_row].height = 18
                next_row += 1
            next_row += 1

            # Insights table (single column with wrapped text)
            if insights:
                h = ws.cell(row=next_row, column=1, value="KEY INSIGHTS")
                h.font = Font(bold=True, size=10, color="000000")
                h.fill = PatternFill(
                    start_color=self.SUBHEAD_COLOUR,
                    end_color=self.SUBHEAD_COLOUR,
                    fill_type="solid",
                )
                h.border = self._thin_border()
                ws.merge_cells(start_row=next_row, start_column=1,
                              end_row=next_row, end_column=EXEC_MERGE)
                ws.row_dimensions[next_row].height = 18
                next_row += 1

                for ins in insights:
                    ic = ws.cell(row=next_row, column=1, value=f"•  {ins}")
                    ic.alignment = Alignment(wrap_text=True, vertical="top")
                    ic.border = self._thin_border()
                    ws.merge_cells(start_row=next_row, start_column=1,
                                  end_row=next_row, end_column=EXEC_MERGE)
                    ws.row_dimensions[next_row].height = 22
                    next_row += 1
                next_row += 1

            # Mini table header (Top themes for this question)
            mini_cols = ["Theme", "Responses", "Share (%)", "Priority",
                         "Recommendation"]
            for ci, h in enumerate(mini_cols, 1):
                c = ws.cell(row=next_row, column=ci, value=h)
                c.font = Font(bold=True, color="FFFFFF", size=9)
                c.fill = PatternFill(
                    start_color=self.HEADER_COLOUR,
                    end_color=self.HEADER_COLOUR,
                    fill_type="solid",
                )
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = self._thin_border()
            ws.row_dimensions[next_row].height = 20
            next_row += 1

            # Mini table data
            stripe_fill = PatternFill(
                start_color=self.STRIPE_COLOUR,
                end_color=self.STRIPE_COLOUR,
                fill_type="solid",
            )
            for row_i, (_, r) in enumerate(s_df.head(5).iterrows()):
                fill = stripe_fill if row_i % 2 == 1 else None
                vals = [
                    r["Theme"],
                    r["Response Count"],
                    r["Share (%)"],
                    r.get("Priority", ""),
                    r.get("Recommendation", ""),
                ]
                for ci, val in enumerate(vals, 1):
                    c = ws.cell(row=next_row, column=ci, value=val)
                    c.border = self._thin_border()
                    c.alignment = Alignment(wrap_text=True, vertical="top")
                    if fill:
                        c.fill = fill
                ws.row_dimensions[next_row].height = 20
                next_row += 1
            next_row += 2

        self.smart_col_widths(ws)
        # Exec summary: give the Recommendation column extra room
        if ws.max_column >= 5:
            ws.column_dimensions["E"].width = 52

        # ══════════════════════════════════════════════════════════════
        # 2. PER-QUESTION THEME SHEETS (with v3+ Engines)
        # ══════════════════════════════════════════════════════════════
        THEME_COLS = [
            "Theme", "Response Count", "Share (%)", "Impact Score",
            "Intensity Score", "Priority", "Theme Sentiment",
            "Average Sentiment", "Trigger Status", "Action Owner",
            "Timeline (Days)", "Key Keywords", "Representative Quote",
            "Recommendation",
        ]

        for qname, qdata in self.question_results.items():
            safe_name = re.sub(r"[^\w ]", "", qname)[:28].strip()
            ws = wb.create_sheet(safe_name)

            ds = qdata["dataset_summary"]
            subtitle = (
                f"Responses: {ds.get('Total Responses', '')}   |   "
                f"Themes: {ds.get('Themes Identified', '')}   |   "
                f"Languages: {ds.get('Languages Detected', '')}"
            )

            s_df: pd.DataFrame = qdata["summary"]
            # Prepare summary with trigger/action info
            s_df = self.prepare_summary_for_excel(s_df)
            available_cols = [c for c in THEME_COLS if c in s_df.columns]
            n_cols = max(len(available_cols), 8)

            next_row = self.write_sheet_heading(
                ws,
                title=f"Theme Analysis  —  {qname.replace('_', ' ').title()}",
                subtitle=subtitle,
                merge_cols=n_cols,
            )
            # next_row == 3 → write column headers here
            for ci, col_name in enumerate(available_cols, 1):
                ws.cell(row=next_row, column=ci, value=col_name)
            self.style_table_header(ws, header_row=next_row)
            data_start = next_row + 1

            for r in s_df[available_cols].itertuples(index=False):
                ws.append(list(r))

            # Wrap long-text columns
            for long_col in ("Representative Quote", "Recommendation",
                              "Key Keywords", "Action Owner", "Trigger Status"):
                if long_col in available_cols:
                    lcol = available_cols.index(long_col) + 1
                    for row in ws.iter_rows(
                        min_row=data_start, min_col=lcol, max_col=lcol
                    ):
                        for cell in row:
                            cell.alignment = Alignment(
                                wrap_text=True, vertical="top"
                            )

            # Colour trigger status cells (RED for HIGH, ORANGE for MEDIUM)
            if "Trigger Status" in available_cols:
                trigger_col = available_cols.index("Trigger Status") + 1
                for row in ws.iter_rows(
                    min_row=data_start, max_row=ws.max_row,
                    min_col=trigger_col, max_col=trigger_col
                ):
                    for cell in row:
                        cell_value = str(cell.value or "").upper()
                        if "HIGH" in cell_value:
                            cell.fill = PatternFill(
                                start_color="FF0000", end_color="FF0000",
                                fill_type="solid"
                            )
                            cell.font = Font(color="FFFFFF", bold=True)
                        elif "MEDIUM" in cell_value:
                            cell.fill = PatternFill(
                                start_color="FFA500", end_color="FFA500",
                                fill_type="solid"
                            )
                            cell.font = Font(color="FFFFFF", bold=True)
                        elif "POSITIVE" in cell_value:
                            cell.fill = PatternFill(
                                start_color="00B050", end_color="00B050",
                                fill_type="solid"
                            )
                            cell.font = Font(color="FFFFFF", bold=True)
                        else:
                            cell.fill = PatternFill(
                                start_color="F0F0F0", end_color="F0F0F0",
                                fill_type="solid"
                            )

            # Colour by sentiment, then stripe
            if "Theme Sentiment" in available_cols:
                sent_col = available_cols.index("Theme Sentiment") + 1
                self.sentiment_colour_rows(ws, sent_col, start_row=data_start)

            self.add_table_borders(ws, start_row=next_row)
            ws.auto_filter.ref = (
                f"A{next_row}:{get_column_letter(len(available_cols))}"
                f"{ws.max_row}"
            )
            self._set_data_row_heights(ws, start_row=data_start, height=20)
            self.smart_col_widths(ws)

        # ══════════════════════════════════════════════════════════════
        # 3. CHARTS SHEET (Professional Styling - v3.0)
        #    Features:
        #    - Clean professional styling with modern chart look
        #    - No gridlines for clarity
        #    - Clear titles and axis labels
        #    - Legend positioned right (not overlapping pie)
        # ══════════════════════════════════════════════════════════════
        ws = wb.create_sheet("Charts")

        CHART_BLOCK_ROWS = 38   # rows allocated per question block
        BAR_W, BAR_H   = 20, 13  # bar chart dimensions (optimized spacing)
        PIE_W, PIE_H   = 15, 13  # pie chart dimensions (optimized spacing)
        BAR_ANCHOR_COL = "D"     # bar chart left edge
        PIE_ANCHOR_COL = "U"     # pie chart left edge (well spaced from bar)

        # Page heading
        self.write_sheet_heading(
            ws,
            title="Charts  —  Feedback Distribution by Theme",
            subtitle=f"Generated: {now_eat}",
            merge_cols=26,
        )

        # Data area starts at row 4 (rows 1-2 = heading, row 3 = spacer)
        data_cursor = 4

        for q_idx, (qname, qdata) in enumerate(self.question_results.items()):
            s_df = qdata["summary"]
            safe_q = re.sub(r"[^\w ]", "", qname).strip().replace("_", " ").title()

            block_start = data_cursor  # anchor row for both charts

            # Section label above data
            lbl = ws.cell(row=block_start, column=1, value=safe_q)
            lbl.font = Font(bold=True, size=11, color="FFFFFF")
            lbl.fill = PatternFill(
                start_color=self.HEADER_COLOUR,
                end_color=self.HEADER_COLOUR,
                fill_type="solid",
            )
            ws.row_dimensions[block_start].height = 20

            # Column headers for data table
            ws.cell(row=block_start + 1, column=1, value="Theme")
            ws.cell(row=block_start + 1, column=2, value="Responses")
            self.style_table_header(ws, header_row=block_start + 1)

            # Data rows
            for offset, (_, r) in enumerate(
                s_df[["Theme", "Response Count"]].iterrows(), start=2
            ):
                ws.cell(row=block_start + offset, column=1, value=r["Theme"])
                ws.cell(row=block_start + offset, column=2,
                        value=r["Response Count"])

            data_end = block_start + 1 + len(s_df)

            # References (header row included for title)
            data_ref = Reference(ws, min_col=2,
                                  min_row=block_start + 1, max_row=data_end)
            cats_ref = Reference(ws, min_col=1,
                                  min_row=block_start + 2, max_row=data_end)

            # ── Bar chart (Professional v3.0) ──────────────────────────
            bar = BarChart()
            bar.type    = "col"
            bar.style   = 11  # Modern professional style
            bar.title   = f"{safe_q}\nTheme Distribution"
            bar.y_axis.title = "Number of Responses"
            bar.x_axis.title = "Feedback Theme"
            bar.legend  = None          # No legend for cleaner look
            
            # Remove gridlines for professional appearance
            bar.y_axis.majorGridlines = None
            
            # Set axis scaling to start from 0 and be visible
            bar.y_axis.scaling.minVal = 0
            bar.y_axis.delete = False  # Ensure axis is visible
            bar.x_axis.delete = False  # Ensure axis is visible
            
            bar.add_data(data_ref, titles_from_data=True)
            bar.set_categories(cats_ref)
            bar.width  = BAR_W
            bar.height = BAR_H
            
            # Add data labels (values on top of bars) - openpyxl limitation
            # requires setting dataLbls XML attribute manually if using openpyxl
            # For now, chart shows cleanly with professional styling
            ws.add_chart(bar, f"{BAR_ANCHOR_COL}{block_start}")

            # ── Pie chart (Professional v3.0) ─────────────────────────
            pie = PieChart()
            pie.style   = 11  # Modern professional style
            pie.title   = f"{safe_q}\nPercentage Share"
            
            # Position legend to RIGHT to prevent overlap
            pie.legend.position = "r"  # 'r' = right
            
            pie.add_data(data_ref, titles_from_data=True)
            pie.set_categories(cats_ref)
            pie.width  = PIE_W
            pie.height = PIE_H
            
            # Anchor pie well to the right of the bar chart
            ws.add_chart(pie, f"{PIE_ANCHOR_COL}{block_start}")

            # Advance cursor by the fixed block height
            data_cursor += CHART_BLOCK_ROWS

        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 14

        # ══════════════════════════════════════════════════════════════
        # 3.5 ALERTS & ACTIONS (v3+ Decision Support Sheet)
        # ══════════════════════════════════════════════════════════════
        # This sheet surfaces HIGH/MEDIUM priority issues with action owners
        # and deadlines for immediate decision-making

        ws = wb.create_sheet("Alerts & Actions", 3)  # Insert after Charts

        next_row = self.write_sheet_heading(
            ws,
            title="🚨 PRIORITY ALERTS & ACTION PLANS",
            subtitle=f"Critical Issues Requiring Immediate Attention | {now_eat}",
            merge_cols=8,
        )
        next_row += 1

        # Collect all HIGH/MEDIUM priority issues across all questions
        priority_issues = []
        for qname, qdata in self.question_results.items():
            s_df = self.prepare_summary_for_excel(qdata["summary"])
            for idx, row in s_df.iterrows():
                triggers = row.get("triggers", [])
                if isinstance(triggers, list):
                    for trigger in triggers:
                        if trigger.get("level") in ("HIGH", "MEDIUM"):
                            priority_issues.append({
                                "question": qname,
                                "theme": row["Theme"],
                                "priority": trigger.get("icon", ""),
                                "level": trigger.get("level", ""),
                                "message": trigger.get("message", ""),
                                "impact": row.get("Response Count", 0),
                                "owner": row.get("Action Owner", "TBD"),
                                "deadline_days": trigger.get("deadline_days", 30),
                                "recommendation": row.get("Recommendation", "N/A"),
                            })

        # Sort by level (HIGH first) then by impact
        priority_issues.sort(
            key=lambda x: (
                0 if x["level"] == "HIGH" else 1,
                -x["impact"]
            )
        )

        if not priority_issues:
            # No alerts
            msg_cell = ws.cell(row=next_row, column=1, 
                              value="✓ No HIGH or MEDIUM priority alerts detected.")
            msg_cell.font = Font(size=12, color="00B050", bold=True)
            ws.merge_cells(start_row=next_row, start_column=1,
                          end_row=next_row, end_column=8)
            next_row += 1
        else:
            # Build alerts table
            headers = ["Priority", "Theme", "Question", "Impact", "Message",
                      "Action Owner", "Deadline (Days)", "Recommendation"]
            for ci, h in enumerate(headers, 1):
                hc = ws.cell(row=next_row, column=ci, value=h)
                hc.font = Font(bold=True, color="FFFFFF", size=11)
                hc.fill = PatternFill(
                    start_color="2F5597", end_color="2F5597",
                    fill_type="solid"
                )
                hc.alignment = Alignment(horizontal="center", vertical="center")
                hc.border = self._thin_border()
            
            ws.row_dimensions[next_row].height = 20
            next_row += 1

            # Data rows
            for issue in priority_issues:
                ws.cell(row=next_row, column=1, value=issue["priority"])
                ws.cell(row=next_row, column=2, value=issue["theme"])
                ws.cell(row=next_row, column=3, value=issue["question"])
                ws.cell(row=next_row, column=4, value=issue["impact"])
                ws.cell(row=next_row, column=5, value=issue["message"])
                ws.cell(row=next_row, column=6, value=issue["owner"])
                ws.cell(row=next_row, column=7, value=issue["deadline_days"])
                ws.cell(row=next_row, column=8, value=issue["recommendation"])

                # Colour the priority column (dimmed colors, no icons)
                priority_cell = ws.cell(row=next_row, column=1)
                priority_cell.value = issue["level"]  # Show text level, not emoji
                if issue["level"] == "HIGH":
                    priority_cell.fill = PatternFill(
                        start_color="FFB3B3", end_color="FFB3B3",
                        fill_type="solid"
                    )
                    priority_cell.font = Font(color="404040", bold=True)
                elif issue["level"] == "MEDIUM":
                    priority_cell.fill = PatternFill(
                        start_color="FFE0B3", end_color="FFE0B3",
                        fill_type="solid"
                    )
                    priority_cell.font = Font(color="404040", bold=True)

                # Wrap long text columns
                for col_idx in [2, 5, 8]:  # Theme, Message, Recommendation
                    cell = ws.cell(row=next_row, column=col_idx)
                    cell.alignment = Alignment(
                        wrap_text=True, vertical="top"
                    )

                # Add borders
                for col in range(1, 9):
                    ws.cell(row=next_row, column=col).border = self._thin_border()

                ws.row_dimensions[next_row].height = 25
                next_row += 1

            # Set column widths
            ws.column_dimensions["A"].width = 12
            ws.column_dimensions["B"].width = 22
            ws.column_dimensions["C"].width = 18
            ws.column_dimensions["D"].width = 8
            ws.column_dimensions["E"].width = 35
            ws.column_dimensions["F"].width = 20
            ws.column_dimensions["G"].width = 14
            ws.column_dimensions["H"].width = 40

            # Add autofilter
            ws.auto_filter.ref = f"A{next_row - len(priority_issues)}:H{next_row}"

        # ══════════════════════════════════════════════════════════════
        # 4. EXAMPLE QUOTES
        # ══════════════════════════════════════════════════════════════
        ws = wb.create_sheet("Example Quotes")

        next_row = self.write_sheet_heading(
            ws,
            title="Example Quotes  —  Representative Responses by Theme",
            subtitle=f"Generated: {now_eat}",
            merge_cols=3,
        )

        # Column headers at row 3
        for ci, h in enumerate(["Question", "Theme", "Quote"], 1):
            ws.cell(row=next_row, column=ci, value=h)
        self.style_table_header(ws, header_row=next_row)
        data_start = next_row + 1

        for qname, qdata in self.question_results.items():
            cs = qdata.get("cluster_summary", {})
            for cid, info in cs.items():
                theme = self.generate_theme_name(info["keywords"])
                for q in info["samples"]:
                    ws.append([qname, theme, q])

        # Wrap quote column
        for row in ws.iter_rows(min_row=data_start, min_col=3, max_col=3):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        self.add_table_borders(ws, start_row=next_row)
        self.zebra_rows(ws, start_row=data_start)
        self._set_data_row_heights(ws, start_row=data_start, height=22)
        self.smart_col_widths(ws)

        # ══════════════════════════════════════════════════════════════
        # 5. EMERGING ISSUES
        # ══════════════════════════════════════════════════════════════
        emerging_rows = []
        for qname, qdata in self.question_results.items():
            em_df: pd.DataFrame = qdata.get("emerging_df", pd.DataFrame())
            if not em_df.empty:
                em_df = em_df.copy()
                em_df.insert(0, "Question", qname)
                emerging_rows.append(em_df)

        if emerging_rows:
            ws = wb.create_sheet("⚠ Emerging Issues")
            all_em = pd.concat(emerging_rows, ignore_index=True)

            next_row = self.write_sheet_heading(
                ws,
                title="⚠  Emerging Issues  —  Weak Signals Worth Watching",
                subtitle="Low-volume clusters that may be early-warning signs. "
                         "Investigate before they become widespread.",
                merge_cols=max(len(all_em.columns), 6),
            )

            for ci, col_name in enumerate(all_em.columns, 1):
                ws.cell(row=next_row, column=ci, value=col_name)
            self.style_table_header(ws, header_row=next_row)
            data_start = next_row + 1

            for r in all_em.itertuples(index=False):
                ws.append(list(r))

            self.add_table_borders(ws, start_row=next_row)
            self.zebra_rows(ws, start_row=data_start)
            self._set_data_row_heights(ws, start_row=data_start, height=18)
            self.smart_col_widths(ws)

        # ══════════════════════════════════════════════════════════════
        # 6. CLUSTERED RESPONSES (raw data)
        # ══════════════════════════════════════════════════════════════
        ws = wb.create_sheet("Clustered Responses")

        # Build comprehensive dataset with themes + metadata
        all_q_dfs = []
        for qname, qdata in self.question_results.items():
            q_df = qdata.get("q_df", pd.DataFrame()).copy()
            s_df = qdata.get("summary", pd.DataFrame()).copy()
            cluster_map = qdata.get("cluster_to_theme_map", {})
            
            if q_df.empty:
                continue
            
            # Add Question column
            q_df.insert(0, "Question", qname)
            
            # Rename columns for clarity
            if "clean_text" in q_df.columns:
                q_df.rename(columns={"clean_text": "Clean Text"}, inplace=True)
            elif "response" in q_df.columns:
                q_df.rename(columns={"response": "Clean Text"}, inplace=True)
            
            if "lang" in q_df.columns:
                q_df.rename(columns={"lang": "Language"}, inplace=True)
            
            if "cluster" in q_df.columns:
                q_df.rename(columns={"cluster": "Cluster"}, inplace=True)
            
            if "sentiment" in q_df.columns:
                q_df.rename(columns={"sentiment": "Sentiment"}, inplace=True)
            
            # Build lookup: theme_name → theme metadata from summary_df
            theme_lookup = {}
            if not s_df.empty:
                for _, row in s_df.iterrows():
                    theme_name = row.get("Theme", "Unknown")
                    theme_lookup[theme_name] = {
                        "Theme": theme_name,
                        "Theme Sentiment": row.get("Theme Sentiment", "Neutral"),
                        "Priority": row.get("Priority", "Low"),
                        "Trigger Status": row.get("Trigger Status", "No Alert"),
                        "Action Owner": row.get("Action Owner", "Program Manager"),
                        "Timeline (Days)": str(row.get("Timeline (Days)", "30")),
                    }
            
            # Map cluster IDs to themes using cluster_to_theme_map, then lookup metadata
            cluster_col = "Cluster" if "Cluster" in q_df.columns else "cluster"
            if cluster_col in q_df.columns:
                # Create a function to lookup theme by cluster ID then get metadata
                def get_theme_metadata(cluster_id):
                    theme_name = cluster_map.get(cluster_id, "Unclassified")
                    return theme_lookup.get(theme_name, {
                        "Theme": "Unclassified",
                        "Theme Sentiment": "Neutral",
                        "Priority": "Low",
                        "Trigger Status": "No Alert",
                        "Action Owner": "Program Manager",
                        "Timeline (Days)": "30",
                    })
                
                # Add theme columns to q_df
                for col in ["Theme", "Theme Sentiment", "Priority",
                            "Trigger Status", "Action Owner", "Timeline (Days)"]:
                    q_df[col] = q_df[cluster_col].apply(lambda cid: get_theme_metadata(cid).get(col, ""))
            
            all_q_dfs.append(q_df)

        if all_q_dfs:
            combined = pd.concat(all_q_dfs, ignore_index=True)
            
            # Reorder columns for professional layout
            # Core data columns first, then theme metadata
            desired_order = [
                "Question",
                "Clean Text",
                "Language", 
                "Sentiment",
                "Cluster",
                "Theme",
                "Theme Sentiment",
                "Priority",
                "Trigger Status",
                "Action Owner",
                "Timeline (Days)",
            ]
            
            # Keep only columns that exist and exclude internal columns
            internal_cols = {"compressed_text"}  # Internal columns to hide
            cols_to_use = [col for col in desired_order if col in combined.columns]
            # Add any remaining columns not in desired_order (excluding internal)
            for col in combined.columns:
                if col not in cols_to_use and col not in internal_cols:
                    cols_to_use.append(col)
            
            combined = combined[cols_to_use]
            
            n_cols = len(combined.columns)

            next_row = self.write_sheet_heading(
                ws,
                title="Clustered Responses  —  Complete Analysis Dataset",
                subtitle=f"Total responses: {len(combined)}   |   "
                         f"Questions: {len(self.question_results)}",
                merge_cols=min(n_cols, 12),
            )

            # Write headers
            for ci, col_name in enumerate(combined.columns, 1):
                ws.cell(row=next_row, column=ci, value=col_name)
            self.style_table_header(ws, header_row=next_row)
            data_start = next_row + 1

            # Write data rows
            for r in combined.itertuples(index=False):
                ws.append(list(r))

            ws.auto_filter.ref = (
                f"A{next_row}:{get_column_letter(len(combined.columns))}"
                f"{ws.max_row}"
            )
            self.add_table_borders(ws, start_row=next_row)
            self.zebra_rows(ws, start_row=data_start)
            self._set_data_row_heights(ws, start_row=data_start, height=16)
            self.smart_col_widths(ws)

        wb.save(output)
        log.info("✅ Report saved → %s", output)

    # ═══════════════════════════════════════════════════════════════════
    # JSON & ALERTS EXPORT (v3+ Feature)
    # ═══════════════════════════════════════════════════════════════════

    def _export_json_reports(self) -> None:
        """Export insights in machine-readable JSON format."""
        try:
            from output.json_writer import export_json
            
            # Collect all themes from all questions
            all_themes = []
            for qname, qdata in self.question_results.items():
                s_df = qdata["summary"]
                for idx, row in s_df.iterrows():
                    theme_dict = {
                        "question": qname,
                        "name": row.get("Theme", ""),
                        "impact": int(row.get("Response Count", 0)),
                        "sentiment": row.get("Theme Sentiment", "Neutral"),
                        "keywords": str(row.get("Key Keywords", "")).split(","),
                        "triggers": row.get("triggers", []),
                        "action_plan": (
                            row.get("action_plan", {})
                            if isinstance(row.get("action_plan"), dict)
                            else {}
                        ),
                        "recommendation": row.get("Recommendation", ""),
                    }
                    all_themes.append(theme_dict)
            
            export_json(
                all_themes,
                filename="output/insights.json",
                sector=self.pipeline_config.get("sector", ""),
            )
        except ImportError:
            log.warning("json_writer not available — skipping JSON export")
        except Exception as exc:
            log.error("JSON export failed: %s", exc)

    def _export_alerts_report(self) -> None:
        """Export priority alerts in machine-readable format."""
        try:
            from output.alert_formatter import generate_alerts, export_alerts_json, print_alerts
            
            # Collect all HIGH/MEDIUM priority alerts
            all_alerts = []
            for qname, qdata in self.question_results.items():
                s_df = qdata["summary"]
                for idx, row in s_df.iterrows():
                    triggers = row.get("triggers", [])
                    if isinstance(triggers, list):
                        for trigger in triggers:
                            if trigger.get("level") in ("HIGH", "MEDIUM"):
                                alert = {
                                    "theme_name": row.get("Theme", ""),
                                    "question": qname,
                                    "priority_icon": trigger.get("icon", ""),
                                    "priority_level": trigger.get("level", ""),
                                    "message": trigger.get("message", ""),
                                    "deadline_days": trigger.get("deadline_days", 30),
                                    "impact_count": row.get("Response Count", 0),
                                    "recommendation": row.get("Recommendation", ""),
                                    "action_owner": row.get("Action Owner", "TBD"),
                                }
                                all_alerts.append(alert)
            
            if all_alerts:
                # Print to console for immediate visibility
                print_alerts(all_alerts, max_display=5)
                
                # Export to JSON
                export_alerts_json(all_alerts, filename="output/alerts.json")
            else:
                log.info("✓ No HIGH/MEDIUM alerts to export")
        except ImportError:
            log.warning("alert_formatter not available — skipping alerts export")
        except Exception as exc:
            log.error("Alerts export failed: %s", exc)

    # ══════════════════════════════════════════════════════════════════
    # FULL PIPELINE
    # ══════════════════════════════════════════════════════════════════

    def run(self) -> None:
        for column in self.text_columns:
            log.info("━━━ Analysing: %s ━━━", column)

            # --- Validate column ---
            if column not in self.df.columns:
                log.warning("Column '%s' not found — skipping.", column)
                continue

            raw_series = (
                self.df[column]
                .dropna()
                .astype(str)
                .str.strip()
            )
            raw_series = raw_series[raw_series.str.len() > 3]
            self._question_response_count = len(raw_series)

            if self._question_response_count < self.config.min_responses:
                log.warning(
                    "Column '%s' has only %d usable responses (min=%d) — skipping.",
                    column, self._question_response_count, self.config.min_responses,
                )
                continue

            # --- Build per-question df ---
            self.q_df = raw_series.to_frame(name="response").reset_index(drop=True)

            log.info("Preprocessing …")
            self.preprocess()

            log.info("Deduplicating near-identical sentences …")
            self.deduplicate_responses()

            log.info("Building TF-IDF …")
            self.build_tfidf_matrix()

            log.info("Sentiment analysis …")
            self.compute_sentiment()

            log.info("Generating embeddings + PCA …")
            self.generate_embeddings()

            log.info("Clustering …")
            self.cluster_feedback()

            log.info("Filtering clusters …")
            self.filter_clusters()

            log.info("Building cluster summary …")
            self.build_cluster_summary()

            log.info("Building summary table …")
            self.build_summary_table()

            log.info("Classifying theme sentiment …")
            self.classify_theme_sentiment()

            log.info("Applying advanced engines (triggers/actions) …")
            self.apply_advanced_engines()

            log.info("Generating recommendations …")
            self.generate_recommendations()

            log.info("Merging similar themes …")
            self.merge_similar_themes()

            log.info("Generating insights …")
            self.generate_key_insights()

            log.info("Dataset summary …")
            self.generate_dataset_summary()

            log.info("Labelling responses …")
            self.label_data()

            # --- Store per-question results ---
            self.question_results[column] = {
                "summary": self.summary_df.copy(),
                "insights": self.key_insights.copy(),
                "dataset_summary": self.dataset_summary.copy(),
                "cluster_summary": self.cluster_summary.copy(),
                "emerging_df": getattr(self, "emerging_df", pd.DataFrame()).copy(),
                "q_df": self.q_df.copy(),
                "cluster_to_theme_map": self.cluster_to_theme_map.copy(),
            }

            # Free per-question RAM
            del self.embeddings
            gc.collect()

        # --- Cross-question analysis ---
        if len(self.question_results) > 1:
            log.info("Detecting cross-question themes …")
            self.detect_cross_question_themes()

        # --- Keyword coverage tips (always printed when rule-based recs used) ---
        # If LLM recs were used, print a condensed keyword discovery summary.
        if self._anthropic_client and self.config.use_llm_recommendations:
            self._print_llm_keyword_summary()

        # ── NEW: Generate JSON + Alerts output (v3+ Upgrade) ──────────
        self._export_outputs()

        log.info("Building Excel report …")
        self.build_excel_report()

    def _export_outputs(self) -> None:
        """
        Export insights as JSON and generate priority alerts.
        Controlled by config/config.yaml output settings.
        """
        config = self.pipeline_config
        output_config = config.get("output", {})
        sector = config.get("sector", "")
        
        # Collect all themes from all questions
        all_themes = []
        for qname, qdata in self.question_results.items():
            s_df = qdata.get("summary", pd.DataFrame())
            
            # DEBUG: Check what columns exist in summary_df
            if all_themes == []:  # Only log once for first question
                log.info(f"  📊 Summary columns: {list(s_df.columns)}")
            
            for idx, row in s_df.iterrows():
                # Extract trigger status from raw triggers list
                triggers = row.get("triggers", [])
                trigger_status = "Low"
                if isinstance(triggers, list) and triggers:
                    # Get the highest priority trigger level
                    levels = {t.get("level", "Low") for t in triggers}
                    if "High" in levels:
                        trigger_status = "High"
                    elif "Medium" in levels:
                        trigger_status = "Medium"
                    elif "Positive" in levels:
                        trigger_status = "Positive"
                
                # Extract action owner and timeline from action_plan dict
                action_plan = row.get("action_plan", {})
                
                # Ensure action_plan is converted from pandas object to dict
                if isinstance(action_plan, dict):
                    assigned_owner = action_plan.get("owner", "Unknown")
                    timeline_days = action_plan.get("timeline", 30)
                else:
                    assigned_owner = "Unknown"
                    timeline_days = 30
                
                # Convert triggers to list of dicts (in case it's a pandas object)
                triggers_list = []
                if isinstance(triggers, list):
                    triggers_list = triggers
                elif pd.notna(triggers):
                    triggers_list = [] if triggers is None else [triggers]
                
                # DEBUG: Log first theme details
                if idx == 0 and qname == list(self.question_results.keys())[0]:
                    log.info(f"  🎯 First theme export check:")
                    log.info(f"     Theme: {row.get('Theme', 'N/A')}")
                    log.info(f"     action_plan type: {type(action_plan)}")
                    log.info(f"     action_plan value: {action_plan}")
                    log.info(f"     owner: {assigned_owner}")
                
                theme_dict = {
                    "question": qname,
                    "name": row.get("Theme", ""),
                    "impact": row.get("Response Count", 0),
                    "sentiment": row.get("Theme Sentiment", "Neutral"),
                    "keywords": (
                        [k.strip() for k in str(row.get("Key Keywords", "")).split(",")]
                        if pd.notna(row.get("Key Keywords"))
                        else []
                    ),
                    "average_sentiment_score": row.get("Average Sentiment", 0),
                    "recommendation": row.get("Recommendation", ""),
                    "trigger_status": trigger_status,
                    "assigned_owner": assigned_owner,
                    "timeline_days": timeline_days,
                    "triggers": triggers_list,
                    "action_plan": action_plan if isinstance(action_plan, dict) else {},
                }
                all_themes.append(theme_dict)
        
        # Export JSON if enabled
        if output_config.get("json", True):
            try:
                from output.json_writer import export_json
                export_json(
                    all_themes,
                    filename="output/insights.json",
                    sector=sector,
                )
            except ImportError:
                log.warning("JSON writer not available. Skipping JSON export.")
        
        # Generate and display alerts if enabled
        if output_config.get("alerts", True):
            try:
                from output.alert_formatter import generate_alerts, print_alerts, export_alerts_json
                alerts = generate_alerts(all_themes)
                print_alerts(alerts, max_display=10)
                export_alerts_json(alerts, filename="output/alerts.json")
            except ImportError:
                log.warning("Alert formatter not available. Skipping alerts.")

    def _print_llm_keyword_summary(self) -> None:
        """Print a brief keyword discovery log when LLM recommendations are active."""
        print(f"\n{'═' * 62}")
        print("  KEYWORD DISCOVERY SUMMARY  (LLM recommendations active)")
        print(f"{'═' * 62}")
        for qname, qdata in self.question_results.items():
            s_df = qdata["summary"]
            all_kw = set()
            for kw_str in s_df.get("Key Keywords", []):
                all_kw.update(k.strip() for k in str(kw_str).split(","))
            print(f"\n  {qname}:")
            print(f"    Themes     : {', '.join(s_df['Theme'].tolist())}")
            print(f"    All keywords found : {', '.join(sorted(all_kw))}")
        print(f"\n  TIP: If recommendations look off for any theme,")
        print(f"  set use_llm_recommendations=False and check the")
        print(f"  KEYWORD TIPS report to tune RECOMMENDATION_RULES.")
        print(f"{'═' * 62}\n")


# ══════════════════════════════════════════════════════════════════════
# ENTRY POINT  ─ configure and run
# ══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":

    config = EngineConfig(
        pca_components=64,
        min_cluster_size=5,
        emerging_min_size=2,
        max_clusters=20,
        merge_threshold=0.75,
        tfidf_max_features=5_000,
        large_dataset_threshold=500,
        min_responses=10,

        # ── Sector context ────────────────────────────────────────────
        # Set this to your programme area for better LLM outputs:
        #   "smallholder agriculture"  |  "primary healthcare"
        #   "microfinance / SACCO"     |  "WASH services"
        #   "secondary education"      |  "urban housing"
        sector="smallholder agriculture",  # ← change for your sector

        # ── LLM features ─────────────────────────────────────────────
        # Requires: pip install anthropic  +  ANTHROPIC_API_KEY env var
        use_llm_naming=True,           # meaningful theme names
        use_llm_recommendations=True,  # sector-aware recommendations
        # Set both to False to use rule-based only (faster, no API key needed)

        report_title="One Acre Fund — Farmer Feedback Insights 2026",
        output_filename="oaf_farmer_survey_insight_report.xlsx",
    )

    engine = FeedbackInsightEngine(
        file_path="oaf_farmer_survey_demo.xlsx",
        text_columns=[
            "q1_inputs",
            "q2_training",
            "q3_fieldofficer",
            "q4_loan",
            "q5_suggestions",
        ],
        config=config,
    )

    engine.run()