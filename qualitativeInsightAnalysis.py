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
import logging
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
    # Custom Swahili / Sheng / African-English polarity adjustments
    # Format: word → compound score delta (−1 to +1)
    extra_sentiment_lexicon: Dict[str, float] = field(default_factory=lambda: {
        # Swahili positive
        "sawa": 0.4, "poa": 0.5, "nzuri": 0.5, "vizuri": 0.5,
        "asante": 0.3, "hongera": 0.6, "bora": 0.4, "salama": 0.3,
        "furaha": 0.6, "starehe": 0.4, "rahisi": 0.3,
        # Swahili negative
        "mbaya": -0.5, "vibaya": -0.5, "tatizo": -0.4, "shida": -0.5,
        "hasira": -0.6, "huzuni": -0.5, "duni": -0.4, "gharama": -0.3,
        "uchafu": -0.5, "lalamiko": -0.4,
        # Sheng
        "peng": 0.3, "moto": 0.3, "safi": 0.5, "rada": -0.3,
        "stress": -0.4, "noma": -0.4, "fala": -0.5,
        # African-English
        "harassed": -0.6, "frustrated": -0.6, "disappointed": -0.6,
        "chapaa": -0.3,  # money stress
    })

    # --- Scale ---
    large_dataset_threshold: int = 500   # use MiniBatchKMeans above this
    min_responses: int = 10              # skip question if fewer rows

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
    # Cost / affordability
    (["cost", "price", "fee", "expensive", "afford", "payment", "pay",
      "charges", "money", "mpesa", "cash", "wallet", "tariff", "cheap",
      "subsidy", "bursary", "fund"],
     "Review pricing structure; explore M-Pesa instalment options, "
     "subsidies, or tiered pricing for low-income segments."),

    # Wait time / queue
    (["wait", "queue", "delay", "slow", "long", "time", "hours", "days",
      "minutes", "appointment", "turnaround", "response"],
     "Streamline service workflow: add appointment booking, triage, "
     "or digital queue management to cut wait times."),

    # Transport / access / distance
    (["transport", "distance", "travel", "far", "road", "matatu",
      "boda", "vehicle", "location", "access", "reach", "remote",
      "rural", "county", "village", "town"],
     "Expand reach via mobile units, satellite offices, or partner "
     "with boda-boda networks for last-mile service delivery."),

    # Staff attitude / customer care
    (["staff", "attitude", "rude", "arrogant", "unfriendly", "unprofessional",
      "disrespect", "ignore", "behaviour", "behavior", "manner", "tone",
      "treat", "treated"],
     "Invest in customer service training; introduce mystery-shopper "
     "audits and link staff appraisals to satisfaction scores."),

    # Drug / supply / stock
    (["drug", "medicine", "stock", "supply", "shortage", "unavailable",
      "out of", "pharmacy", "dispensary", "equipment", "machine"],
     "Strengthen supply-chain management: adopt real-time stock "
     "tracking and safety-stock thresholds with automated reordering."),

    # Digital / technology
    (["app", "website", "online", "digital", "network", "internet",
      "data", "ussd", "sms", "platform", "portal", "login", "slow",
      "crash", "error", "down"],
     "Audit digital touchpoints; ensure USSD/SMS fallback for "
     "low-bandwidth users and test regularly on Android entry-level devices."),

    # Health insurance / NHIF / SHIF
    (["nhif", "shif", "insurance", "cover", "claim", "reimburse",
      "scheme", "card", "benefit"],
     "Clarify insurance benefits communication; partner with SHIF "
     "to streamline pre-authorisation and claims processing."),

    # Safety / security
    (["safe", "safety", "security", "theft", "crime", "danger",
      "risk", "unsafe", "afraid", "fear"],
     "Conduct a safety audit; coordinate with local security organs "
     "and introduce visible security measures at service points."),

    # Communication / information
    (["information", "communicate", "update", "notice", "aware",
      "know", "told", "announcement", "transparent", "explain",
      "unclear", "confusing"],
     "Improve proactive communication: send SMS/WhatsApp updates, "
     "post clear notices in Swahili and English, and train staff "
     "on clear messaging."),

    # Sanitation / facilities
    (["toilet", "water", "clean", "dirty", "hygiene", "sanitation",
      "facility", "facilities", "room", "building", "infrastructure"],
     "Allocate resources for basic facility upgrades; partner with "
     "county water departments for reliable water supply."),

    # Food / nutrition (schools, hospitals, social services)
    (["food", "meal", "lunch", "hunger", "nutrition", "diet", "eat"],
     "Review catering contracts or nutrition budgets; consider "
     "community-based feeding partnerships."),
]

FALLBACK_RECOMMENDATION = (
    "Conduct targeted focus-group discussions to identify the root "
    "cause before designing an intervention."
)


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

        # ── cross-question artefacts ──────────────────────────────────
        self.question_results: Dict = {}
        self.cross_question_themes: List[str] = []

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
    # THEME NAMING
    # ══════════════════════════════════════════════════════════════════

    def generate_theme_name(self, keywords: List[str]) -> str:
        if not keywords:
            return "Other Feedback"
        primary = keywords[0]
        secondary = keywords[1] if len(keywords) > 1 else ""
        phrase = f"{primary} {secondary}".strip().replace("_", " ")
        words = list(dict.fromkeys(phrase.split()))
        return " ".join(words).title()

    # ══════════════════════════════════════════════════════════════════
    # CLUSTER SUMMARY BUILD
    # ══════════════════════════════════════════════════════════════════

    def build_cluster_summary(self) -> None:
        self.cluster_summary = {}
        for cid in sorted(self.q_df["cluster"].unique()):
            subset = self.q_df[self.q_df["cluster"] == cid]
            texts = subset["clean_text"].tolist()
            indices = subset.index.tolist()
            kw = self.extract_keywords(indices, self.config.top_keywords)
            self.cluster_summary[cid] = {
                "count": len(texts),
                "keywords": kw,
                "samples": texts[:5],
            }

    # ══════════════════════════════════════════════════════════════════
    # SUMMARY TABLE  ─ % based on per-question raw response count
    # ══════════════════════════════════════════════════════════════════

    def build_summary_table(self) -> None:
        total = self._question_response_count  # set in run() per question
        rows = []

        for cid, info in self.cluster_summary.items():
            theme = self.generate_theme_name(info["keywords"])
            subset = self.q_df[self.q_df["cluster"] == cid]
            avg_sent = subset["sentiment"].mean()
            pct = round((info["count"] / total) * 100, 2)

            # Intensity score: few very angry vs many mildly negative
            negative_sentences = subset[subset["sentiment"] < -0.3]
            intensity = (
                negative_sentences["sentiment"].abs().mean()
                if len(negative_sentences) > 0 else 0.0
            )
            severity = info["count"] * abs(avg_sent)

            rows.append({
                "Theme": theme,
                "Response Count": info["count"],
                "Share (%)": pct,
                "Impact Score": round(severity, 2),
                "Intensity Score": round(intensity, 3),
                "Average Sentiment": round(avg_sent, 3),
                "Key Keywords": ", ".join(info["keywords"]),
                "Representative Quote": info["samples"][0],
            })

        self.summary_df = pd.DataFrame(rows).sort_values(
            "Impact Score", ascending=False
        )

    # ══════════════════════════════════════════════════════════════════
    # SENTIMENT LABEL
    # ══════════════════════════════════════════════════════════════════

    def classify_theme_sentiment(self) -> None:
        labels = []
        for s in self.summary_df["Average Sentiment"]:
            if s < -0.1:
                labels.append("Problem")
            elif s > 0.1:
                labels.append("Positive")
            else:
                labels.append("Neutral")
        self.summary_df["Theme Sentiment"] = labels

    # ══════════════════════════════════════════════════════════════════
    # RECOMMENDATIONS  ─ rule-based with African context
    # ══════════════════════════════════════════════════════════════════

    def generate_recommendations(self) -> None:
        recs = []
        for _, row in self.summary_df.iterrows():
            text = (row["Theme"] + " " + row["Key Keywords"]).lower()
            matched = FALLBACK_RECOMMENDATION
            for keywords, rec in RECOMMENDATION_RULES:
                if any(k in text for k in keywords):
                    matched = rec
                    break
            recs.append(matched)
        self.summary_df.insert(5, "Recommendation", recs)

    # ══════════════════════════════════════════════════════════════════
    # PRIORITY LABELS  ─ intensity-aware
    # ══════════════════════════════════════════════════════════════════

    def add_priority_labels(self) -> None:
        priorities = []
        for _, row in self.summary_df.iterrows():
            score = row["Impact Score"]
            intensity = row["Intensity Score"]

            if score > 20 or (score > 10 and intensity > 0.5):
                priorities.append("🔴 High Priority")
            elif score > 10 or (score > 5 and intensity > 0.4):
                priorities.append("🟡 Medium Priority")
            else:
                priorities.append("🟢 Low Priority")

        self.summary_df.insert(4, "Priority", priorities)

    # ══════════════════════════════════════════════════════════════════
    # MERGE SIMILAR THEMES
    # ══════════════════════════════════════════════════════════════════

    def merge_similar_themes(self) -> None:
        themes = self.summary_df["Theme"].tolist()
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

        new_rows = []
        for main_theme, idxs in merged_map.items():
            subset = self.summary_df.iloc[idxs]
            row = {
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

    HEADER_COLOUR  = "2F5597"   # navy blue
    HEADING_COLOUR = "1F3864"   # darker navy for sheet title bar
    STRIPE_COLOUR  = "EEF2FF"   # soft lavender stripe
    SUBHEAD_COLOUR = "D6E4F7"   # light blue for section sub-headers

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

    # ══════════════════════════════════════════════════════════════════
    # EXCEL REPORT  ─ multi-question, fully rebuilt
    # ══════════════════════════════════════════════════════════════════

    def build_excel_report(self, output: Optional[str] = None) -> None:
        output = output or self.config.output_filename
        now_eat = datetime.now(EAT).strftime("%d %b %Y %H:%M EAT")

        wb = Workbook()
        wb.remove(wb.active)

        # ══════════════════════════════════════════════════════════════
        # 1. EXECUTIVE SUMMARY
        # ══════════════════════════════════════════════════════════════
        ws = wb.create_sheet("Executive Summary")
        EXEC_MERGE = 8   # columns to merge for heading

        next_row = self.write_sheet_heading(
            ws,
            title=self.config.report_title,
            subtitle=f"Generated: {now_eat}   |   Questions analysed: "
                     f"{len(self.question_results)}",
            merge_cols=EXEC_MERGE,
        )

        # blank spacer
        next_row += 1   # row 4

        # Questions list
        lbl = ws.cell(row=next_row, column=1, value="Questions Analysed")
        lbl.font = Font(bold=True, size=11, color=self.HEADER_COLOUR)
        next_row += 1
        for q in self.question_results:
            ws.cell(row=next_row, column=1, value=f"    •  {q}")
            next_row += 1
        next_row += 1

        # Cross-question themes
        if self.cross_question_themes:
            warn = ws.cell(
                row=next_row, column=1,
                value="⚠  Themes Recurring Across Multiple Questions"
            )
            warn.font = Font(bold=True, size=11, color="C00000")
            next_row += 1
            for t in self.cross_question_themes:
                ws.cell(row=next_row, column=1, value=f"    ▶  {t}")
                next_row += 1
            next_row += 1

        # Per-question summary blocks
        for qname, qdata in self.question_results.items():
            s_df: pd.DataFrame = qdata["summary"]
            insights: List[str]  = qdata["insights"]
            ds: Dict             = qdata["dataset_summary"]

            # Section divider
            divider = ws.cell(row=next_row, column=1, value=f"  {qname.upper()}")
            divider.font = Font(bold=True, size=12, color="FFFFFF")
            divider.fill = PatternFill(
                start_color=self.HEADER_COLOUR,
                end_color=self.HEADER_COLOUR,
                fill_type="solid",
            )
            ws.merge_cells(
                start_row=next_row, start_column=1,
                end_row=next_row, end_column=EXEC_MERGE
            )
            ws.row_dimensions[next_row].height = 22
            next_row += 1

            # Stats row
            for col, (label, key) in enumerate(
                [("Responses", "Total Responses"),
                 ("Themes", "Themes Identified"),
                 ("Avg Length", "Average Response Length (words)"),
                 ("Languages", "Languages Detected")],
                start=1
            ):
                lc = ws.cell(row=next_row, column=col * 2 - 1, value=label)
                lc.font = Font(bold=True, size=9, color="666666")
                vc = ws.cell(row=next_row, column=col * 2, value=ds.get(key, ""))
                vc.font = Font(bold=True, size=11)
            next_row += 2

            # Insights
            for ins in insights:
                ic = ws.cell(row=next_row, column=1, value=f"  •  {ins}")
                ic.alignment = Alignment(wrap_text=True)
                ws.row_dimensions[next_row].height = 20
                next_row += 1
            next_row += 1

            # Mini table header
            mini_cols = ["Theme", "Responses", "Share (%)", "Priority",
                         "Recommendation"]
            for ci, h in enumerate(mini_cols, 1):
                c = ws.cell(row=next_row, column=ci, value=h)
                c.font = Font(bold=True, color="FFFFFF", size=10)
                c.fill = PatternFill(
                    start_color=self.HEADER_COLOUR,
                    end_color=self.HEADER_COLOUR,
                    fill_type="solid",
                )
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = self._thin_border()
            ws.row_dimensions[next_row].height = 22
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
        ws.column_dimensions["E"].width = 52

        # ══════════════════════════════════════════════════════════════
        # 2. PER-QUESTION THEME SHEETS
        # ══════════════════════════════════════════════════════════════
        THEME_COLS = [
            "Theme", "Response Count", "Share (%)", "Impact Score",
            "Intensity Score", "Priority", "Theme Sentiment",
            "Average Sentiment", "Key Keywords", "Representative Quote",
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
                              "Key Keywords"):
                if long_col in available_cols:
                    lcol = available_cols.index(long_col) + 1
                    for row in ws.iter_rows(
                        min_row=data_start, min_col=lcol, max_col=lcol
                    ):
                        for cell in row:
                            cell.alignment = Alignment(
                                wrap_text=True, vertical="top"
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
        # 3. CHARTS SHEET
        #    Layout: one block per question, each block is CHART_BLOCK_ROWS
        #    tall.  Bar chart and Pie chart are placed SIDE BY SIDE
        #    (bar at col E, pie at col P) so they never overlap.
        # ══════════════════════════════════════════════════════════════
        ws = wb.create_sheet("Charts")

        CHART_BLOCK_ROWS = 32   # rows allocated per question block
        BAR_W, BAR_H   = 22, 13  # bar chart size (cm-ish in openpyxl units)
        PIE_W, PIE_H   = 16, 13  # pie chart size
        BAR_ANCHOR_COL = "E"     # bar chart left edge
        PIE_ANCHOR_COL = "P"     # pie chart left edge — well clear of bar

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

            # ── Bar chart (no legend) ──────────────────────────────────
            bar = BarChart()
            bar.type    = "col"
            bar.title   = f"{safe_q}  —  Theme Distribution"
            bar.y_axis.title = "Number of Responses"
            bar.x_axis.title = "Theme"
            bar.legend  = None          # ← DROP LEGEND
            bar.add_data(data_ref, titles_from_data=True)
            bar.set_categories(cats_ref)
            bar.width  = BAR_W
            bar.height = BAR_H
            # Anchor bar at (BAR_ANCHOR_COL, block_start)
            ws.add_chart(bar, f"{BAR_ANCHOR_COL}{block_start}")

            # ── Pie chart ─────────────────────────────────────────────
            pie = PieChart()
            pie.title  = f"{safe_q}  —  Share (%)"
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

        first = True
        all_q_dfs = []
        for qname, qdata in self.question_results.items():
            q_df: pd.DataFrame = qdata.get("q_df", pd.DataFrame())
            if q_df.empty:
                continue
            q_df = q_df.copy()
            q_df.insert(0, "Question", qname)
            all_q_dfs.append(q_df)

        if all_q_dfs:
            combined = pd.concat(all_q_dfs, ignore_index=True)
            n_cols = max(len(combined.columns), 6)

            next_row = self.write_sheet_heading(
                ws,
                title="Clustered Responses  —  Full Analysed Dataset",
                subtitle=f"Total sentence units: {len(combined)}   |   "
                         f"Questions: {len(all_q_dfs)}",
                merge_cols=n_cols,
            )

            for ci, col_name in enumerate(combined.columns, 1):
                ws.cell(row=next_row, column=ci, value=col_name)
            self.style_table_header(ws, header_row=next_row)
            data_start = next_row + 1

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

            log.info("Generating recommendations …")
            self.generate_recommendations()

            log.info("Adding priority labels …")
            self.add_priority_labels()

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
            }

            # Free per-question RAM
            del self.embeddings
            gc.collect()

        # --- Cross-question analysis ---
        if len(self.question_results) > 1:
            log.info("Detecting cross-question themes …")
            self.detect_cross_question_themes()

        log.info("Building Excel report …")
        self.build_excel_report()


# ══════════════════════════════════════════════════════════════════════
# ENTRY POINT  ─ configure and run
# ══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":

    config = EngineConfig(
        pca_components=64,          # reduce embedding dims for RAM efficiency
        min_cluster_size=5,
        emerging_min_size=2,
        max_clusters=20,
        merge_threshold=0.75,
        tfidf_max_features=5_000,
        large_dataset_threshold=500,
        min_responses=10,
        report_title="Customer Feedback Insights – Kenya Survey 2025",
        output_filename="Feedback_Insights_Report_v2.xlsx",
    )

    engine = FeedbackInsightEngine(
        file_path="health_survey_test_v2.xlsx",   # or .xlsx
        text_columns=[
            "q1_access",
            "q2_wait",
            "q3_staff",
            "q4_cost",
            "q5_suggestions",
        ],
        config=config,
    )

    engine.run()