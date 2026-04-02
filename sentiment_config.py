"""
========================================================================
sentiment_config.py  —  Complaint-Signal Scoring Layer
FeedbackInsightEngine — Phase 1 Sentiment Accuracy Upgrade
========================================================================

PURPOSE
-------
This file implements a complaint-signal pre-scoring layer that sits
BEFORE VADER in the sentiment pipeline.  It addresses the core failure
mode: VADER defaults to 0.0 (neutral/positive) for any word it does
not know, which means Swahili sentences and domain-specific English
complaints both score falsely positive.

HOW IT WORKS
------------
Three independent signal detectors run on every sentence:

  1. complaint_signal_score()
     Detects structural signals of dissatisfaction — contrastive
     conjunctions ("but", "however", "lakini"), unmet expectation
     phrases ("was supposed to", "never came"), and complaint verbs
     ("delayed", "failed", "missed").  These work regardless of
     whether the surrounding words are in VADER's dictionary.

  2. domain_lexicon_score()
     A domain-specific word → polarity dictionary covering agricultural
     and NGO programme vocabulary that VADER scores as 0.  Words like
     "delayed", "rotten", "penalised", "infrequent" are complaints in
     this context and are scored accordingly.

  3. negation_score()
     An extended negation window (up to 8 tokens, vs VADER's 3) that
     catches constructions like "did not germinate well after planting"
     where the negation and the scored word are far apart.

The three scores are combined into a single COMPLAINT ADJUSTMENT
(always ≤ 0.0) that is added to the VADER compound score BEFORE
language-weighted blending.  The adjustment is clamped so it cannot
push a genuinely positive sentence into negative territory by more
than a defined maximum.

INTEGRATION
-----------
Import this module in qualitativeInsightAnalysis.py and call:

    from sentiment_config import compute_complaint_adjustment

Then in compute_sentiment(), replace:

    hedge_adj   = self.detect_hedged_sentiment(text)
    vader_score = max(-1.0, min(1.0, base + hedge_adj))

with:

    hedge_adj       = self.detect_hedged_sentiment(text)
    complaint_adj   = compute_complaint_adjustment(text, lang)
    vader_score     = max(-1.0, min(1.0, base + hedge_adj + complaint_adj))

CUSTOMIZATION
-------------
All signal lists and weights are defined as module-level constants.
To extend for a new sector (healthcare, WASH, education), add words
to the relevant lists below — no changes to the pipeline needed.

To DISABLE a specific detector without editing the pipeline, set its
weight to 0.0 in SIGNAL_WEIGHTS at the bottom of this file.

========================================================================
"""

import re
from typing import List, Tuple

# ══════════════════════════════════════════════════════════════════════
# 1.  CONTRASTIVE CONJUNCTIONS
#     These words signal that the second half of a sentence contradicts
#     or undermines the first half.  Even if the first clause is
#     positive, the presence of these words shifts overall sentiment
#     toward negative.
#     Weight applied: -0.30 per match (capped at -0.50 total)
# ══════════════════════════════════════════════════════════════════════

CONTRASTIVE_EN: List[str] = [
    "but", "however", "though", "although", "yet", "except",
    "unless", "despite", "still", "even though", "even so",
    "on the other hand", "that said", "nevertheless", "nonetheless",
    "in spite of", "regardless", "unfortunately", "sadly",
    "regrettably", "only that", "the problem is", "the issue is",
    "the challenge is",
]

CONTRASTIVE_SW: List[str] = [
    "lakini", "ingawa", "hata hivyo", "ijapokuwa", "bado",
    "isipokuwa", "hata kama", "pamoja na hayo", "kwa bahati mbaya",
    "tatizo ni", "changamoto ni", "shida ni",
]

# Combined and pre-compiled for speed
_CONTRASTIVE_PATTERN = re.compile(
    r"\b(" +
    "|".join(re.escape(p) for p in sorted(
        CONTRASTIVE_EN + CONTRASTIVE_SW, key=len, reverse=True
    )) +
    r")\b",
    re.IGNORECASE,
)

# ══════════════════════════════════════════════════════════════════════
# 2.  UNMET EXPECTATION PHRASES
#     These signal that something promised, expected, or required did
#     not happen.  Structurally a complaint even if no negative word
#     appears.
#     Weight applied: -0.35 per match (capped at -0.60 total)
# ══════════════════════════════════════════════════════════════════════

UNMET_EXPECTATION_EN: List[str] = [
    # Negated auxiliary verbs — the most common English complaint form
    "did not", "didn't", "could not", "couldn't", "was not", "wasn't",
    "were not", "weren't", "has not", "hasn't", "have not", "haven't",
    "had not", "hadn't", "would not", "wouldn't", "should not",
    "shouldn't", "does not", "doesn't", "do not", "don't",
    "cannot", "can't", "will not", "won't",
    # Expectation / promise violation
    "supposed to", "should have", "was expected", "were expected",
    "promised", "never came", "never arrived", "never received",
    "failed to", "unable to", "not able to", "too late",
    "no longer", "no support", "no help", "no response",
    "left us", "abandoned", "ignored",
    # Scarcity / insufficiency
    "not enough", "too little", "too few", "too small", "not sufficient",
    "ran out", "no stock", "out of stock", "missing",
]

UNMET_EXPECTATION_SW: List[str] = [
    # Negated Swahili auxiliary patterns
    "haikuja", "haikufika", "haikutosha", "hawakuja", "hawakufika",
    "sikupata", "hatukupata", "sikuweza", "hatukuweza",
    "haijakamilika", "haijafika", "haijawahi",
    # Expectation / scarcity
    "haikutimia", "haikufaulu", "haikutekelezwa",
    "hakuna msaada", "hakuna jibu", "hakuna mtu",
    "si ya kutosha", "haitoshi", "pungufu",
    # Abandonment
    "walituacha", "hawakusaidia", "hawakujibu",
    # Non-arrival / absence
    "hawafiki", "hafiki", "hawaji", "hakufika", "hakukuja",
    "haikuja", "haikufika", "hawaendi", "hawaonekani",
]

_UNMET_PATTERN = re.compile(
    r"\b(" +
    "|".join(re.escape(p) for p in sorted(
        UNMET_EXPECTATION_EN + UNMET_EXPECTATION_SW, key=len, reverse=True
    )) +
    r")\b",
    re.IGNORECASE,
)

# ══════════════════════════════════════════════════════════════════════
# 3.  COMPLAINT VERBS & PROBLEM NOUNS
#     Action words and nouns that describe a problem situation.  VADER
#     scores many of these as 0 (neutral) because they are not
#     inherently negative in all contexts, but in survey feedback about
#     service delivery they are virtually always complaints.
#     Weight applied: -0.20 per match (capped at -0.45 total)
# ══════════════════════════════════════════════════════════════════════

COMPLAINT_VERBS_EN: List[str] = [
    "delayed", "delay", "delays", "delaying",
    "missed", "miss", "missing",
    "failed", "fail", "failing", "failure",
    "struggled", "struggle", "struggling",
    "waited", "wait", "waiting",
    "lost", "lose", "losing",
    "affected", "affect",
    "forced", "force",
    "suffered", "suffer", "suffering",
    "complained", "complain", "complaining",
    "penalised", "penalized", "penalise",
    "punished", "punish",
    "charged", "overcharged",
    "rejected", "reject",
    "ignored", "ignore",
    "abandoned", "abandon",
    "disappointed", "disappoint",
    "frustrated", "frustrate",
    "worried", "worry",
    "confused", "confuse",
    "cheated", "cheat",
    "lied", "lie",
]

COMPLAINT_NOUNS_EN: List[str] = [
    "problem", "problems",
    "issue", "issues",
    "challenge", "challenges",
    "complaint", "complaints",
    "delay", "delays",
    "failure", "failures",
    "shortage", "shortages",
    "penalty", "penalties",
    "burden", "burdens",
    "obstacle", "obstacles",
    "barrier", "barriers",
    "loss", "losses",
    "damage", "damages",
    "mistake", "mistakes",
    "error", "errors",
    "difficulty", "difficulties",
    "inconvenience",
    "frustration",
    "disappointment",
    "confusion",
    "conflict", "conflicts",
]

COMPLAINT_SW: List[str] = [
    # Verbs
    "chelewa", "chukuliwa", "poteza", "shinda", "lalamika",
    "adhibiwa", "onewa", "dhulumiwa", "kamatiwa",
    # Nouns
    "tatizo", "matatizo", "shida", "changamoto", "lalamiko",
    "malalamiko", "hasara", "hasara", "adhabu", "kikwazo",
    "ugumu", "msongo", "wasiwasi", "hofu", "onyo",
    # Difficulty / access
    "difficult", "difficulty", "hard", "harder",
    "unable", "inaccessible",
    # Sheng
    "stress", "noma", "ngori",
]

_COMPLAINT_PATTERN = re.compile(
    r"\b(" +
    "|".join(re.escape(p) for p in sorted(
        COMPLAINT_VERBS_EN + COMPLAINT_NOUNS_EN + COMPLAINT_SW,
        key=len, reverse=True
    )) +
    r")\b",
    re.IGNORECASE,
)

# ══════════════════════════════════════════════════════════════════════
# 4.  DOMAIN-SPECIFIC NEGATIVE LEXICON
#     Words that are neutral in general English but clearly negative
#     in the context of agricultural programmes and NGO service delivery.
#     VADER scores these as 0.0 — this lexicon corrects that.
# ══════════════════════════════════════════════════════════════════════

DOMAIN_NEGATIVE_LEXICON: dict = {

    # ── Delivery / timing failures ────────────────────────────────────
    "late":          -0.50,
    "too late":      -0.60,
    "overdue":       -0.55,
    "delayed":       -0.55,
    "delay":         -0.50,
    "untimely":      -0.50,
    "missed":        -0.50,
    "miss":          -0.45,

    # ── Input / product quality ───────────────────────────────────────
    "rotten":        -0.70,
    "spoiled":       -0.65,
    "defective":     -0.65,
    "substandard":   -0.60,
    "poor quality":  -0.65,
    "germinate":     -0.20,   # almost always "did not germinate" in context
    "germination":   -0.20,
    "underperform":  -0.55,
    "inconsistent":  -0.45,

    # ── Access / distance ─────────────────────────────────────────────
    "far":           -0.35,
    "too far":       -0.50,
    "distant":       -0.40,
    "inaccessible":  -0.55,
    "remote":        -0.25,
    "travel far":    -0.50,

    # ── Staff / officer issues ────────────────────────────────────────
    "infrequent":    -0.45,
    "unreachable":   -0.55,
    "unavailable":   -0.50,
    "unresponsive":  -0.55,
    "unhelpful":     -0.55,
    "rude":          -0.65,
    "disrespectful": -0.70,
    "turnover":      -0.40,
    "replaced":      -0.25,   # only negative when staff replaced frequently

    # ── Loan / finance ────────────────────────────────────────────────
    "penalised":     -0.60,
    "penalized":     -0.60,
    "penalty":       -0.55,
    "defaulted":     -0.50,
    "default":       -0.45,
    "debt":          -0.45,
    "overcharged":   -0.60,
    "predatory":     -0.75,
    "trapped":       -0.60,
    "tight":         -0.35,
    "burden":        -0.50,

    # ── Group / community dysfunction ────────────────────────────────
    "conflict":      -0.55,
    "distrust":      -0.60,
    "absent":        -0.40,
    "absent members":-0.45,
    "defaulting":    -0.45,

    # ── Market / income ───────────────────────────────────────────────
    "below market":  -0.55,
    "low price":     -0.50,
    "no market":     -0.55,
    "no buyer":      -0.55,
    "unsupported":   -0.50,
    "abandoned":     -0.65,

    # ── Training / knowledge gaps ────────────────────────────────────
    "unclear":       -0.45,
    "confusing":     -0.45,
    "insufficient":  -0.50,
    "inadequate":    -0.50,
    "incomplete":    -0.45,
    "rushed":        -0.40,
    "too fast":      -0.40,
    "too short":     -0.35,

    # ── Emotional / strong negative ──────────────────────────────────
    "furious":       -0.80,
    "outraged":      -0.80,
    "betrayed":      -0.75,
    "humiliated":    -0.75,
    "appalling":     -0.75,
    "disaster":      -0.70,
    "ruined":        -0.70,
    "terrible":      -0.70,
    "awful":         -0.70,
    "dreadful":      -0.70,
    "worst":         -0.75,
    "useless":       -0.65,
    "pathetic":      -0.65,

    # ── Difficulty / access ──────────────────────────────────────────
    "difficult to reach":  -0.55,
    "hard to reach":       -0.55,
    "difficult to contact":-0.50,
    "hard to contact":     -0.50,
    "hard to find":        -0.45,
    "difficult to find":   -0.45,
    "no visit":            -0.50,
    "no visits":           -0.50,
    "rarely visits":       -0.45,
    "never visits":        -0.55,
    "does not visit":      -0.55,
    "changed":             -0.00,   # too ambiguous — removed from scoring

    # ── Swahili domain negatives ──────────────────────────────────────
    "hawafiki":    -0.55,   # they don't come / don't arrive
    "hafiki":      -0.50,   # he/she doesn't come
    "hawaji":      -0.50,   # they don't come
    "mbali":       -0.35,   # far / distant
    "mbali sana":  -0.50,   # very far
    "ngumu":       -0.50,   # difficult / hard
    "ngumu sana":  -0.60,   # very difficult
    "pungufu":     -0.45,   # insufficient / deficient
    "haitoshi":    -0.50,   # it is not enough
    "haikutosha":  -0.55,   # it was not enough
    "haikufika":   -0.50,   # it did not arrive
    "haukufika":   -0.50,   # it did not arrive (alt form)
    "haikuja":     -0.50,   # it did not come
    "hawakuja":    -0.55,   # they did not come
    "hakukuja":    -0.50,   # he/she did not come
    "hakufika":    -0.50,   # he/she did not arrive
    "imechelewa":  -0.50,   # it has been delayed
    "ilichelewa":  -0.55,   # it was delayed
    "walichelewa": -0.55,   # they were late
    "inachukua muda": -0.40, # it takes too long

    # ── Agricultural domain neutrals that imply loss in context ──────
    "drought":       -0.55,
    "flood":         -0.50,
    "flooding":      -0.50,
    "pests":         -0.45,
    "pest":          -0.40,
    "crop failure":  -0.70,
    "yield loss":    -0.65,
    "low yield":     -0.60,
    "poor harvest":  -0.65,
    "soil":          -0.05,   # mild — only in "poor soil" context
}

# ══════════════════════════════════════════════════════════════════════
# 5.  EXTENDED NEGATION WINDOW
#     VADER's negation window is only 3 tokens wide.
#     "Seeds did not germinate well after planting" — by the time VADER
#     reaches "well", the negation from "not" has already expired.
#     We extend this to 8 tokens and handle the most common patterns.
# ══════════════════════════════════════════════════════════════════════

# English negation trigger words
NEGATION_TRIGGERS_EN: List[str] = [
    "not", "no", "never", "neither", "nor", "without",
    "lack", "lacking", "lacks", "absent", "absence",
    "fail", "failed", "fails", "failure",
    "unable", "cannot", "can't", "couldn't", "wouldn't",
    "didn't", "doesn't", "don't", "won't", "wasn't", "weren't",
    "hasn't", "haven't", "hadn't",
]

# Swahili negation trigger words
NEGATION_TRIGGERS_SW: List[str] = [
    "si", "bila", "hapana", "kamwe", "la", "hata",
    "sikuweza", "hatukuweza", "hawakuweza",
    "haikuja", "haikufika", "haikutosha",
]

NEGATION_WINDOW_SIZE = 8   # tokens to look ahead after a negation trigger

# Words that are positive but become negative when negated
# (Only score if a negation token is within NEGATION_WINDOW_SIZE tokens)
NEGATABLE_POSITIVES: dict = {
    "good":       -0.50,
    "well":       -0.45,
    "great":      -0.55,
    "fine":       -0.40,
    "okay":       -0.35,
    "sufficient": -0.50,
    "enough":     -0.45,
    "arrive":     -0.50,   # "did not arrive"
    "arrived":    -0.50,
    "come":       -0.45,   # "never came"
    "came":       -0.45,
    "work":       -0.45,   # "did not work"
    "worked":     -0.45,
    "help":       -0.45,   # "did not help"
    "helped":     -0.45,
    "support":    -0.40,
    "supported":  -0.40,
    "receive":    -0.50,   # "did not receive"
    "received":   -0.50,
    "deliver":    -0.50,   # "did not deliver"
    "delivered":  -0.50,
    "understand": -0.40,
    "understood": -0.40,
    # Swahili positives that become negative when negated
    "nzuri":      -0.55,
    "vizuri":     -0.55,
    "tosha":      -0.50,
    "fika":       -0.50,
    "kuja":       -0.45,
    "saidia":     -0.45,
    "pata":       -0.45,
}

# ══════════════════════════════════════════════════════════════════════
# 6.  SIGNAL WEIGHTS
#     Controls how much each detector contributes to the final
#     complaint adjustment score.
#     Set a weight to 0.0 to disable that detector entirely.
# ══════════════════════════════════════════════════════════════════════

SIGNAL_WEIGHTS = {
    "contrastive":       0.30,   # weight per contrastive conjunction found
    "contrastive_cap":   0.50,   # maximum total from contrastive signals
    "unmet":             0.35,   # weight per unmet-expectation phrase found
    "unmet_cap":         0.65,   # maximum total from unmet-expectation signals
    "complaint":         0.20,   # weight per complaint verb/noun found
    "complaint_cap":     0.50,   # maximum total from complaint signals
    "domain":            1.00,   # domain lexicon scores are used directly
    "negation":          1.00,   # negation window scores are used directly
    "final_cap":         0.85,   # maximum total negative adjustment
}


# ══════════════════════════════════════════════════════════════════════
# PUBLIC API
# ══════════════════════════════════════════════════════════════════════

def complaint_signal_score(text: str, lang: str = "en") -> float:
    """
    Detect structural complaint signals and return a negative adjustment.

    Checks three signal types:
      1. Contrastive conjunctions  (but, however, lakini …)
      2. Unmet expectation phrases (did not, was supposed to, haikuja …)
      3. Complaint verbs / nouns   (delayed, failed, tatizo …)

    Args:
        text: cleaned, lowercased sentence text
        lang: detected language code ("sw", "en", or other)

    Returns:
        float ≤ 0.0  — negative adjustment to add to VADER score
    """
    text_lower = text.lower()
    total_adj = 0.0

    # Signal 1: Contrastive conjunctions
    contrastive_matches = _CONTRASTIVE_PATTERN.findall(text_lower)
    if contrastive_matches:
        adj = min(
            len(contrastive_matches) * SIGNAL_WEIGHTS["contrastive"],
            SIGNAL_WEIGHTS["contrastive_cap"],
        )
        total_adj += adj

    # Signal 2: Unmet expectation phrases
    unmet_matches = _UNMET_PATTERN.findall(text_lower)
    if unmet_matches:
        adj = min(
            len(unmet_matches) * SIGNAL_WEIGHTS["unmet"],
            SIGNAL_WEIGHTS["unmet_cap"],
        )
        total_adj += adj

    # Signal 3: Complaint verbs / nouns
    complaint_matches = _COMPLAINT_PATTERN.findall(text_lower)
    if complaint_matches:
        adj = min(
            len(complaint_matches) * SIGNAL_WEIGHTS["complaint"],
            SIGNAL_WEIGHTS["complaint_cap"],
        )
        total_adj += adj

    return -min(total_adj, SIGNAL_WEIGHTS["final_cap"])


def domain_lexicon_score(text: str) -> float:
    """
    Score text against the domain-specific negative lexicon.

    Checks multi-word phrases first (longest match wins), then single
    words.  Returns the mean of all matched scores, or 0.0 if no
    domain words are found.

    Args:
        text: cleaned, lowercased sentence text

    Returns:
        float in [-1.0, 0.0]  (domain lexicon only has negative entries)
    """
    text_lower = text.lower()
    matched_scores = []

    # Sort by length descending so multi-word phrases match before parts
    for phrase, score in sorted(
        DOMAIN_NEGATIVE_LEXICON.items(), key=lambda x: len(x[0]), reverse=True
    ):
        if re.search(r"\b" + re.escape(phrase) + r"\b", text_lower):
            matched_scores.append(score)
            # Remove the matched phrase so its parts don't double-count
            text_lower = re.sub(
                r"\b" + re.escape(phrase) + r"\b", " ", text_lower
            )

    if not matched_scores:
        return 0.0

    # Use the mean of matched scores — multiple complaints compound
    raw = sum(matched_scores) / len(matched_scores)
    return max(-1.0, raw)


def negation_window_score(text: str) -> float:
    """
    Detect negated positives using an extended 8-token window.

    VADER's negation window is only 3 tokens wide and misses patterns
    like "did not germinate well after planting" where the positive word
    is 5+ tokens after the negation trigger.

    Args:
        text: cleaned, lowercased sentence text

    Returns:
        float in [-1.0, 0.0]
    """
    tokens = text.lower().split()
    n = len(tokens)
    if n < 2:
        return 0.0

    all_negators = set(NEGATION_TRIGGERS_EN + NEGATION_TRIGGERS_SW)
    scores = []

    for i, tok in enumerate(tokens):
        clean_tok = tok.strip(".,!?;:'\"")
        if clean_tok in all_negators:
            # Look ahead up to NEGATION_WINDOW_SIZE tokens
            window_end = min(i + NEGATION_WINDOW_SIZE + 1, n)
            for j in range(i + 1, window_end):
                candidate = tokens[j].strip(".,!?;:'\"")
                if candidate in NEGATABLE_POSITIVES:
                    scores.append(NEGATABLE_POSITIVES[candidate])
                    break   # one negation trigger flips one positive

    if not scores:
        return 0.0

    raw = sum(scores) / len(scores)
    return max(-1.0, raw)


def compute_complaint_adjustment(text: str, lang: str = "en") -> float:
    """
    Master function — combines all three detectors into one adjustment.

    This is the single function imported by qualitativeInsightAnalysis.py.
    It replaces the single detect_hedged_sentiment() call with a
    comprehensive complaint-signal adjustment.

    Call order (additive):
      1. complaint_signal_score()   — structural complaint signals
      2. domain_lexicon_score()     — domain-specific negative words
      3. negation_window_score()    — extended negation window

    The final adjustment is clamped to [-final_cap, 0.0] so it cannot
    push an already-negative score further than the cap allows, and
    cannot make a genuinely positive score go deeply negative on its own.

    Args:
        text: cleaned sentence text (the clean_text column value)
        lang: detected language code ("sw", "en", or other)

    Returns:
        float in [-0.85, 0.0]  — always ≤ 0.0
    """
    signal_adj  = complaint_signal_score(text, lang)
    domain_adj  = domain_lexicon_score(text)
    negation_adj = negation_window_score(text)

    total = signal_adj + domain_adj + negation_adj
    return max(-SIGNAL_WEIGHTS["final_cap"], min(0.0, total))


# ══════════════════════════════════════════════════════════════════════
# QUICK SELF-TEST  (run:  python sentiment_config.py)
# ══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    TEST_CASES = [
        # (text, lang, expected_direction)
        # English complaints
        ("The inputs were delivered late and I missed the planting window", "en", "negative"),
        ("The field officer is difficult to reach on the phone between visits", "en", "negative"),
        ("I was penalised for late repayment even though the harvest was delayed", "en", "negative"),
        ("The seeds provided did not germinate well after planting", "en", "negative"),
        ("Training was good but delivery was late", "en", "negative"),
        ("I am absolutely furious with the program they have ruined my season", "en", "negative"),
        ("The loan system is predatory and designed to trap poor farmers in debt", "en", "negative"),
        # Swahili complaints
        ("Mbegu zilifika baada ya mvua kuanza na sikuweza kupanda kwa wakati", "sw", "negative"),
        ("Afisa wa shamba hawafiki kwa wakulima wa mbali", "sw", "negative"),
        ("Mbolea haikutosha kwa shamba langu lote", "sw", "negative"),
        ("Ratiba ya malipo ni ngumu sana", "sw", "negative"),
        # Genuine positives (should stay positive or low negative)
        ("My yield doubled after following One Acre Fund planting advice", "en", "positive"),
        ("One Acre Fund has truly changed my farming and my family income", "en", "positive"),
        ("The training on spacing and planting techniques improved my yield", "en", "positive"),
        ("Nilifurahi sana na huduma ya One Acre Fund msimu huu", "sw", "positive"),
    ]

    print("\n" + "═" * 72)
    print("  SENTIMENT CONFIG — SELF TEST")
    print("═" * 72)
    print(f"  {'Text':<55} {'Lang':<5} {'Adj':>7}  {'Expected'}")
    print("─" * 72)

    correct = 0
    for text, lang, expected in TEST_CASES:
        adj = compute_complaint_adjustment(text, lang)
        direction = "negative" if adj < -0.05 else "positive"
        ok = "✅" if direction == expected else "❌"
        short = text[:52] + "..." if len(text) > 52 else text
        print(f"  {ok} {short:<55} {lang:<5} {adj:>7.3f}  {expected}")
        if direction == expected:
            correct += 1

    print("─" * 72)
    print(f"  Passed: {correct}/{len(TEST_CASES)}")
    print("═" * 72 + "\n")