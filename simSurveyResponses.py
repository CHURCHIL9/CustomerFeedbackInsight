"""
========================================================================
Synthetic Dataset Generator  ─  v2 Pipeline Stress-Test Edition
========================================================================

Designed to exercise EVERY new feature in FeedbackInsightEngine v2:

  ✅  Swahili / Sheng responses          → language detection, extended stop-words, custom sentiment lexicon
  ✅  Code-switched sentences            → mixed English-Swahili in one response
  ✅  Hedged / sarcastic positives       → "not bad", "could be better", "I guess it's okay"
  ✅  Near-duplicate responses           → deduplication logic
  ✅  Emerging / weak-signal themes      → rare issues appear only 3-6 times
  ✅  Loud minority                      → few respondents with extreme anger language
  ✅  Cross-question theme recurrence    → same theme seeded across multiple questions
  ✅  Kenya-specific context             → M-Pesa, boda-boda, matatu, NHIF/SHIF, county, jua kali
  ✅  African-English fillers            → "kindly", "whereby", "humbly"
  ✅  Output as .xlsx                    → tests chardet + openpyxl loader

Run:
    pip install pandas openpyxl
    python generate_test_dataset.py
Outputs:
    health_survey_test_v2.xlsx   (500 rows × 6 columns)
========================================================================
"""

import random
import pandas as pd

random.seed(42)

# ══════════════════════════════════════════════════════════════════════
# 1.  CORE THEMES  ─ English (same as v1 baseline)
# ══════════════════════════════════════════════════════════════════════

THEMES_EN = {
    "waiting_time": [
        "I waited for many hours before being attended to",
        "The queue was too long at the clinic",
        "We spent the whole day waiting for treatment",
        "Service was very slow and disorganized",
        "Too much delay before seeing a doctor",
        "Patients are kept waiting without any explanation",
        "The appointment system does not work properly",
    ],
    "high_cost": [
        "The consultation fee is too expensive",
        "I could not afford the treatment cost",
        "Medicine prices are very high here",
        "Healthcare services are not affordable for ordinary people",
        "The charges are beyond my income level",
        "Even with NHIF the copay is still too much",
        "I had to borrow money just to see the doctor",
    ],
    "stock_outs": [
        "There were no medicines available at the dispensary",
        "The drugs were completely out of stock",
        "I was told to buy medicine at a private pharmacy",
        "The pharmacy shelves were empty when I visited",
        "Essential medicines were unavailable for weeks",
        "The facility keeps running out of basic supplies",
    ],
    "distance": [
        "The health center is very far from my village",
        "I have to walk long distances to reach the clinic",
        "Transport to the hospital is very difficult",
        "The facility is too far away from where I live",
        "It takes more than two hours to get to the hospital",
        "There are no matatus that go directly to the health center",
    ],
    "staff_attitude": [
        "The nurses were rude and dismissive",
        "Staff were not friendly at all",
        "The doctor did not listen to my concerns",
        "Healthcare workers were disrespectful to patients",
        "Poor customer service at the facility",
        "The receptionist shouted at me in front of others",
        "Staff ignored patients who were in pain",
    ],
    "equipment": [
        "The hospital lacks proper medical equipment",
        "The ultrasound machine was not working during my visit",
        "There are no diagnostic tools available at this facility",
        "The facility is poorly equipped for basic procedures",
        "Broken equipment causes unnecessary delays in treatment",
        "They do not have a functioning X-ray machine",
    ],
    "transport": [
        "There is no reliable transport to reach the clinic",
        "Public transport to the hospital is very expensive",
        "Bad roads make it extremely hard to reach the hospital",
        "Transport costs are too high especially for boda-boda",
        "Ambulances are not available when you need them most",
        "The matatu stops far from the hospital gate",
    ],
    "positive": [
        "The service was excellent and I am satisfied",
        "Staff were helpful and very professional",
        "The clinic attended to me quickly and efficiently",
        "Doctors treated me well and explained everything clearly",
        "I received very good care during my stay",
        "The nurses were kind and patient with me",
        "Overall the facility has improved a lot recently",
    ],
}

# ══════════════════════════════════════════════════════════════════════
# 2.  SWAHILI RESPONSES  →  tests language detection + Swahili lexicon
# ══════════════════════════════════════════════════════════════════════

THEMES_SW = {
    "waiting_time": [
        "Nilipiga foleni muda mrefu sana bila msaada wowote",
        "Tulisubiri masaa mengi kabla ya kupata daktari",
        "Huduma ni polepole sana na haipendezi",
        "Watu wanang'ang'ania kupata nambari tangu asubuhi",
    ],
    "high_cost": [
        "Bei ya dawa ni ghali sana huwezi kumudu",
        "Hawakubali NHIF na unatakiwa kulipa pesa taslimu",
        "Gharama ya matibabu imezidi uwezo wangu",
        "Kulipa kupitia M-Pesa pia wanakuchaji ziada",
    ],
    "staff_attitude": [
        "Muuguzi alikuwa mkali sana na hakusikiliza",
        "Wahudumu wa afya hawana heshima kwa wagonjwa",
        "Daktari aliniambia niende haraka bila kueleza vizuri",
        "Nilihisi vibaya sana baada ya jinsi walivyonishughulikia",
    ],
    "positive": [
        "Huduma ilikuwa nzuri sana na nilifurahi",
        "Wafanyakazi walisaidia vizuri na walikuwa na subira",
        "Nimeridhika na matibabu niliyopata",
        "Kliniki ilinisaidia haraka na kwa upole",
    ],
    "distance": [
        "Kituo cha afya kiko mbali sana na simu yangu haina data",
        "Barabara mbaya inafanya usafiri kuwa mgumu",
        "Hakuna matatu yanayokwenda moja kwa moja hospitali",
    ],
}

# ══════════════════════════════════════════════════════════════════════
# 3.  CODE-SWITCHED RESPONSES (Sheng / mixed)
#     →  tests the language-aware cleaner; these should NOT be stripped
# ══════════════════════════════════════════════════════════════════════

CODE_SWITCHED = [
    "Maze the queue ilikuwa ndefu sana, nilichoka kustand",
    "Dawa ziko out of stock, wanakuambia ununue nje — si poa hii",
    "Staff wako rada sana, hawakujali hata kidogo",
    "Bana the hospital is too far, na boda-boda fare ni ghali mno",
    "Nimechoka kusubiri, stress mingi bila reason yoyote",
    "The charges ziko juu sana, sina chapaa kiasi hicho",
    "Walisema M-Pesa inafanya kazi but machine ilikuwa down",
    "Manze equipment yao iko fala, machine ya X-ray haifanyi kazi",
    "Sawa sawa but service ingeweza kuwa better honestly",
    "Mimi niko poa na huduma, wafanyakazi walikuwa safi",
    "Unapata dawa nzuri lakini bei inakuua mfano",
    "Transport hadi hapa ni stress, matatu hazifiki hapa usiku",
]

# ══════════════════════════════════════════════════════════════════════
# 4.  HEDGED / SARCASTIC POSITIVES
#     →  tests detect_hedged_sentiment(); these should score negative
# ══════════════════════════════════════════════════════════════════════

HEDGED = [
    "The service was not bad I suppose, but could be better",
    "I guess the waiting time was okay, nothing special really",
    "Not the best experience but not terrible either",
    "The staff were sort of okay, could be more professional",
    "It was acceptable I suppose but I expected more",
    "So so experience, the facility is average at best",
    "Meh, the service was not great but not awful",
    "Nothing special about this facility if I am being honest",
    "The doctor was okay I guess, did not really explain much",
    "Below average service but at least they attended to me",
    "I would not call it good but could have been worse",
    "The facility is okay in some ways but quite lacking overall",
]

# ══════════════════════════════════════════════════════════════════════
# 5.  NEAR-DUPLICATES
#     →  tests deduplicate_responses(); 20 rows repeat with minor variation
# ══════════════════════════════════════════════════════════════════════

NEAR_DUPLICATES = [
    "The queue was very long and I waited too long",
    "The queue was too long and I waited for many hours",
    "The queue was long and the waiting was too much",
    "I waited in a very long queue for many hours",
    "Long queue, too much waiting, very frustrating",
]

# ══════════════════════════════════════════════════════════════════════
# 6.  LOUD MINORITY  (very high-intensity negative language)
#     →  tests Intensity Score vs Impact Score separation
#     Only ~10 respondents but very strong language
# ══════════════════════════════════════════════════════════════════════

LOUD_MINORITY = [
    "This facility is absolutely terrible and I am disgusted by the treatment",
    "I have never been so disrespected and humiliated in my entire life at a clinic",
    "The staff behavior was completely unacceptable and should be reported",
    "I am outraged by the complete lack of care and professionalism here",
    "This is the worst healthcare experience I have ever had, deeply frustrated",
    "The negligence here is criminal, patients are suffering while staff do nothing",
    "I am beyond furious with the disrespectful attitude of every single staff member",
    "Appalling conditions, broken equipment, zero medicine — this is shameful",
    "I was humiliated, overcharged and ignored. This place is a disgrace",
    "Deeply disappointed and angry. My family member suffered because of this facility",
]

# ══════════════════════════════════════════════════════════════════════
# 7.  EMERGING / WEAK-SIGNAL THEMES  (only 4-6 occurrences in dataset)
#     →  tests ⚠ Emerging Issues sheet
# ══════════════════════════════════════════════════════════════════════

EMERGING = [
    # Mental health — rare but important signal
    "I feel like my mental health concerns are completely ignored here",
    "There is no counselling service available at this facility",
    "Depression and anxiety are not taken seriously by the staff",
    # Digital / USSD / app issues — growing concern
    "The online booking system keeps crashing on my phone",
    "The facility portal does not work properly on Android",
    "USSD appointment booking failed three times before I gave up",
    # Water / sanitation — infrastructure
    "The toilets at the facility are in a terrible condition",
    "There is no running water in the patient wards",
]

# ══════════════════════════════════════════════════════════════════════
# 8.  AFRICAN-ENGLISH FILLER RESPONSES
#     →  tests COMBINED_STOP_WORDS filtering
# ══════════════════════════════════════════════════════════════════════

AFRICAN_ENGLISH_FILLERS = [
    "Kindly the waiting time should be improved whereby patients do not suffer",
    "Humbly I request that the staff be trained to treat us with respect sincerely",
    "Whereby the transport is bad, the government should intervene accordingly",
    "Basically the charges are too much and the services are not up to standard",
    "Truly the medicine shortage is a serious issue that needs to be addressed",
    "Obviously the queue management needs to be improved going forward",
]

# ══════════════════════════════════════════════════════════════════════
# 9.  KENYA-SPECIFIC CONTEXT PHRASES
#     →  tests expanded recommendation engine
# ══════════════════════════════════════════════════════════════════════

KENYA_SPECIFIC = [
    "NHIF card was rejected and I had to pay out of pocket",
    "SHIF registration is confusing and nobody explains how it works",
    "The M-Pesa payment option was down and I had no cash",
    "Boda-boda from my village to the clinic costs three hundred shillings",
    "County government should build a dispensary closer to our ward",
    "Matatu fare to the hospital takes most of my daily wage",
    "Jua kali workers like me cannot afford to miss a whole day for treatment",
    "The facility is in the next county and crossing the border is expensive",
    "NHIF should cover more services at the Level 2 dispensary near us",
    "Mobile clinic visits used to come monthly but stopped last year",
]

# ══════════════════════════════════════════════════════════════════════
# QUESTION DEFINITIONS
# ══════════════════════════════════════════════════════════════════════

QUESTIONS = [
    "q1_access",
    "q2_wait",
    "q3_staff",
    "q4_cost",
    "q5_suggestions",
]

# Theme-to-question affinity — makes cross-question detection realistic
# (same theme recurs but naturally across different questions)
QUESTION_THEME_AFFINITY = {
    "q1_access":      ["distance", "transport", "high_cost", "stock_outs"],
    "q2_wait":        ["waiting_time", "equipment", "staff_attitude"],
    "q3_staff":       ["staff_attitude", "positive", "waiting_time"],
    "q4_cost":        ["high_cost", "stock_outs", "transport"],
    "q5_suggestions": ["waiting_time", "high_cost", "distance", "positive",
                       "staff_attitude", "equipment"],
}

# ══════════════════════════════════════════════════════════════════════
# RESPONSE PICKER  ─ question-aware
# ══════════════════════════════════════════════════════════════════════

def pick_response(
    question: str,
    respondent_themes: list,
    use_swahili: bool,
    use_code_switch: bool,
    use_hedged: bool,
    use_kenya: bool,
    use_filler: bool,
) -> str:

    # Override: special response types
    if use_hedged and random.random() < 0.7:
        return random.choice(HEDGED)

    if use_code_switch and random.random() < 0.8:
        return random.choice(CODE_SWITCHED)

    if use_kenya and random.random() < 0.7:
        return random.choice(KENYA_SPECIFIC)

    if use_filler and random.random() < 0.6:
        return random.choice(AFRICAN_ENGLISH_FILLERS)

    if use_swahili:
        # Pick a Swahili theme that fits this question if possible
        affinity = QUESTION_THEME_AFFINITY.get(question, list(THEMES_SW.keys()))
        sw_themes = [t for t in respondent_themes if t in THEMES_SW and t in affinity]
        if sw_themes:
            return random.choice(THEMES_SW[random.choice(sw_themes)])
        # fallback to any Swahili theme
        t = random.choice(list(THEMES_SW.keys()))
        return random.choice(THEMES_SW[t])

    # Standard English — prefer affinity themes
    affinity = QUESTION_THEME_AFFINITY.get(question, list(THEMES_EN.keys()))
    candidate_themes = [t for t in respondent_themes if t in affinity]
    if not candidate_themes:
        candidate_themes = respondent_themes

    theme = random.choice(candidate_themes)
    return random.choice(THEMES_EN[theme])


# ══════════════════════════════════════════════════════════════════════
# GENERATE DATASET
# ══════════════════════════════════════════════════════════════════════

def generate_dataset(n: int = 500, output: str = "health_survey_test_v2.xlsx") -> pd.DataFrame:

    all_themes = list(THEMES_EN.keys())
    rows = []

    # --- Special respondent pools (indices) ---
    n_swahili      = int(n * 0.12)   # 12%  pure Swahili
    n_code_switch  = int(n * 0.10)   # 10%  Sheng / code-switched
    n_hedged       = int(n * 0.08)   # 8%   hedged / sarcastic
    n_near_dup     = int(n * 0.04)   # 4%   near-duplicate queue complaints
    n_loud         = 10              # fixed: loud minority (extreme anger)
    n_kenya        = int(n * 0.08)   # 8%   Kenya-specific context phrases
    n_filler       = int(n * 0.04)   # 4%   African-English filler language

    # Assign special types to respondent indices (non-overlapping)
    idx = list(range(n))
    random.shuffle(idx)

    swahili_idx    = set(idx[:n_swahili])
    idx            = idx[n_swahili:]
    code_idx       = set(idx[:n_code_switch])
    idx            = idx[n_code_switch:]
    hedged_idx     = set(idx[:n_hedged])
    idx            = idx[n_hedged:]
    near_dup_idx   = set(idx[:n_near_dup])
    idx            = idx[n_near_dup:]
    loud_idx       = set(idx[:n_loud])
    idx            = idx[n_loud:]
    kenya_idx      = set(idx[:n_kenya])
    idx            = idx[n_kenya:]
    filler_idx     = set(idx[:n_filler])

    for i in range(n):
        # Assign 1–3 base themes per respondent
        respondent_themes = random.sample(all_themes, random.randint(1, 3))

        # Respondent type flags
        is_swahili     = i in swahili_idx
        is_code_switch = i in code_idx
        is_hedged      = i in hedged_idx
        is_near_dup    = i in near_dup_idx
        is_loud        = i in loud_idx
        is_kenya       = i in kenya_idx
        is_filler      = i in filler_idx

        row = {"response_id": i + 1}

        for q in QUESTIONS:
            # Loud minority: extreme language on staff and access questions
            if is_loud and q in ("q3_staff", "q1_access"):
                row[q] = random.choice(LOUD_MINORITY)
                continue

            # Near-duplicates: always queue complaint
            if is_near_dup:
                row[q] = random.choice(NEAR_DUPLICATES)
                continue

            row[q] = pick_response(
                question=q,
                respondent_themes=respondent_themes,
                use_swahili=is_swahili,
                use_code_switch=is_code_switch,
                use_hedged=is_hedged,
                use_kenya=is_kenya,
                use_filler=is_filler,
            )

        rows.append(row)

    # --- Inject emerging / weak-signal themes ---
    # Sprinkle 2-4 times each into q5_suggestions (so they appear but rarely)
    emerging_targets = random.sample(range(n), min(len(EMERGING) * 3, n))
    for j, target_i in enumerate(emerging_targets):
        rows[target_i]["q5_suggestions"] = EMERGING[j % len(EMERGING)]

    df = pd.DataFrame(rows)

    # --- Save as .xlsx (tests chardet + openpyxl loader in v2 engine) ---
    df.to_excel(output, index=False)
    print(f"✅ Dataset saved → {output}")
    print(f"   Shape: {df.shape}")
    print()

    # --- Print respondent type breakdown ---
    print("Respondent composition:")
    print(f"  Standard English      : {n - n_swahili - n_code_switch - n_hedged - n_near_dup - n_loud - n_kenya - n_filler}")
    print(f"  Swahili responses     : {n_swahili}  ({n_swahili/n*100:.0f}%)")
    print(f"  Code-switched (Sheng) : {n_code_switch}  ({n_code_switch/n*100:.0f}%)")
    print(f"  Hedged / sarcastic    : {n_hedged}  ({n_hedged/n*100:.0f}%)")
    print(f"  Near-duplicate        : {n_near_dup}  ({n_near_dup/n*100:.0f}%)")
    print(f"  Loud minority         : {n_loud}   ({n_loud/n*100:.0f}%)")
    print(f"  Kenya-specific ctx    : {n_kenya}  ({n_kenya/n*100:.0f}%)")
    print(f"  African-English filler: {n_filler}  ({n_filler/n*100:.0f}%)")
    print(f"  Emerging issues       : {min(len(EMERGING)*3, n)} responses")
    print()
    print("Features this tests in v2 pipeline:")
    print("  ✅ Language detection (sw / en / mixed)")
    print("  ✅ Swahili stop-words + custom sentiment lexicon")
    print("  ✅ Hedged/sarcastic sentiment down-scoring")
    print("  ✅ Near-duplicate deduplication")
    print("  ✅ Emerging issues (⚠ sheet)")
    print("  ✅ Loud minority  (Intensity Score vs Impact Score)")
    print("  ✅ Cross-question theme recurrence (Executive Summary)")
    print("  ✅ Kenya-specific recommendations (M-Pesa, NHIF, boda-boda)")
    print("  ✅ African-English filler word filtering")
    print("  ✅ .xlsx input (chardet + openpyxl loader)")
    print()
    print(df.head(8).to_string(index=False))

    return df


# ══════════════════════════════════════════════════════════════════════
# RUN
# ══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    df = generate_dataset(n=500, output="health_survey_test_v2.xlsx")