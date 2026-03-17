"""
========================================================================
Synthetic Dataset Generator  ─  One Acre Fund Demo Edition
========================================================================

Sector : Smallholder Agriculture  |  Organisation: One Acre Fund
Target : M&E Officer Demo

Covers the full One Acre Fund service model:
  • Seed & fertilizer input bundles
  • Field officer training & farm visits
  • Group loan & repayment
  • Market linkage & post-harvest support
  • Input delivery logistics
  • Mobile money (M-Pesa) payments

Stress-tests ALL v2 pipeline features:
  ✅  Swahili / Sheng responses          → language detection, extended stop-words, custom sentiment lexicon
  ✅  Code-switched sentences            → mixed English-Swahili / Sheng
  ✅  Hedged / sarcastic positives       → "not bad", "could be better", "I guess it's okay"
  ✅  Near-duplicate responses           → deduplication logic
  ✅  Emerging / weak-signal themes      → rare issues (3-6 occurrences)
  ✅  Loud minority                      → few respondents with extreme frustration
  ✅  Cross-question theme recurrence    → same theme seeded across multiple questions
  ✅  Kenya/Africa-specific context      → M-Pesa, boda-boda, county, village, barazas
  ✅  African-English fillers            → "kindly", "whereby", "humbly"
  ✅  Output as .xlsx                    → tests chardet + openpyxl loader

Questions mirror a typical One Acre Fund M&E survey:
  q1_inputs     : Experience with seed and fertilizer inputs
  q2_training   : Quality and usefulness of farmer training
  q3_fieldofficer : Field officer support and communication
  q4_loan       : Loan and repayment process experience
  q5_suggestions: Suggestions for improving One Acre Fund services

Run:
    pip install pandas openpyxl
    python simSurveyResponses.py
Outputs:
    oaf_farmer_survey_demo.xlsx   (500 rows × 6 columns)
========================================================================
"""

import random
import pandas as pd

random.seed(42)

# ══════════════════════════════════════════════════════════════════════
# 1.  CORE THEMES  ─ English
# ══════════════════════════════════════════════════════════════════════

THEMES_EN = {

    "input_quality": [
        "The seeds provided did not germinate well after planting",
        "Some of the maize seeds were of poor quality this season",
        "I received the fertilizer late and it affected my planting schedule",
        "The hybrid seeds performed well and gave a good harvest",
        "Some seeds were rotten when I opened the bag",
        "The fertilizer quantity was not enough for my entire farm",
        "Input quality has been inconsistent across the seasons",
    ],

    "input_delivery": [
        "The inputs were delivered very late after the rains had already started",
        "I had to travel far to collect my seed and fertilizer bundle",
        "Delivery to our village is always delayed compared to other areas",
        "The collection point is too far from where most farmers live",
        "I missed the planting window because the inputs arrived too late",
        "Inputs were delivered on time and I was able to plant early",
        "The boda-boda delivery option helped but the cost was added to our loan",
    ],

    "training_quality": [
        "The training on spacing and planting techniques improved my yield",
        "The field training sessions are too short and do not cover enough",
        "I did not understand some of the training because it was in English only",
        "Post-harvest handling training was very useful for reducing losses",
        "The demonstration plots helped me understand proper fertilizer application",
        "Training sessions are held during harvest time when farmers are too busy",
        "We need more training on drought-resistant crop varieties",
    ],

    "field_officer": [
        "Our field officer visits the group regularly and answers our questions",
        "The field officer is difficult to reach on the phone between visits",
        "Our field officer was changed three times this season without notice",
        "The field officer does not visit remote farmers in our sub-location",
        "Our field officer is knowledgeable and treats farmers with respect",
        "Field officer visits are too infrequent especially during growing season",
        "The field officer pushed us to take larger loans than we could repay",
    ],

    "loan_repayment": [
        "The repayment schedule is too tight and does not match the harvest cycle",
        "I was penalised for late repayment even though the harvest was delayed",
        "The loan amount is too small to cover my full farm input needs",
        "Group liability means I had to pay for a member who defaulted",
        "The interest rate is fair and I was able to repay after selling my produce",
        "M-Pesa repayment is convenient but sometimes the system fails",
        "I would like a longer grace period before repayments begin",
    ],

    "market_access": [
        "One Acre Fund helped us find a buyer for our surplus maize",
        "After training we had more produce but no reliable market to sell it",
        "The market price offered through the program was below the local market",
        "We were left to sell on our own after harvest with no support",
        "The aggregation model helped us negotiate better prices as a group",
        "Transport costs to the nearest market eat into our profits significantly",
        "We need support selling vegetables and legumes not just maize",
    ],

    "group_dynamics": [
        "Some group members do not attend meetings and it affects all of us",
        "The group leader does not communicate information to all members",
        "Our group works well together and we support each other during hard times",
        "Group meetings are held far away and transport is a challenge",
        "Conflicts over loan repayment have broken trust in the group",
        "New members do not understand the group rules and cause problems",
        "The group savings model has helped us invest in small farm improvements",
    ],

    "positive": [
        "My yield doubled after following One Acre Fund planting advice",
        "One Acre Fund has truly changed my farming and my family income",
        "I am now food secure for the whole year for the first time",
        "The program gave me access to inputs I could never afford before",
        "I am very satisfied with the support I receive from One Acre Fund",
        "Thanks to One Acre Fund training I can now train other farmers in my village",
        "The combination of seeds, fertilizer and training has made a real difference",
    ],
}

# ══════════════════════════════════════════════════════════════════════
# 2.  SWAHILI RESPONSES
# ══════════════════════════════════════════════════════════════════════

THEMES_SW = {

    "input_quality": [
        "Mbegu walizotupa hazikuota vizuri na tulipoteza muda wa kupanda",
        "Mbolea haikutosha kwa shamba langu lote na mavuno yalikuwa kidogo",
        "Mwaka huu mbegu zilikuwa bora sana na nilivuna vizuri sana",
        "Mfuko wa mbolea ulikuwa na uzito mdogo kuliko ulivyosemwa",
    ],

    "input_delivery": [
        "Mbegu zilifika baada ya mvua kuanza na sikuweza kupanda kwa wakati",
        "Lazima niende mbali sana kuchukua pembejeo zangu kila msimu",
        "Usafirishaji wa pembejeo kwa vijiji vya mbali ni tatizo kubwa sana",
        "Nilipokea pembejeo mapema na mavuno yangu yalikuwa mazuri mwaka huu",
    ],

    "training_quality": [
        "Mafunzo ya upandaji yalisaidia sana na mavuno yangu yaliongezeka",
        "Mafunzo yanafanyika wakati wa mavuno na wakulima hawana muda",
        "Sijui Kiingereza na mafunzo mengi hayatolewa kwa Kiswahili",
        "Mashamba ya maonyesho yalisaidia kuelewa matumizi sahihi ya mbolea",
    ],

    "field_officer": [
        "Afisa wetu wa shamba anatutembelea mara kwa mara na anasaidia sana",
        "Afisa wa shamba habadiliki mara moja na tunapoteza muda wa kujifunza",
        "Afisa wa shamba hawafiki kwa wakulima wa mbali katika jimbo letu",
        "Afisa wetu anajua kilimo vizuri na anatuambia ukweli",
    ],

    "loan_repayment": [
        "Ratiba ya malipo ni ngumu sana na haifanani na wakati wa mavuno",
        "Kulipa kupitia M-Pesa ni rahisi lakini mfumo unashindwa mara nyingi",
        "Mwanakikundi mmoja hakurudisha mkopo na sisi wote tuliadhibiwa",
        "Niliweza kulipa mkopo wangu kwa wakati baada ya kuuza mahindi yangu",
    ],

    "positive": [
        "One Acre Fund imenibadilisha kabisa na sasa ninalisha familia yangu",
        "Kwa mara ya kwanza nina chakula cha kutosha kwa mwaka wote",
        "Programu hii imenisaidia sana kupata mbegu bora na mafunzo mazuri",
        "Nilifurahi sana na huduma ya One Acre Fund msimu huu wa kilimo",
    ],

    "market_access": [
        "Baada ya mavuno hatuna soko la uhakika la kuuzia mazao yetu",
        "Bei waliyotoa kwa mahindi ilikuwa chini ya bei ya soko la kawaida",
        "Usafiri hadi sokoni unachukua faida nyingi za mkulima mdogo",
    ],
}

# ══════════════════════════════════════════════════════════════════════
# 3.  CODE-SWITCHED (Sheng / mixed)
# ══════════════════════════════════════════════════════════════════════

CODE_SWITCHED = [
    "Maze mbegu zilitupwa late sana, tulimiss planting season kabisa",
    "Field officer wetu yuko rada, haji shambani mara nyingi kama anapaswa",
    "Bana the loan repayment inakuua, deadline ni tight sana after harvest",
    "Mbolea ilikuwa poa but delivery ilichelewa by almost three weeks",
    "Nimechoka na group meetings, wengine hawaji but tunalipa wote",
    "The training ilikuwa safi, yield yangu ilikuwa moto this season",
    "Walisema M-Pesa itafanya kazi but system ilikuwa down nikijaribu kulipa",
    "Unapata inputs nzuri lakini delivery point iko mbali sana manze",
    "Sawa sawa na program but field officer anabadilika kila msimu bila reason",
    "Mimi niko poa na One Acre Fund, wamenisaidia sana kilimo",
    "Boda-boda ya kuleta mbegu iliongeza cost ya loan yangu bila notice",
    "Training ilikuwa too fast, hatukuelewa spacing ya fertilizer vizuri",
    "Harvest ilikuwa noma this year, seeds walifanya kazi sawa sawa",
    "Stress ya group liability inafanya watu wengine waache program",
]

# ══════════════════════════════════════════════════════════════════════
# 4.  HEDGED / SARCASTIC POSITIVES
# ══════════════════════════════════════════════════════════════════════

HEDGED = [
    "The training was not bad I suppose but it could have been more practical",
    "I guess the seeds performed okay this season, nothing extraordinary really",
    "Not the best experience with repayment but I managed somehow",
    "The field officer was sort of helpful, could have visited more often",
    "It was acceptable I suppose but I expected better yields after all the training",
    "So so experience with One Acre Fund, some things work and some do not",
    "Meh, the inputs arrived eventually but the delay was frustrating",
    "Nothing special about the market linkage support if I am being honest",
    "The loan terms are okay I guess, did not really suit my harvest timeline",
    "Below average experience with delivery but at least the inputs came",
    "I would not say the program transformed my farm, but it helped a little",
    "The group model works in some ways but creates problems in others",
]

# ══════════════════════════════════════════════════════════════════════
# 5.  NEAR-DUPLICATES  (late delivery complaints)
# ══════════════════════════════════════════════════════════════════════

NEAR_DUPLICATES = [
    "The inputs were delivered late and I missed the planting window",
    "Late delivery of seeds meant I could not plant on time this season",
    "My seeds and fertilizer arrived too late and I missed the rains",
    "Delivery was very late and I was not able to plant at the right time",
    "Late inputs delivery caused me to plant late and lose yield",
]

# ══════════════════════════════════════════════════════════════════════
# 6.  LOUD MINORITY  (extreme frustration)
# ══════════════════════════════════════════════════════════════════════

LOUD_MINORITY = [
    "I am absolutely furious with One Acre Fund, they have ruined my entire season",
    "This program is a complete disaster and I have lost money every single year",
    "The field officer lied to us about the loan terms and I am deeply betrayed",
    "I am outraged that group members who default face zero consequences while I suffer",
    "This is the worst agricultural program I have ever joined, deeply disappointed",
    "They promised market support and completely abandoned us after harvest",
    "I am beyond frustrated with the seed quality this season, total crop failure",
    "The loan system is predatory and designed to trap poor farmers in debt",
    "I was humiliated at the group meeting and the field officer did nothing",
    "Appalling service, wrong fertilizer delivered, no apology, no replacement",
]

# ══════════════════════════════════════════════════════════════════════
# 7.  EMERGING / WEAK-SIGNAL THEMES  (3-6 occurrences only)
# ══════════════════════════════════════════════════════════════════════

EMERGING = [
    # Climate adaptation — growing concern
    "We need training on drought-tolerant varieties as rains are becoming unpredictable",
    "Flooding destroyed my crop and One Acre Fund did not offer any support after",
    "Climate change is making our planting calendar unreliable season after season",
    # Gender issues — rare but important
    "Female farmers in our group are given less support than male farmers",
    "My husband controls the loan but I do all the farming work on our plot",
    # Mental health / farmer wellbeing
    "The pressure of loan repayment after crop failure caused me serious stress",
    # Digital / app literacy
    "The One Acre Fund app is too complicated for farmers who are not literate",
    "I cannot use the USSD service because I do not have enough mobile data",
]

# ══════════════════════════════════════════════════════════════════════
# 8.  AFRICAN-ENGLISH FILLERS
# ══════════════════════════════════════════════════════════════════════

AFRICAN_ENGLISH_FILLERS = [
    "Kindly One Acre Fund should improve the delivery whereby farmers do not suffer",
    "Humbly I request that field officers be increased so that they visit us more sincerely",
    "Whereby the rains come early, the inputs should also be delivered early accordingly",
    "Basically the loan repayment schedule is too tight and does not suit small farmers",
    "Truly the seed quality this season was a serious issue that needs to be addressed",
    "Obviously the training should be done in local languages for better understanding",
]

# ══════════════════════════════════════════════════════════════════════
# 9.  KENYA / AFRICA-SPECIFIC CONTEXT
# ══════════════════════════════════════════════════════════════════════

KENYA_SPECIFIC = [
    "M-Pesa repayment is convenient but failed when I tried to pay before deadline",
    "Boda-boda delivery of inputs to our village adds cost to our loan without notice",
    "The county extension officer and One Acre Fund field officer give different advice",
    "Our sub-location chief holds barazas but One Acre Fund never sends a representative",
    "Jua kali farmers like us cannot afford to keep attending afternoon training sessions",
    "The nearest collection point is in the next ward which requires matatu fare",
    "CRB listing for loan default is very harsh for subsistence farmers in our area",
    "One Acre Fund should partner with Equity Bank so we can borrow more affordably",
    "We want the program to include dairy goats and chickens not just maize and beans",
    "The village elder said One Acre Fund is not registered with the county government",
    "USSD code for loan balance check does not work on Airtel, only on Safaricom",
    "Women in our group want separate training sessions that fit their household schedule",
]

# ══════════════════════════════════════════════════════════════════════
# SURVEY QUESTIONS
# ══════════════════════════════════════════════════════════════════════

QUESTIONS = [
    "q1_inputs",
    "q2_training",
    "q3_fieldofficer",
    "q4_loan",
    "q5_suggestions",
]

# Theme-to-question affinity — drives cross-question pattern detection
QUESTION_THEME_AFFINITY = {
    "q1_inputs":      ["input_quality", "input_delivery", "market_access"],
    "q2_training":    ["training_quality", "field_officer", "positive"],
    "q3_fieldofficer":["field_officer", "group_dynamics", "training_quality"],
    "q4_loan":        ["loan_repayment", "group_dynamics", "market_access"],
    "q5_suggestions": ["input_delivery", "training_quality", "loan_repayment",
                       "market_access", "positive", "field_officer"],
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

    if use_hedged and random.random() < 0.7:
        return random.choice(HEDGED)

    if use_code_switch and random.random() < 0.8:
        return random.choice(CODE_SWITCHED)

    if use_kenya and random.random() < 0.7:
        return random.choice(KENYA_SPECIFIC)

    if use_filler and random.random() < 0.6:
        return random.choice(AFRICAN_ENGLISH_FILLERS)

    if use_swahili:
        affinity = QUESTION_THEME_AFFINITY.get(question, list(THEMES_SW.keys()))
        sw_themes = [
            t for t in respondent_themes
            if t in THEMES_SW and t in affinity
        ]
        if sw_themes:
            return random.choice(THEMES_SW[random.choice(sw_themes)])
        t = random.choice(list(THEMES_SW.keys()))
        return random.choice(THEMES_SW[t])

    # Standard English — prefer question-affinity themes
    affinity = QUESTION_THEME_AFFINITY.get(question, list(THEMES_EN.keys()))
    candidate_themes = [t for t in respondent_themes if t in affinity]
    if not candidate_themes:
        candidate_themes = respondent_themes

    theme = random.choice(candidate_themes)
    return random.choice(THEMES_EN[theme])


# ══════════════════════════════════════════════════════════════════════
# GENERATE DATASET
# ══════════════════════════════════════════════════════════════════════

def generate_dataset(
    n: int = 500,
    output: str = "oaf_farmer_survey_demo.xlsx",
) -> pd.DataFrame:

    all_themes = list(THEMES_EN.keys())
    rows = []

    # --- Special respondent pools ---
    n_swahili     = int(n * 0.15)   # 15%  pure Swahili  (rural farmers)
    n_code_switch = int(n * 0.12)   # 12%  Sheng / mixed (peri-urban farmers)
    n_hedged      = int(n * 0.08)   # 8%   hedged / sarcastic
    n_near_dup    = int(n * 0.04)   # 4%   near-duplicate late-delivery complaints
    n_loud        = 10              # fixed: loud minority
    n_kenya       = int(n * 0.10)   # 10%  Kenya-specific context
    n_filler      = int(n * 0.04)   # 4%   African-English filler language

    idx = list(range(n))
    random.shuffle(idx)

    swahili_idx    = set(idx[:n_swahili]);          idx = idx[n_swahili:]
    code_idx       = set(idx[:n_code_switch]);       idx = idx[n_code_switch:]
    hedged_idx     = set(idx[:n_hedged]);            idx = idx[n_hedged:]
    near_dup_idx   = set(idx[:n_near_dup]);          idx = idx[n_near_dup:]
    loud_idx       = set(idx[:n_loud]);              idx = idx[n_loud:]
    kenya_idx      = set(idx[:n_kenya]);             idx = idx[n_kenya:]
    filler_idx     = set(idx[:n_filler])

    for i in range(n):
        respondent_themes = random.sample(all_themes, random.randint(1, 3))

        is_swahili     = i in swahili_idx
        is_code_switch = i in code_idx
        is_hedged      = i in hedged_idx
        is_near_dup    = i in near_dup_idx
        is_loud        = i in loud_idx
        is_kenya       = i in kenya_idx
        is_filler      = i in filler_idx

        row = {"response_id": i + 1}

        for q in QUESTIONS:
            # Loud minority: extreme language on loan and field officer questions
            if is_loud and q in ("q4_loan", "q3_fieldofficer"):
                row[q] = random.choice(LOUD_MINORITY)
                continue

            # Near-duplicates: always a late-delivery complaint
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

    # --- Inject emerging / weak-signal themes into q5_suggestions ---
    emerging_targets = random.sample(range(n), min(len(EMERGING) * 3, n))
    for j, target_i in enumerate(emerging_targets):
        rows[target_i]["q5_suggestions"] = EMERGING[j % len(EMERGING)]

    df = pd.DataFrame(rows)
    df.to_excel(output, index=False)

    # ── Summary ──────────────────────────────────────────────────────
    n_standard = (
        n - n_swahili - n_code_switch - n_hedged
        - n_near_dup - n_loud - n_kenya - n_filler
    )

    print(f"\n✅  Dataset saved → {output}")
    print(f"    Shape : {df.shape[0]} rows × {df.shape[1]} columns\n")
    print("Respondent composition:")
    print(f"  Standard English        : {n_standard}")
    print(f"  Swahili (rural farmers) : {n_swahili}  ({n_swahili/n*100:.0f}%)")
    print(f"  Code-switched / Sheng   : {n_code_switch}  ({n_code_switch/n*100:.0f}%)")
    print(f"  Hedged / sarcastic      : {n_hedged}  ({n_hedged/n*100:.0f}%)")
    print(f"  Near-duplicate (delivery): {n_near_dup}  ({n_near_dup/n*100:.0f}%)")
    print(f"  Loud minority           : {n_loud}   ({n_loud/n*100:.0f}%)")
    print(f"  Kenya-specific context  : {n_kenya}  ({n_kenya/n*100:.0f}%)")
    print(f"  African-English filler  : {n_filler}  ({n_filler/n*100:.0f}%)")
    print(f"  Emerging issues injected: {min(len(EMERGING)*3, n)} responses\n")

    print("Pipeline features exercised:")
    print("  ✅  Language detection (sw / en / mixed Sheng)")
    print("  ✅  Swahili stop-words + custom sentiment lexicon")
    print("  ✅  Hedged / sarcastic sentiment down-scoring")
    print("  ✅  Near-duplicate deduplication (late delivery theme)")
    print("  ✅  Emerging issues sheet (climate, gender, digital literacy)")
    print("  ✅  Loud minority — Intensity Score vs Impact Score separation")
    print("  ✅  Cross-question theme recurrence (Executive Summary)")
    print("  ✅  Kenya-specific recommendations (M-Pesa, boda-boda, barazas)")
    print("  ✅  African-English filler word filtering")
    print("  ✅  .xlsx input (chardet + openpyxl loader)\n")

    print("Survey questions generated:")
    descs = {
        "q1_inputs":       "Seed & fertilizer input experience",
        "q2_training":     "Training quality & usefulness",
        "q3_fieldofficer": "Field officer support & communication",
        "q4_loan":         "Loan & repayment process",
        "q5_suggestions":  "Suggestions for improvement",
    }
    for q, desc in descs.items():
        print(f"  {q:20s} → {desc}")

    print()
    print(df.head(6).to_string(index=False))

    return df


# ══════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    df = generate_dataset(n=500, output="oaf_farmer_survey_demo.xlsx")