import random
import pandas as pd

random.seed(42)

themes = {
    "waiting_time": [
        "I waited for many hours before being attended to.",
        "The queue was too long at the clinic.",
        "We spent the whole day waiting.",
        "Service was very slow and disorganized.",
        "Too much delay before seeing a doctor."
    ],
    "high_cost": [
        "The consultation fee is too expensive.",
        "I could not afford the treatment cost.",
        "Medicine prices are very high.",
        "Healthcare services are not affordable.",
        "The charges are beyond my income level."
    ],
    "stock_outs": [
        "There were no medicines available.",
        "The drugs were out of stock.",
        "I was told to buy medicine elsewhere.",
        "The pharmacy had empty shelves.",
        "Essential medicines were unavailable."
    ],
    "distance": [
        "The health center is very far from my home.",
        "I have to walk long distances to reach the clinic.",
        "Transport to the hospital is difficult.",
        "The facility is too far away.",
        "It takes hours to get to the hospital."
    ],
    "staff_attitude": [
        "The nurses were rude.",
        "Staff were not friendly.",
        "The doctor did not listen to me.",
        "Healthcare workers were disrespectful.",
        "Poor customer service at the facility."
    ],
    "equipment": [
        "The hospital lacks proper equipment.",
        "Machines were not working.",
        "There are no diagnostic tools available.",
        "The facility is poorly equipped.",
        "Broken equipment delays treatment."
    ],
    "transport": [
        "There is no reliable transport to the clinic.",
        "Public transport is expensive.",
        "Bad roads make it hard to reach the hospital.",
        "Transport costs are too high.",
        "Ambulances are not available."
    ],
    "no_issue": [
        "I did not face any major challenges.",
        "The service was satisfactory.",
        "Everything went smoothly.",
        "No problems accessing healthcare.",
        "I received help without issues."
    ]
}

responses = []

for _ in range(300):
    theme = random.choice(list(themes.keys()))
    sentence = random.choice(themes[theme])
    responses.append(sentence)

df = pd.DataFrame({
    "response_id": range(1, 301),
    "question": "What challenges did you face when accessing healthcare services in your area?",
    "response": responses
})

df.to_csv("health_survey_responses.csv", index=False)

print("Dataset created with shape:", df.shape)
df.head()