"""
╔═══════════════════════════════════════════════════════════════════════════╗
║                           TRIGGER ENGINE                                 ║
╠═══════════════════════════════════════════════════════════════════════════╣
║                                                                           ║
║ PURPOSE: Detect high/medium-priority issues from feedback themes.        ║
║                                                                           ║
║ USAGE:                                                                    ║
║   engine = TriggerEngine(thresholds)                                     ║
║   triggers = engine.evaluate(theme)                                      ║
║                                                                           ║
║ TRIGGER STATUS ASSIGNMENT CRITERIA (RULE-BASED):                        ║
║ ┌─────────────────────────────────────────────────────────────────────┐  ║
║ │ HIGH PRIORITY:                                                     │  ║
║ │    Trigger: Impact > 15 AND Sentiment = "Problem"                  │  ║
║ │    Meaning: Widespread, strongly negative issue                    │  ║
║ │    Example: "Late Delivery" with 28 mentions + Problem sentiment   │  ║
║ │    Action: IMMEDIATE intervention (within 14 days)                 │  ║
║ │                                                                   │  ║
║ │ MEDIUM PRIORITY:                                                   │  ║
║ │    Trigger: Impact > 8 AND Impact <= 15 AND Sentiment = "Problem"  │  ║
║ │    Meaning: Significant concern from subset of farmers             │  ║
║ │    Example: "Loan Repayment" with 12 mentions + Problem sentiment  │  ║
║ │    Action: PLANNED response (within 30 days)                       │  ║
║ │                                                                   │  ║
║ │ POSITIVE:                                                          │  ║
║ │    Trigger: Impact > 15 AND Sentiment = "Positive"                 │  ║
║ │    Meaning: Strength to amplify (working solutions)                │  ║
║ │    Example: "Acre Fund Support" with 27 mentions + Positive        │  ║
║ │    Action: AMPLIFY (highlight in communications)                   │  ║
║ │                                                                   │  ║
║ │ LOW/NEUTRAL:                                                       │  ║
║ │    Everything else (low impact OR neutral sentiment)               │  ║
║ │    Action: Monitor and review periodically                         │  ║
║ │                                                                   │  ║
║ │ CUSTOMIZATION NOTE:                                                │  ║
║ │    Thresholds are SECTOR-DEPENDENT. In your agriculture context:  │  ║
║ │    - high_impact: 15 mentions (adjust based on typical problem     │  ║
║ │      volumes in OTHER SECTORS: healthcare=10, WASH=12)            │  ║
║ │    - medium_impact: 8 mentions (adjust based on significance)      │  ║
║ │    Update thresholds in config/config.yaml for sector changes.     │  ║
║ └─────────────────────────────────────────────────────────────────────┘  ║
║                                                                           ║
║ THRESHOLDS (CUSTOMIZABLE IN config/config.yaml):                        ║
║   Smallholder agriculture: high_impact=15, medium_impact=8               ║
║   Primary healthcare: high_impact=10, medium_impact=5                    ║
║   WASH services: high_impact=12, medium_impact=6                         ║
║   Microfinance: high_impact=10, medium_impact=6                          ║
╚═══════════════════════════════════════════════════════════════════════════╝
"""


class TriggerEngine:
    """
    Evaluates themes for priority-level escalation.
    
    Assigns priority levels (HIGH, MEDIUM) based on:
    - Impact count (how many responses mention this)
    - Sentiment (is this a problem or positive feedback?)
    """

    def __init__(self, thresholds: dict):
        """
        Initialize with configurable thresholds.
        
        Args:
            thresholds: dict with keys:
                - high_impact: count threshold for RED escalation
                - medium_impact: count threshold for YELLOW escalation
        """
        self.thresholds = thresholds

    def evaluate(self, theme: dict) -> list:
        """
        Evaluate a theme and return list of triggers (always returns ONE trigger).
        
        Args:
            theme: dict with keys like 'impact' (count), 'sentiment', 'name'
        
        Returns:
            list with a single trigger dict (ALWAYS present for reporting consistency)
        """
        triggers = []
        
        impact_count = theme.get("impact", 0)
        sentiment = str(theme.get("sentiment", "Neutral")).strip().title()

        high_threshold = self.thresholds.get("high_impact", 15)
        medium_threshold = self.thresholds.get("medium_impact", 8)

        # ─────────────────────────────────────────────
        # NORMALIZE SENTIMENT (safety layer)
        # ─────────────────────────────────────────────
        if sentiment not in ["Problem", "Negative", "Positive"]:
            sentiment = "Neutral"

        # ─────────────────────────────────────────────
        # PRIORITY LOGIC (FIXED BOUNDARIES)
        # ─────────────────────────────────────────────

        # 🔴 HIGH PRIORITY
        if impact_count >= high_threshold and sentiment in ("Problem", "Negative"):
            triggers.append({
                "level": "HIGH",
                "icon": "",
                "message": "Immediate intervention required",
                "deadline_days": 14,
                "color": "#FF0000"
            })

        # 🟡 MEDIUM PRIORITY
        elif impact_count >= medium_threshold and sentiment in ("Problem", "Negative"):
            triggers.append({
                "level": "MEDIUM",
                "icon": "",
                "message": "Plan response and monitor closely",
                "deadline_days": 30,
                "color": "#FFA500"
            })

        # 🟢 POSITIVE (STRENGTH)
        elif impact_count >= high_threshold and sentiment == "Positive":
            triggers.append({
                "level": "POSITIVE",
                "icon": "",
                "message": "Strength to amplify and replicate",
                "deadline_days": 60,
                "color": "#00AA00"
            })

        # ⚪ LOW / NEUTRAL (NEW — IMPORTANT FOR REPORTING)
        else:
            triggers.append({
                "level": "LOW",
                "icon": "",
                "message": "Monitor — no immediate action required",
                "deadline_days": 60,
                "color": "#999999"
            })

        return triggers

