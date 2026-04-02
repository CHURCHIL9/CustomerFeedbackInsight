"""
╔═══════════════════════════════════════════════════════════════════════════╗
║                          ACTION ENGINE                                   ║
╠═══════════════════════════════════════════════════════════════════════════╣
║                                                                           ║
║ PURPOSE: Auto-assign owners, timelines, and action plans to themes.      ║
║                                                                           ║
║ USAGE:                                                                    ║
║   engine = ActionEngine(sector="smallholder agriculture")                ║
║   action_plan = engine.generate(theme)                                   ║
║                                                                           ║
║ ASSIGNMENT CRITERIA (RULE-BASED KEYWORD ROUTING):                        ║
║ ┌─────────────────────────────────────────────────────────────────────┐  ║
║ │ ⚠️  IMPORTANT: This is NOT AI-driven. It uses KEYWORD MATCHING.     │  ║
║ │                                                                   │  ║
║ │ OWNER ASSIGNMENT (current agriculture sector):                   │  ║
║ │ ├─ Logistics/Supply Chain: delivery, delivered, late, transport   │  ║
║ │ ├─ Finance/Loan Management: loan, repay, repayment, payment      │  ║
║ │ ├─ Training/Extension: training, education, learning, mafunzo    │  ║
║ │ ├─ Field Operations/Staffing: officer, field, coverage, visits   │  ║
║ │ ├─ Group Mobilization: group, member, community, chama, model    │  ║
║ │ ├─ Market Linkage: market, price, buyer, sell, linkage           │  ║
║ │ ├─ Input Quality/Sourcing: seed, mbolea, quality, germination    │  ║
║ │ └─ Program Manager (default): When no keywords match above        │  ║
║ │                                                                   │  ║
║ │ TIMELINE ASSIGNMENT (days to action):                           │  ║
║ │ ├─ Training/Extension issues: 7 days (urgent - affects farming)   │  ║
║ │ ├─ Finance/Logistics: 14 days (operational fixes)                 │  ║
║ │ ├─ Other issues: 30 days (normal planning cycle)                  │  ║
║ │ └─ Positive feedback: 60 days (amplification campaign)            │  ║
║ │                                                                   │  ║
║ │ WHY ALL "Program Manager"?:                                      │  ║
║ │ Theme names & keywords might not contain routing keywords.       │  ║
║ │ Example: Theme "Loans Are Too Tight" doesn't contain "repay"    │  ║
║ │ Solution: Update keyword list in assign_owner() to match data.   │  ║
║ └─────────────────────────────────────────────────────────────────────┘  ║
║                                                                           ║
║ CUSTOMIZATION FOR YOUR SECTOR:                                           ║
║ ┌─────────────────────────────────────────────────────────────────────┐  ║
║ │ STEP 1: Identify your teams (replace listed above)               │  ║
║ │         e.g., Healthcare: Pharmacy, Nursing, Admin, IT           │  ║
║ │                                                                   │  ║
║ │ STEP 2: Map keywords to your themes                              │  ║
║ │         Review your Excel report's "Theme" column                │  ║
║ │         Add missing keywords to routing dict below               │  ║
║ │                                                                   │  ║
║ │ STEP 3: Update assign_owner() method keywords                    │  ║
║ │         Example for HEALTHCARE:                                  │  ║
║ │         if "medication" in keywords: return "Pharmacy Lead"       │  ║
║ │         if "staff" in keywords: return "HR/Training"             │  ║
║ │                                                                   │  ║
║ │ STEP 4: Update timeline() method for your urgency levels         │  ║
║ │         Example for WASH:                                        │  ║
║ │         if "borehole" in keywords: return 7 (critical)           │  ║
║ │         elif "maintenance" in keywords: return 14 (important)    │  ║
║ │━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━  │  ║
║ │ TEST: After updating, re-run pipeline and check Excel report.   │  ║
║ │       "Action Owner" column should show variety of team names.   │  ║
║ └─────────────────────────────────────────────────────────────────────┘  ║
╚═══════════════════════════════════════════════════════════════════════════╝
"""


class ActionEngine:
    """
    Generates actionable action plans from themes.
    Assigns owners, deadlines, and concrete next steps.
    """

    def __init__(self, sector: str = ""):
        """
        Initialize with sector context for personalized ownership assignment.
        
        Args:
            sector: e.g., "smallholder agriculture", "healthcare", "WASH"
        """
        self.sector = sector

    def generate(self, theme: dict) -> dict:
        """
        Generate a complete action plan for a theme.
        
        Args:
            theme: dict with keys 'name', 'impact', 'priority', 'keywords', etc.
        
        Returns:
            dict with keys: owner, timeline, actions, urgency_level
        """
        return {
            "owner": self.assign_owner(theme),
            "timeline": self.timeline(theme),
            "actions": self.map_actions(theme),
            "urgency": self.urgency_level(theme),
        }

    def assign_owner(self, theme: dict) -> str:
        """
        Route issue to responsible team based on keywords (improved matching).
        
        
        ⚠️  RULE-BASED & SECTOR-AWARE:
        This is KEYWORD-BASED ROUTING, not AI-driven. Keywords must match
        theme names or problem statements to trigger assignment.
        
        If you're using this on a new sector:
        1. Update TEAM_KEYWORDS dict below with your sector's keywords
        2. Add new teams as needed
        3. Re-run pipeline
        
        Example: HEALTHCARE version
            TEAM_KEYWORDS = {
                "Pharmacy": ["medication", "drugs", "prescription"],
                "HR/Training": ["staff", "nurse", "doctor", "training"],
                "Operations": ["waiting", "queue", "patient", "bed"],
            }
        
        Example: WASH version
            TEAM_KEYWORDS = {
                "Infrastructure": ["borehole", "pipe", "tap", "water point"],
                "Field Ops": ["maintenance", "repair", "broken"],
                "Behavior Change": ["hygiene", "handwashing", "latrine"],
            }
        """

        keywords = " ".join(theme.get("keywords", [])).lower()
        theme_name = theme.get("name", "").lower()
        combined_text = f"{keywords} {theme_name}"

        # Define keyword groups (expanded)
        ROUTING_RULES = {
            "Logistics/Supply Chain": [
                "delivery", "delivered", "late", "delay", "transport",
                "collection", "boda", "distance", "access"
            ],
            "Finance/Loan Management": [
                "loan", "repay", "repayment", "payment", "mpesa",
                "credit", "cost", "expensive", "price", "penalty"
            ],
            "Training/Extension": [
                "training", "education", "learning", "mafunzo",
                "session", "teach", "knowledge", "skills"
            ],
            "Field Operations/Staffing": [
                "officer", "field", "visit", "coverage",
                "staff", "support", "reach", "availability"
            ],
            "Group Mobilization": [
                "group", "member", "community", "chama",
                "leader", "meeting", "coordination"
            ],
            "Market Linkage": [
                "market", "buyer", "sell", "price",
                "income", "profit", "linkage"
            ],
            "Input Quality/Sourcing": [
                "seed", "seeds", "fertilizer", "mbolea",
                "quality", "germination", "input"
            ],
        }

        # Score-based matching (instead of first-match)
        scores = {}
        for team, words in ROUTING_RULES.items():
            score = sum(1 for word in words if word in combined_text)
            if score > 0:
                scores[team] = score

        # Select best match
        if scores:
            return max(scores, key=scores.get)

        return "Program Manager"

    def timeline(self, theme: dict) -> int:
        """
        Return action deadline in days based ONLY on priority level.
        Priority is computed by TriggerEngine (single source of truth).
        
        Args:
            theme: dict with 'priority' field (HIGH/MEDIUM/LOW/POSITIVE)
        
        Returns:
            int: days to action (14/30/60)
        """
        priority = str(theme.get("priority", "MEDIUM")).strip().upper()
        
        # Clean priority mapping (TriggerEngine is source of truth)
        if priority == "HIGH":
            return 14
        elif priority == "MEDIUM":
            return 30
        elif priority == "POSITIVE":
            return 60
        else:  # LOW, NEUTRAL, or unknown
            return 60

    def map_actions(self, theme: dict) -> list:
        """
        Extract recommended actions from theme recommendation field.
        
        Args:
            theme: dict with 'recommendation' field (human-readable text)
        
        Returns:
            list of action strings (cleaned and limited)
        """
        recommendation = theme.get("recommendation", "")
        
        if not recommendation:
            return ["Conduct focus-group discussion to identify root cause"]
        
        # Split by common delimiters
        actions = []
        for delimiter in [";", "\n", "•"]:
            if delimiter in recommendation:
                raw_actions = recommendation.split(delimiter)
                actions = [a.strip() for a in raw_actions if a.strip()]
                break
        
        if not actions:
            # Single action / no delimiters
            actions = [recommendation.strip()]
        
        # Limit to max 5 actions (safe for Excel readability)
        return actions[:5]

    def urgency_level(self, theme: dict) -> str:
        """
        Simple urgency classification based on timeline (days).
        Returns one of: 'HIGH', 'MEDIUM', 'LOW'.
        """
        try:
            days = int(self.timeline(theme))
        except Exception:
            return "MEDIUM"

        if days <= 7:
            return "HIGH"
        if days <= 14:
            return "HIGH"
        if days <= 30:
            return "MEDIUM"
        return "LOW"
