# In[1]

import pandas as pd
import numpy as np
import re

from sentence_transformers import SentenceTransformer
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.feature_extraction.text import TfidfVectorizer

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.utils import get_column_letter

from nltk.sentiment import SentimentIntensityAnalyzer
import nltk

nltk.download("vader_lexicon")


class FeedbackInsightEngine:

    def __init__(self, file_path, text_columns):

        self.file_path = file_path
        self.text_columns = text_columns

        self.df = pd.read_csv(file_path)

        self.vectorizer = TfidfVectorizer(
            stop_words="english",
            ngram_range=(1,2),
            min_df=2,
            max_df=0.9,
            max_features=8000
        )

        self.model = SentenceTransformer("all-MiniLM-L6-v2")

        self.sentiment_model = SentimentIntensityAnalyzer()

        # store results per question
        self.question_results = {}

    # ------------------------------------------------
    # CLEAN TEXT
    # ------------------------------------------------

    def clean_text(self, text):

        text = str(text).lower()
        text = re.sub(r"[^a-z\s]", " ", text)
        text = re.sub(r"\s+", " ", text).strip()

        return text

    # ------------------------------------------------
    # SPLIT RESPONSES INTO SENTENCES
    # ------------------------------------------------

    def split_sentences(self, text):

        parts = re.split(r"[.!?;]", text)

        sentences = []

        for p in parts:

            p = p.strip()

            if len(p.split()) >= 3:
                sentences.append(p)

        return sentences

    # ------------------------------------------------
    # COMPRESS TEXT (FOR BETTER CLUSTERING)
    # ------------------------------------------------

    def compress_text(self, text):

        tokens = text.split()

        filtered = [
            t for t in tokens
            if len(t) > 3
        ]

        return " ".join(filtered)

    def preprocess(self):

        rows = []

        for text in self.q_df["response"]:

            clean = self.clean_text(text)

            sentences = self.split_sentences(clean)

            for s in sentences:

                rows.append(s)

        self.q_df = pd.DataFrame({"clean_text": rows})

        # compressed version for clustering
        self.q_df["compressed_text"] = self.q_df["clean_text"].apply(self.compress_text)

    # ------------------------------------------------
    # BUILD TF-IDF MATRIX
    # ------------------------------------------------

    def build_tfidf_matrix(self):

        texts = self.q_df["compressed_text"].tolist()

        self.tfidf_matrix = self.vectorizer.fit_transform(texts)

        self.feature_names = self.vectorizer.get_feature_names_out()

    # ------------------------------------------------
    # SENTIMENT ANALYSIS
    # ------------------------------------------------

    def compute_sentiment(self):

        scores = []

        for text in self.q_df["clean_text"]:
            score = self.sentiment_model.polarity_scores(text)["compound"]
            scores.append(score)

        self.q_df["sentiment"] = scores

    # ------------------------------------------------
    # EMBEDDINGS
    # ------------------------------------------------

    def create_embeddings(self):

        texts = self.q_df["compressed_text"].tolist()

        embeddings = self.model.encode(
            texts,
            batch_size=64,
            show_progress_bar=True
        )

        self.embeddings = np.array(embeddings, dtype="float32")

    # ------------------------------------------------
    # CLUSTER SELECTION
    # ------------------------------------------------

    def choose_clusters(self):

        n = len(self.q_df)

        if n < 200:

            scores = []

            for k in range(2, 8):

                kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)

                labels = kmeans.fit_predict(self.embeddings)

                score = silhouette_score(self.embeddings, labels)

                scores.append((k, score))

            return max(scores, key=lambda x: x[1])[0]

        elif n < 2000:

            return int(np.sqrt(n / 2))

        else:

            return min(20, int(np.sqrt(n / 2)))

    # ------------------------------------------------
    # CLUSTERING
    # ------------------------------------------------

    def cluster_feedback(self):

        k = self.choose_clusters()

        self.kmeans = KMeans(
            n_clusters=k,
            random_state=42,
            n_init=10
        )

        self.q_df["cluster"] = self.kmeans.fit_predict(self.embeddings)

    # ------------------------------------------------
    # FILTER LOW QUALITY CLUSTERS
    # ------------------------------------------------

    def filter_clusters(self, min_size=5):

        cluster_counts = self.q_df["cluster"].value_counts()

        valid_clusters = cluster_counts[cluster_counts >= min_size].index

        self.q_df = self.q_df[self.q_df["cluster"].isin(valid_clusters)].copy()

    # ------------------------------------------------
    # KEYWORD EXTRACTION
    # ------------------------------------------------

    # ------------------------------------------------
    # PHRASE-AWARE KEYWORD EXTRACTION
    # ------------------------------------------------

    def extract_keywords(self, cluster_indices, top_n=6):

        cluster_matrix = self.tfidf_matrix[cluster_indices]

        cluster_scores = np.asarray(cluster_matrix.mean(axis=0)).ravel()

        global_scores = np.asarray(self.tfidf_matrix.mean(axis=0)).ravel()

        importance = cluster_scores - global_scores

        top_idx = importance.argsort()[::-1]

        keywords = []

        generic_words = {
            "service","good","bad","thing","things",
            "experience","people","place","really",
            "very","much"
        }

        for idx in top_idx:

            term = self.feature_names[idx]

            # skip generic terms
            if term in generic_words:
                continue

            # prioritize phrases
            if " " in term:
                keywords.append(term)
            else:
                keywords.append(term)

            if len(keywords) >= top_n:
                break

        return keywords

    # ------------------------------------------------
    # THEME NAME GENERATION
    # ------------------------------------------------

    def generate_theme_name(self, keywords):

        if not keywords:
            return "Other Feedback"

        primary = keywords[0]
        secondary = keywords[1] if len(keywords) > 1 else ""

        phrase = f"{primary} {secondary}".strip()

        phrase = phrase.replace("_", " ")

        words = phrase.split()

        # remove duplicates
        words = list(dict.fromkeys(words))

        theme = " ".join(words)

        # clean formatting
        theme = theme.title()

        return theme

    # ------------------------------------------------
    # CLUSTER SUMMARY
    # ------------------------------------------------

    def build_cluster_summary(self):

        summary = {}

        for cid in sorted(self.q_df["cluster"].unique()):

            subset = self.q_df[self.q_df["cluster"] == cid]

            texts = subset["clean_text"].tolist()

            indices = subset.index.tolist()

            keywords = self.extract_keywords(indices)

            summary[cid] = {

                "count": len(texts),
                "keywords": keywords,
                "samples": texts[:5]

            }

        self.cluster_summary = summary

    # ------------------------------------------------
    # SUMMARY TABLE
    # ------------------------------------------------

    def build_summary_table(self):

        rows = []

        total = len(self.df)

        for cid, info in self.cluster_summary.items():

            theme = self.generate_theme_name(info["keywords"])

            subset = self.q_df[self.q_df["cluster"] == cid]

            avg_sentiment = subset["sentiment"].mean()

            pct = round((info["count"] / total) * 100, 2)

            severity = info["count"] * abs(avg_sentiment)

            rows.append({

                "Theme": theme,

                "Response Count": info["count"],

                "Share (%)": pct,

                "Impact Score": round(severity, 2),

                "Average Sentiment": round(avg_sentiment, 3),

                "Key Keywords": ", ".join(info["keywords"]),

                "Representative Quote": info["samples"][0]

            })

        summary_df = pd.DataFrame(rows)

        summary_df = summary_df.sort_values(
            "Impact Score",
            ascending=False
        )

        self.summary_df = summary_df

    # ------------------------------------------------
    # THEME SENTIMENT LABEL
    # ------------------------------------------------

    def classify_theme_sentiment(self):

        labels = []

        for score in self.summary_df["Average Sentiment"]:

            if score < -0.1:
                labels.append("Problem")

            elif score > 0.1:
                labels.append("Positive")

            else:
                labels.append("Neutral")

        self.summary_df["Theme Sentiment"] = labels

    # ------------------------------------------------
    # MERGE SIMILAR THEMES
    # ------------------------------------------------

    def merge_similar_themes(self, similarity_threshold=0.75):

        themes = self.summary_df["Theme"].tolist()

        if len(themes) <= 1:
            return

        embeddings = self.model.encode(
            themes,
            convert_to_numpy=True,
            normalize_embeddings=True
        )

        merged_map = {}

        used = set()

        for i, theme_i in enumerate(themes):

            if i in used:
                continue

            merged_map[theme_i] = [i]

            for j in range(i + 1, len(themes)):

                if j in used:
                    continue

                sim = np.dot(embeddings[i], embeddings[j])

                if sim >= similarity_threshold:

                    merged_map[theme_i].append(j)
                    used.add(j)

        new_rows = []

        for main_theme, idxs in merged_map.items():

            subset = self.summary_df.iloc[idxs]

            combined = {

                "Theme": main_theme,
                "Response Count": subset["Response Count"].sum(),
                "Share (%)": subset["Share (%)"].sum(),
                "Impact Score": subset["Impact Score"].sum(),
                "Average Sentiment": subset["Average Sentiment"].mean(),
                "Key Keywords": ", ".join(subset["Key Keywords"].tolist()),
                "Representative Quote": subset["Representative Quote"].iloc[0],
                "Recommendation": subset["Recommendation"].iloc[0],
                "Priority": subset["Priority"].iloc[0]

            }

            new_rows.append(combined)

        self.summary_df = pd.DataFrame(new_rows)

        self.summary_df = self.summary_df.sort_values(
            "Impact Score",
            ascending=False
        )

    # ------------------------------------------------
    # RECOMMENDATIONS
    # ------------------------------------------------

    def generate_recommendations(self):

        recs = []

        for theme in self.summary_df["Theme"]:

            t = theme.lower()

            if "cost" in t or "price" in t or "fee" in t:

                rec = "Review pricing structure or introduce subsidies."

            elif "wait" in t or "delay" in t:

                rec = "Increase staffing or optimize service workflow."

            elif "transport" in t:

                rec = "Improve transportation access or provide mobile services."

            elif "staff" in t:

                rec = "Provide staff training on service quality."

            elif "distance" in t or "walk" in t:

                rec = "Consider expanding service coverage or satellite clinics."

            else:

                rec = "Investigate further through targeted surveys."

            recs.append(rec)

        self.summary_df.insert(5, "Recommendation", recs)

    #-------------------------------------------------
    # INSIGHT GENERATION
    #-------------------------------------------------

    def generate_key_insights(self):

        insights = []

        top = self.summary_df.iloc[0]

        insights.append(
            f"The most common issue reported by respondents is '{top['Theme']}', "
            f"mentioned in {top['Share (%)']}% of responses."
        )

        negative = self.summary_df.sort_values("Average Sentiment").iloc[0]

        insights.append(
            f"The theme with the most negative sentiment is '{negative['Theme']}', "
            f"suggesting strong dissatisfaction among respondents."
        )

        high_severity = self.summary_df.iloc[0]

        insights.append(
            f"Based on severity scoring, '{high_severity['Theme']}' appears to be the "
            f"most critical issue requiring attention."
        )

        self.key_insights = insights

    #-------------------------------------------------
    # PRIORITY LEVELS
    #-------------------------------------------------

    def add_priority_labels(self):

        priorities = []

        for score in self.summary_df["Impact Score"]:

            if score > 20:
                priorities.append("High Priority")

            elif score > 10:
                priorities.append("Medium Priority")

            else:
                priorities.append("Low Priority")

        self.summary_df.insert(4, "Priority", priorities)

    #-------------------------------------------------
    # DATASET SUMMARY
    #-------------------------------------------------

    def generate_dataset_summary(self):

        total = len(self.q_df)

        avg_length = self.q_df["clean_text"].apply(
            lambda x: len(str(x).split())
        ).mean()

        self.dataset_summary = {
            "Total Responses": total,
            "Average Response Length": round(avg_length, 1),
            "Themes Identified": len(self.summary_df)
        }

    # ------------------------------------------------
    # LABEL DATA
    # ------------------------------------------------

    def label_data(self):

        theme_map = {}

        for cid, info in self.cluster_summary.items():
            theme_map[cid] = self.generate_theme_name(info["keywords"])

        self.q_df["Theme"] = self.q_df["cluster"].map(theme_map)

    # ------------------------------------------------
    # EXCEL STYLE HELPERS
    # ------------------------------------------------

    def style_header(self, ws):

        header_fill = PatternFill(
            start_color="2F5597",
            end_color="2F5597",
            fill_type="solid"
        )

        for cell in ws[1]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

    def zebra_rows(self, ws):

        fill = PatternFill(
            start_color="F2F2F2",
            end_color="F2F2F2",
            fill_type="solid"
        )

        for row in range(2, ws.max_row + 1):
            if row % 2 == 0:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = fill

    def auto_width(self, ws):

        for i, column_cells in enumerate(ws.columns, 1):

            max_length = 0
            column_letter = get_column_letter(i)

            for cell in column_cells:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass

            adjusted_width = min(max_length + 4, 60)

            ws.column_dimensions[column_letter].width = adjusted_width

    # ------------------------------------------------
    # EXCEL REPORT
    # ------------------------------------------------

    def build_excel_report(self, output="Feedback_Insights_Report.xlsx"):

        wb = Workbook()
        wb.remove(wb.active)

        # ------------------------------------------------
        # DASHBOARD
        # ------------------------------------------------

        ws = wb.create_sheet("Dashboard")

        ws["A1"] = "Customer Feedback Insights Dashboard"
        ws["A1"].font = Font(size=20, bold=True)

        ws["A3"] = "Total Responses"
        ws["B3"] = self.dataset_summary["Total Responses"]

        ws["A4"] = "Themes Identified"
        ws["B4"] = self.dataset_summary["Themes Identified"]

        ws["A6"] = "Top Issues"
        ws["A6"].font = Font(size=14, bold=True)

        top_themes = self.summary_df.head(5)
        problems = self.summary_df[
        self.summary_df["Theme Sentiment"] == "Problem"
        ].head(5)

        positives = self.summary_df[
        self.summary_df["Theme Sentiment"] == "Positive"
        ].head(5)

        ws.append([])
        ws.append(["Theme", "Response Count", "Share (%)", "Priority"])

        for r in top_themes[
            ["Theme", "Response Count", "Share (%)", "Priority"]
        ].itertuples(index=False):

            ws.append(list(r))

        self.style_header(ws)
        self.zebra_rows(ws)
        self.auto_width(ws)

        # Dashboard Chart

        data = Reference(
            ws,
            min_col=2,
            min_row=8,
            max_row=8 + len(top_themes)
        )

        cats = Reference(
            ws,
            min_col=1,
            min_row=9,
            max_row=8 + len(top_themes)
        )

        bar = BarChart()
        bar.title = "Top Issues by Responses"
        bar.y_axis.title = "Responses"

        bar.add_data(data, titles_from_data=True)
        bar.set_categories(cats)

        ws.add_chart(bar, "F6")

        # ------------------------------------------------
        # EXECUTIVE SUMMARY
        # ------------------------------------------------

        ws = wb.create_sheet("Executive Summary")

        ws["A1"] = "Customer Feedback Insights Report"
        ws["A1"].font = Font(size=20, bold=True)

        ws.merge_cells("A1:F1")

        ws["A3"] = "Dataset Overview"
        ws["A3"].font = Font(size=14, bold=True)

        ws["A5"] = "Total Responses"
        ws["B5"] = self.dataset_summary["Total Responses"]

        ws["A6"] = "Average Response Length"
        ws["B6"] = self.dataset_summary["Average Response Length"]

        ws["A7"] = "Themes Identified"
        ws["B7"] = self.dataset_summary["Themes Identified"]

        ws["A9"] = "Key Insights"
        ws["A9"].font = Font(size=14, bold=True)

        row = 11

        for insight in self.key_insights:
            ws[f"A{row}"] = f"• {insight}"
            row += 1

        self.auto_width(ws)
        
        ws["A15"] = "Top Customer Problems"
        ws["A15"].font = Font(size=14, bold=True)

        row = 17

        for r in problems[
            ["Theme","Response Count","Share (%)"]
        ].itertuples(index=False):

            ws[f"A{row}"] = r[0]
            ws[f"B{row}"] = r[1]
            ws[f"C{row}"] = r[2]

            row += 1
            
        ws["F15"] = "Top Positive Feedback"
        ws["F15"].font = Font(size=14, bold=True)

        row = 17

        for r in positives[
            ["Theme","Response Count","Share (%)"]
        ].itertuples(index=False):

            ws[f"F{row}"] = r[0]
            ws[f"G{row}"] = r[1]
            ws[f"H{row}"] = r[2]

            row += 1

        # ------------------------------------------------
        # THEME ANALYSIS
        # ------------------------------------------------

        ws = wb.create_sheet("Theme Analysis")

        columns = [
            "Theme",
            "Response Count",
            "Share (%)",
            "Impact Score",
            "Priority",
            "Recommendation",
            "Representative Quote",
            "Key Keywords",
            "Average Sentiment"
        ]

        ws.append(columns)

        for r in self.summary_df[columns].itertuples(index=False):
            ws.append(list(r))

        # wrap quote column
        for row in ws.iter_rows(min_row=2, min_col=7, max_col=7):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True)

        self.style_header(ws)
        self.zebra_rows(ws)

        ws.auto_filter.ref = ws.dimensions

        self.auto_width(ws)

        # ------------------------------------------------
        # CHARTS
        # ------------------------------------------------

        ws = wb.create_sheet("Charts")

        ws.append(["Theme", "Response Count"])

        for r in self.summary_df[
            ["Theme", "Response Count"]
        ].itertuples(index=False):

            ws.append(list(r))

        data = Reference(ws, min_col=2, min_row=1, max_row=len(self.summary_df) + 1)
        cats = Reference(ws, min_col=1, min_row=2, max_row=len(self.summary_df) + 1)

        bar = BarChart()
        bar.title = "Feedback Distribution by Theme"
        bar.y_axis.title = "Responses"

        bar.add_data(data, titles_from_data=True)
        bar.set_categories(cats)

        ws.add_chart(bar, "E2")

        pie = PieChart()
        pie.title = "Theme Share"

        pie.add_data(data, titles_from_data=True)
        pie.set_categories(cats)

        ws.add_chart(pie, "E20")

        self.auto_width(ws)

        # ------------------------------------------------
        # EXAMPLE QUOTES
        # ------------------------------------------------

        ws = wb.create_sheet("Example Quotes")

        ws.append(["Theme", "Quote"])

        for cid, info in self.cluster_summary.items():

            theme = self.generate_theme_name(info["keywords"])

            for q in info["samples"]:
                ws.append([theme, q])

        for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True)

        self.style_header(ws)
        self.zebra_rows(ws)

        self.auto_width(ws)

        # ------------------------------------------------
        # FULL DATA
        # ------------------------------------------------

        ws = wb.create_sheet("Clustered Responses")

        ws.append(self.q_df.columns.tolist())

        for r in self.q_df.itertuples(index=False):
            ws.append(list(r))

        self.style_header(ws)

        ws.auto_filter.ref = ws.dimensions

        self.auto_width(ws)

        wb.save(output)

        print("Report generated:", output)

    # ------------------------------------------------
    # FULL PIPELINE
    # ------------------------------------------------

    def run(self):

        for column in self.text_columns:

            print(f"\nAnalyzing question: {column}")

            # create question dataframe
            self.q_df = self.df[[column]].copy()
            self.q_df.columns = ["response"]

            print("Cleaning text...")
            self.preprocess()

            print("Building TF-IDF...")
            self.build_tfidf_matrix()

            print("Sentiment analysis...")
            self.compute_sentiment()

            print("Generating embeddings...")
            self.create_embeddings()

            print("Clustering responses...")
            self.cluster_feedback()

            print("Filtering clusters...")
            self.filter_clusters()

            print("Building cluster summary...")
            self.build_cluster_summary()

            print("Building summary table...")
            self.build_summary_table()

            print("Classifying theme sentiment...")
            self.classify_theme_sentiment()

            print("Generating insights...")
            self.generate_key_insights()

            print("Generating recommendations...")
            self.generate_recommendations()

            print("Adding priorities...")
            self.add_priority_labels()

            print("Merge similar themes...")
            self.merge_similar_themes()

            print("Dataset summary...")
            self.generate_dataset_summary()

            print("Labeling responses...")
            self.label_data()

            self.question_results[column] = {
                "summary": self.summary_df.copy(),
                "insights": self.key_insights.copy()
            }

        print("Building Excel report...")
        self.build_excel_report()

# In[2]

engine = FeedbackInsightEngine(
    "health_survey_responses.csv",
    text_columns=[
    "q1_service",
    "q2_staff",
    "q3_cost",
    "q4_transport"
    ])

engine.run()
# %%
