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

    def __init__(self, file_path, text_column):

        self.file_path = file_path
        self.text_column = text_column

        self.df = pd.read_csv(file_path)

        self.model = SentenceTransformer("all-MiniLM-L6-v2")

        self.sentiment_model = SentimentIntensityAnalyzer()

    # ------------------------------------------------
    # CLEAN TEXT
    # ------------------------------------------------

    def clean_text(self, text):

        text = str(text).lower()
        text = re.sub(r"[^a-z\s]", " ", text)
        text = re.sub(r"\s+", " ", text).strip()

        return text

    def preprocess(self):

        self.df["clean_text"] = self.df[self.text_column].apply(self.clean_text)

    # ------------------------------------------------
    # SENTIMENT ANALYSIS
    # ------------------------------------------------

    def compute_sentiment(self):

        scores = []

        for text in self.df["clean_text"]:
            score = self.sentiment_model.polarity_scores(text)["compound"]
            scores.append(score)

        self.df["sentiment"] = scores

    # ------------------------------------------------
    # EMBEDDINGS
    # ------------------------------------------------

    def create_embeddings(self):

        texts = self.df["clean_text"].tolist()

        embeddings = self.model.encode(
            texts,
            batch_size=32,
            show_progress_bar=True
        )

        self.embeddings = np.array(embeddings, dtype="float32")

    # ------------------------------------------------
    # CLUSTER SELECTION
    # ------------------------------------------------

    def choose_clusters(self):

        n = len(self.df)

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

        self.df["cluster"] = self.kmeans.fit_predict(self.embeddings)

    # ------------------------------------------------
    # KEYWORD EXTRACTION
    # ------------------------------------------------

    def extract_keywords(self, texts, top_n=6):

        vectorizer = TfidfVectorizer(
            stop_words="english",
            max_features=5000
        )

        X = vectorizer.fit_transform(texts)

        scores = np.asarray(X.mean(axis=0)).ravel()

        terms = vectorizer.get_feature_names_out()

        top_idx = scores.argsort()[-top_n:][::-1]

        return [terms[i] for i in top_idx]

    # ------------------------------------------------
    # CLUSTER SUMMARY
    # ------------------------------------------------

    def build_cluster_summary(self):

        summary = {}

        for cid in sorted(self.df["cluster"].unique()):

            subset = self.df[self.df["cluster"] == cid]

            texts = subset[self.text_column].tolist()

            keywords = self.extract_keywords(texts)

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

            theme = ", ".join(info["keywords"][:2])

            subset = self.df[self.df["cluster"] == cid]

            avg_sentiment = subset["sentiment"].mean()

            pct = round((info["count"] / total) * 100, 2)

            severity = info["count"] * abs(avg_sentiment)

            rows.append({

                "Theme": theme,
                "Response_Count": info["count"],
                "Percentage (%)": pct,
                "Average Sentiment": round(avg_sentiment, 3),
                "Severity Score": round(severity, 2),
                "Top Keywords": ", ".join(info["keywords"]),
                "Example Quote": info["samples"][0]

            })

        summary_df = pd.DataFrame(rows)

        summary_df = summary_df.sort_values(
            "Severity Score",
            ascending=False
        )

        self.summary_df = summary_df

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

        self.summary_df["Recommendation"] = recs

    #-------------------------------------------------
    # INSIGHT GENERATION
    #-------------------------------------------------

    def generate_key_insights(self):

        insights = []

        top = self.summary_df.iloc[0]

        insights.append(
            f"The most common issue reported by respondents is '{top['Theme']}', "
            f"mentioned in {top['Percentage (%)']}% of responses."
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

        for score in self.summary_df["Severity Score"]:

            if score > 20:
                priorities.append("High Priority")

            elif score > 10:
                priorities.append("Medium Priority")

            else:
                priorities.append("Low Priority")

        self.summary_df["Priority"] = priorities

    #-------------------------------------------------
    # DATASET SUMMARY
    #-------------------------------------------------

    def generate_dataset_summary(self):

        total = len(self.df)

        avg_length = self.df[self.text_column].apply(lambda x: len(str(x).split())).mean()

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
            theme_map[cid] = ", ".join(info["keywords"][:2])

        self.df["Theme"] = self.df["cluster"].map(theme_map)

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

        # EXECUTIVE SUMMARY

        ws = wb.create_sheet("Executive Summary")

        ws["A1"] = "Customer Feedback Insights Report"
        ws["A1"].font = Font(size=20, bold=True)

        ws.merge_cells("A1:E1")

        # DATASET OVERVIEW

        ws["A3"] = "Total Responses"
        ws["B3"] = self.dataset_summary["Total Responses"]

        ws["A4"] = "Average Response Length"
        ws["B4"] = self.dataset_summary["Average Response Length"]

        ws["A5"] = "Themes Identified"
        ws["B5"] = self.dataset_summary["Themes Identified"]

        # TOP INSIGHTS

        ws["A7"] = "Top Issue"
        ws["B7"] = self.summary_df.iloc[0]["Theme"]

        ws["A8"] = "Most Severe Issue"
        ws["B8"] = self.summary_df.iloc[0]["Theme"]

        # KEY INSIGHTS SECTION

        ws["A10"] = "Key Insights"
        ws["A10"].font = Font(bold=True)

        row = 11

        for insight in self.key_insights:
            ws[f"A{row}"] = f"• {insight}"
            row += 1

        self.auto_width(ws)

        # THEME ANALYSIS

        ws = wb.create_sheet("Theme Analysis")

        ws.append(self.summary_df.columns.tolist())

        for r in self.summary_df.itertuples(index=False):
            ws.append(list(r))

        self.style_header(ws)
        self.zebra_rows(ws)

        ws.auto_filter.ref = ws.dimensions

        self.auto_width(ws)

        # CHARTS

        ws = wb.create_sheet("Charts")

        ws.append(["Theme", "Responses"])

        for r in self.summary_df[["Theme", "Response_Count"]].itertuples(index=False):
            ws.append(r)

        data = Reference(ws, min_col=2, min_row=1, max_row=len(self.summary_df) + 1)
        cats = Reference(ws, min_col=1, min_row=2, max_row=len(self.summary_df) + 1)

        bar = BarChart()
        bar.title = "Feedback Distribution by Theme"
        bar.y_axis.title = "Responses"

        bar.add_data(data, titles_from_data=True)
        bar.set_categories(cats)

        ws.add_chart(bar, "D2")

        pie = PieChart()
        pie.title = "Theme Share"

        pie.add_data(data, titles_from_data=True)
        pie.set_categories(cats)

        ws.add_chart(pie, "D20")

        # EXAMPLE QUOTES

        ws = wb.create_sheet("Example Quotes")

        ws.append(["Theme", "Quote"])

        for cid, info in self.cluster_summary.items():

            theme = ", ".join(info["keywords"][:2])

            for q in info["samples"]:
                ws.append([theme, q])

        self.style_header(ws)
        self.zebra_rows(ws)

        self.auto_width(ws)

        # FULL DATA

        ws = wb.create_sheet("Clustered Responses")

        ws.append(self.df.columns.tolist())

        for r in self.df.itertuples(index=False):
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

        print("Cleaning text...")
        self.preprocess()

        print("Sentiment analysis...")
        self.compute_sentiment()

        print("Generating embeddings...")
        self.create_embeddings()

        print("Clustering responses...")
        self.cluster_feedback()

        print("Building cluster summary...")
        self.build_cluster_summary()

        print("Generating insights...")
        self.build_summary_table()

        print("Adding priority labels...")
        self.add_priority_labels()

        print("Generating recommendations...")
        self.generate_recommendations()

        print("Generating key insights...")
        self.generate_key_insights()

        print("Generating dataset summary...")
        self.generate_dataset_summary()

        print("Labeling responses...")
        self.label_data()

        print("Building Excel report...")
        self.build_excel_report()

        print("Analysis complete.")

# In[2]

engine = FeedbackInsightEngine(
"/home/churchil/Desktop/pythonfiles/env310/CustomerFeedbackInsight/health_survey_responses.csv",
"response"
)

engine.run()
# %%
