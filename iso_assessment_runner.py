from __future__ import annotations

import os
import argparse
import json
import textwrap
from typing import List, Dict, Any

pd = None
genai = None



# Load questionnaire files

def load_questionnaire_files(paths: List[str]) -> str:
    """
    Reads questionnaire files and combines them into one large text string.

    Supports:
      - Excel (.xlsx / .xls)
      - CSV (.csv)
      - Text (.txt / .md)

    Unsupported types are skipped.
    """
    texts = []

    for p in paths:
        ext = os.path.splitext(p)[1].lower()

        if not os.path.exists(p):
            print(f"[WARN] File not found, skipping: {p}")
            continue

        try:
            # Excel files
            if ext in (".xlsx", ".xls"):
                xls = pd.ExcelFile(p)
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(p, sheet_name=sheet_name)
                    texts.append(
                        f"=== FILE: {p} | SHEET: {sheet_name} ===\n"
                        + df.to_csv(index=False)
                    )

            # CSV files
            elif ext == ".csv":
                df = pd.read_csv(p)
                texts.append(f"=== FILE: {p} ===\n" + df.to_csv(index=False))

            # Plain text files
            elif ext in (".txt", ".md"):
                with open(p, "r", encoding="utf-8", errors="ignore") as f:
                    texts.append(f"=== FILE: {p} ===\n" + f.read())

            else:
                print(f"[WARN] Unsupported file type: {p}")

        except Exception as e:
            print(f"[WARN] Failed to load {p}: {e}")

    big_text = "\n\n".join(texts)

    if not big_text.strip():
        raise ValueError("No questionnaire content loaded.")

    return big_text



# Extract template questions

def extract_template_questions(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """
    Converts template rows into a simplified JSON structure.
    """
    required_cols = [
        "Section",
        "Subsection",
        "Control_ID",
        "Control_Name",
        "Question_Text",
        "Criticality_Level",
    ]

    # Make sure required columns exist
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing column: {col}")

    questions = []

    for _, row in df.iterrows():
        control_id = str(row["Control_ID"]).strip()
        if not control_id:
            continue

        questions.append(
            {
                "control_id": control_id,
                "section": str(row["Section"]),
                "subsection": str(row["Subsection"]),
                "control_name": str(row["Control_Name"]),
                "question": str(row["Question_Text"]),
                "criticality": int(row["Criticality_Level"]),
            }
        )

    return questions



# Build Gemini prompt

def build_assessment_prompt(
    company_name: str,
    template_questions: List[Dict[str, Any]],
    questionnaires_text: str,
) -> str:
    """
    Builds the full prompt string sent to Gemini.
    """
    questions_json = json.dumps(template_questions, ensure_ascii=False, indent=2)

    instructions = textwrap.dedent(
        f"""
        You are an ISO 27001 / ISO 27002 information security assessor.

        Company: "{company_name}"

        You are given:
        1) A JSON list of ISO-aligned questions.
        2) Raw questionnaire responses from the company.

        For each control:
          - Review the question.
          - Use only the questionnaire text as evidence.
          - Write a short answer summary.
          - Assign a score from 0 to criticality.
          - Provide a short rationale.

        Scoring:
          0 = no evidence
          partial = incomplete implementation
          max = strong implementation

        Output ONLY valid JSON in this format:

        {{
          "<CONTROL_ID>": {{
              "answer": "...",
              "score": <int>,
              "max_score": <int>,
              "rationale": "..."
          }}
        }}

        Do not include markdown.
        """
    ).strip()

    prompt = (
        instructions
        + "\n\n=== TEMPLATE QUESTIONS ===\n"
        + questions_json
        + "\n\n=== QUESTIONNAIRE TEXT ===\n"
        + questionnaires_text
    )

    return prompt


# Call Gemini and parse JSON

def call_gemini_and_get_scores(
    api_key: str,
    model_name: str,
    prompt: str,
) -> Dict[str, Any]:
    """
    Sends prompt to Gemini and parses the JSON response.
    """
    client = genai.Client(api_key=api_key)

    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
    )

    raw_text = getattr(response, "text", None)

    if raw_text is None:
        try:
            raw_text = response.candidates[0].content.parts[0].text
        except Exception as e:
            raise RuntimeError(f"Could not read Gemini response: {e}")

    cleaned = raw_text.strip()

    if cleaned.startswith("```"):
        cleaned = cleaned.strip("`")
        if cleaned.startswith("json"):
            cleaned = cleaned[len("json"):].strip()

    try:
        scores = json.loads(cleaned)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON from Gemini: {e}")

    if not isinstance(scores, dict):
        raise ValueError("Gemini response is not a JSON object.")

    return scores



# Fill assessment sheet

def fill_assessment_from_scores(
    template_df: pd.DataFrame,
    assessment_df: pd.DataFrame,
    scores: Dict[str, Any],
) -> Dict[str, float]:
    """
    Fills the assessment sheet with Gemini scores and answers.
    """
    earned_points = 0.0
    total_points = 0.0

    assessment_df["Answer"] = assessment_df["Answer"].astype("object")
    assessment_df["Total_Score_Possible"] = assessment_df["Total_Score_Possible"].astype(float)
    assessment_df["Score_Earned"] = assessment_df["Score_Earned"].astype(float)

    for idx, t_row in template_df.iterrows():
        control_id = str(t_row["Control_ID"]).strip()
        max_points = float(t_row["Criticality_Level"])
        total_points += max_points

        info = scores.get(control_id, {})

        answer_text = str(
            info.get("answer")
            or info.get("rationale")
            or ""
        ).strip()

        try:
            score = float(info.get("score", 0))
        except (TypeError, ValueError):
            score = 0.0

        score = max(0.0, min(score, max_points))

        assessment_df.at[idx, "Answer"] = answer_text
        assessment_df.at[idx, "Total_Score_Possible"] = max_points
        assessment_df.at[idx, "Score_Earned"] = score

        earned_points += score

    grade_percent = (earned_points / total_points) * 100.0 if total_points > 0 else 0.0

    return {
        "earned_points": earned_points,
        "total_points": total_points,
        "grade_percent": grade_percent,
    }



# Main execution

def main():
    parser = argparse.ArgumentParser(
        description="Fill and grade ISO assessment using Gemini."
    )

    parser.add_argument("--template", required=True)
    parser.add_argument("--assessment_blank", required=True)
    parser.add_argument("--questionnaires", nargs="+", required=True)
    parser.add_argument("--company", required=True)
    parser.add_argument("--output_assessment", default="assessment_filled.xlsx")
    parser.add_argument("--model", default="models/gemini-2.5-flash")
    parser.add_argument("--batch_size", type=int, default=40)

    args = parser.parse_args()

    # Get API key from environment
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not api_key:
        raise EnvironmentError("Set GEMINI_API_KEY or GOOGLE_API_KEY.")

    template_df = pd.read_excel(args.template)
    assessment_df = pd.read_excel(args.assessment_blank, sheet_name="Assessment")

    questionnaires_text = load_questionnaire_files(args.questionnaires)

    total_rows = len(template_df)
    batch_size = max(1, args.batch_size)

    earned_total = 0.0
    max_total = 0.0

    for start in range(0, total_rows, batch_size):
        end = min(start + batch_size, total_rows)

        t_batch = template_df.iloc[start:end].reset_index(drop=True).copy()
        a_batch = assessment_df.iloc[start:end].reset_index(drop=True).copy()

        template_questions = extract_template_questions(t_batch)

        prompt = build_assessment_prompt(
            company_name=args.company,
            template_questions=template_questions,
            questionnaires_text=questionnaires_text,
        )

        scores = call_gemini_and_get_scores(
            api_key=api_key,
            model_name=args.model,
            prompt=prompt,
        )

        score_info = fill_assessment_from_scores(t_batch, a_batch, scores)

        assessment_df.loc[start:end - 1, "Answer"] = a_batch["Answer"].values
        assessment_df.loc[start:end - 1, "Total_Score_Possible"] = a_batch["Total_Score_Possible"].values
        assessment_df.loc[start:end - 1, "Score_Earned"] = a_batch["Score_Earned"].values

        earned_total += score_info["earned_points"]
        max_total += score_info["total_points"]

    final_grade = (earned_total / max_total) * 100.0 if max_total > 0 else 0.0

    with pd.ExcelWriter(args.output_assessment, engine="openpyxl") as writer:
        assessment_df.to_excel(writer, sheet_name="Assessment", index=False)

    print(
        f"{args.company}: {earned_total:.1f}/{max_total:.1f} "
        f"({final_grade:.1f}%)"
    )


from assessment_runner_core import main as modern_main


if __name__ == "__main__":
    modern_main()
