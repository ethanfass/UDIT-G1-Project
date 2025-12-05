from google import genai
import pandas as pd

# 1. Initialize Gemini client (make sure your API key is set in env or config)
client = genai.Client()

# 2. Load the 3 filled-out questionnaires for Grammarly
file_paths = [
    "UDIT-G1-Project\Assets\Grammarly VSA Questionnaire 2024.xlsx",
    "UDIT-G1-Project\Assets\HECVAT303-Grammarly 2023.xlsx",
    "UDIT-G1-Project\Assets\SIG Core - 2025.xlsx",
]

questionnaire_csvs = []
for path in file_paths:
    # If there are multiple sheets you care about, you may need to loop over them.
    df = pd.read_excel(path)
    questionnaire_csvs.append(df.to_csv(index=False))

q1_csv, q2_csv, q3_csv = questionnaire_csvs

# 3. Build the Gemini prompt: ask ONLY for a generic master template (no Grammarly-specific values)
prompt = f"""
You are a senior cybersecurity risk and compliance analyst.

You are given three filled-out cybersecurity questionnaires for the SAME vendor (Grammarly),
but your job is NOT to evaluate Grammarly. Instead, your task is to DESIGN a single, generic
MASTER CYBERSECURITY DUE DILIGENCE TEMPLATE that could be reused for ANY vendor.

The three questionnaires may:
- Have different structures and column names
- Ask overlapping or duplicate questions in different wording
- Have vendor-specific answers already filled in

Your job is to:
1. Infer the full set of SECURITY TOPICS and CONTROLS that are covered across ALL THREE questionnaires.
2. Normalize and deduplicate the questions into a clean, consistent structure.
3. Design a vendor-agnostic template that captures all the necessary information to assess whether
   a company meets common security standards and guidelines (e.g., SOC 2, ISO 27001, NIST CSF, etc.).

VERY IMPORTANT:
- Completely IGNORE the specific answers for Grammarly and any vendor-specific details.
- Focus ONLY on the QUESTIONS / FIELDS / CONTROL AREAS that a generic template should contain.
- The output must be reusable for any future vendor.

### Output format

Return ONLY **one CSV table** with a header row, no additional commentary, in the following columns
(you can adjust slightly if needed, but keep it structured):

Section,
Subsection,
Control_ID,
Control_Name,
Question_Text,
Answer_Type,
Expected_Response_Format,
Response_Options,
Evidence_Requested,
Criticality_Level,
Standard_Mapping,
Notes

Where:
- Section: high-level domain (e.g., Governance, Access Control, Network Security, Application Security, Incident Response, Business Continuity/DR, Physical Security, Privacy & Data Protection, Third-Party Risk, etc.)
- Subsection: more specific grouping within a section (e.g., Password Policy, MFA, Logging & Monitoring)
- Control_ID: a simple identifier you create (e.g., AC-01, AC-02, IR-01)
- Control_Name: short title for the control/question (e.g., "Multi-Factor Authentication for Admins")
- Question_Text: the actual question to ask the vendor in the template
- Answer_Type: e.g., "Yes/No", "Multiple Choice", "Free Text", "Numeric", "Attachment"
- Expected_Response_Format: brief guidance for how the vendor should answer (e.g., "Describe...", "Provide policy name...", "Upload document...")
- Response_Options: if applicable, list common options (e.g., "Yes/No", or a small set like "In-scope, Out-of-scope, Planned")
- Evidence_Requested: examples of evidence to attach (e.g., "Policy document", "Screenshot", "PenTest report", "SOC 2 Type 2 report")
- Criticality_Level: e.g., "High", "Medium", "Low"
- Standard_Mapping: if you can infer it, note mappings like "ISO 27001 A.9.2.3; SOC 2 CC6.3; NIST CSF PR.AC-7"
- Notes: any short internal note to the assessor (e.g., "Key control for privileged access management")

The goal is:
- Combine EVERYTHING covered in the three questionnaires
- Remove duplicates and near-duplicates (merge them into a single well-written question when possible)
- Make the template as clear and organized as possible, suitable for grading/assessing future vendors.

### Input data

Below are the three filled-out questionnaires as CSV. Use them ONLY to infer what the template
should contain; do NOT reproduce the vendor's specific answers.

=== QUESTIONNAIRE 1 (CSV) ===
{q1_csv}

=== QUESTIONNAIRE 2 (CSV) ===
{q2_csv}

=== QUESTIONNAIRE 3 (CSV) ===
{q3_csv}

Remember: Output ONLY the final unified template as a CSV table with the specified columns.
Do not include any explanation or markdown, only the raw CSV content.
"""

# 4. Call Gemini
resp = client.models.generate_content(
    model="gemini-2.5-flash",
    contents=prompt,
)

# 5. Save the returned template to a CSV file
template_csv = resp.text

with open("master_cyber_template.csv", "w", encoding="utf-8") as f:
    f.write(template_csv)

print("Template saved to master_cyber_template.csv")

with open("master_cyber_template.csv", "r") as f:
    master_cyber_template = f.read()


# 6. Prompt AI to assess the risk
prompt = f"""
You are a senior cybersecurity risk and compliance analyst.

You are given a combined spreadsheat of cybersecurity questionnaires for a vendor and must grade each questions answer based on from 1-5, 5 being the safest.
The grade for each question should be based on the how likely the answer is to be a risk and how bad it would be for that risk to happen.

Your job is to:
1. Assess the security risk of each answer
2. Give an overall grade
3. Point out major flaws in answers

VERY IMPORTANT:
- Analyze each question
- Maximum points should be possible if there is no security risk
- Have the overall grade at the bottom, 2 rows under the section.
- Next to the overall grade, in the 2 rows under the subsection, point out any major security risks you find.

### Output format

Return ONLY **one CSV table** with a header row, no additional commentary, in the following columns
(you can adjust slightly if needed, but keep it structured):

Section,
Subsection,
Control_ID,
Control_Name,
Question_Text,
Answer_Type,
Grade,
Expected_Response_Format,
Response_Options,
Evidence_Requested,
Criticality_Level,
Standard_Mapping,
Notes

Where:
- Section: high-level domain (e.g., Governance, Access Control, Network Security, Application Security, Incident Response, Business Continuity/DR, Physical Security, Privacy & Data Protection, Third-Party Risk, etc.)
- Subsection: more specific grouping within a section (e.g., Password Policy, MFA, Logging & Monitoring)
- Control_ID: a simple identifier you create (e.g., AC-01, AC-02, IR-01)
- Control_Name: short title for the control/question (e.g., "Multi-Factor Authentication for Admins")
- Question_Text: the actual question to ask the vendor in the template
- Answer_Type: e.g., "Yes/No", "Multiple Choice", "Free Text", "Numeric", "Attachment"
- Grade: 1-5 analysis of how safe the answer is to the question
- Expected_Response_Format: brief guidance for how the vendor should answer (e.g., "Describe...", "Provide policy name...", "Upload document...")
- Response_Options: if applicable, list common options (e.g., "Yes/No", or a small set like "In-scope, Out-of-scope, Planned")
- Evidence_Requested: examples of evidence to attach (e.g., "Policy document", "Screenshot", "PenTest report", "SOC 2 Type 2 report")
- Criticality_Level: e.g., "High", "Medium", "Low"
- Standard_Mapping: if you can infer it, note mappings like "ISO 27001 A.9.2.3; SOC 2 CC6.3; NIST CSF PR.AC-7"
- Notes: any short internal note to the assessor (e.g., "Key control for privileged access management")

The goal is:
- Assess the risk of trusting a company based on their answers to these questions.

### Input data

Below is the questionnaire CSV.

=== COMBINED QUESTIONNAIRE ===
{master_cyber_template}


Remember: Output ONLY the final unified template as a CSV table with the specified columns.
Do not include any explanation or markdown, only the raw CSV content.
"""

# 7. Call Gemini
resp2 = client.models.generate_content(
    model="gemini-2.5-flash",
    contents=prompt,
)

# 8. Save the returned  to a CSV file
assessment_csv = resp2.text

with open("security_assessment.csv", "w", encoding="utf-8") as f:
    f.write(assessment_csv)

print("Template saved to security_assessment.csv")