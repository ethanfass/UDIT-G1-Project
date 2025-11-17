# Standard Questionnaire Template

## Overview
This document outlines the standard template format for all questionnaires in the UDIT-G1 project.

## Template Structure

### Header Information
```
Title: [Questionnaire Title]
Date: [MM/DD/YYYY]
Organization: [Organization Name]
Respondent: [Name/ID]
Category: [Category Classification]
```

### Section 1: Organizational Information
- Organization Name
- Department/Division
- Contact Person
- Contact Email
- Contact Phone

### Section 2: Policy Compliance Assessment
For each policy area:
- Policy Area Name
- Current Implementation Status (Not Started / In Progress / Complete)
- Compliance Level (Non-Compliant / Partially Compliant / Fully Compliant)
- Evidence/Documentation
- Comments

### Section 3: Risk Assessment
- Risk Areas Identified
- Risk Level (Low / Medium / High)
- Mitigation Strategies
- Timeline for Resolution

### Section 4: Additional Comments
- General feedback
- Challenges faced
- Resource needs
- Next steps

### Section 5: Signatures
- Respondent Name & Date
- Reviewer Name & Date (if applicable)

---

## Data Format (For AI Conversion)

When converted to standardized format, use this JSON structure:

```json
{
  "metadata": {
    "title": "string",
    "date": "YYYY-MM-DD",
    "organization": "string",
    "respondent": "string",
    "category": "string"
  },
  "organizational_info": {
    "organization_name": "string",
    "department": "string",
    "contact_person": "string",
    "email": "string",
    "phone": "string"
  },
  "policy_assessment": [
    {
      "policy_area": "string",
      "implementation_status": "Not Started|In Progress|Complete",
      "compliance_level": "Non-Compliant|Partially Compliant|Fully Compliant",
      "evidence": "string",
      "comments": "string"
    }
  ],
  "risk_assessment": [
    {
      "risk_area": "string",
      "risk_level": "Low|Medium|High",
      "mitigation": "string",
      "timeline": "string"
    }
  ],
  "additional_comments": "string"
}
```

---

## Notes
- This template should be adapted based on your specific policy standards
- All dates should use ISO 8601 format (YYYY-MM-DD)
- Compliance levels should be consistent across all assessments
- Template will be refined as feedback is collected
