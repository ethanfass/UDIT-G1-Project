# Project Epics & Tasks

## Epic 1: Policy Standards & Rubric
**Assigned to:** Kevin

### Overview
Create a reliable and consistent set of policy standards and a rubric based on these standards to score inputs.

### Tasks
1. **Understand policy standards**
   - Research and document existing policies
   - Identify key compliance areas
   - Document policy hierarchy and relationships

2. **Assemble a rubric for report generation**
   - Create scoring criteria (e.g., 1-5 scale)
   - Define weightings for different policy areas
   - Build template for consistent evaluation

3. **Set up a RAG based on policy standards**
   - Implement Retrieval-Augmented Generation
   - Index policy documents
   - Enable context-aware policy lookup

---

## Epic 2: Data Filtering & Combination
**Assigned to:** Shukria, Zahra & Matthew

### Overview
Create a way to filter data into one input.

### Tasks
1. **Search functionality**
   - Build search interface for questionnaires
   - Filter by date, category, or other criteria
   - Implement sorting and organization

2. **Write initial Python code to combine Excel documents**
   - Normalize data across questionnaires
   - Handle missing/inconsistent fields
   - Generate combined output (Excel/JSON/CSV)

---

## Epic 3: AI Analysis Implementation
**Assigned to:** Ethan

### Overview
Implement an AI to analyze our data input based on set standards.

### Tasks
1. **Set up basic AI implementation**
   - Integrate OpenAI API
   - Build analysis pipeline
   - Test with sample data
   - Generate scored reports

---

## Project Workflow

```
1. Manually Create Template
   └─> Analyze sample questionnaires
   └─> Define standard format

2. Convert to Template (AI)
   └─> Input raw questionnaires
   └─> AI converts to template format
   └─> Validate output

3. Combine Questionnaires
   └─> Select related templates
   └─> Merge into single document
   └─> Validate combined data

4. Analyze & Score (AI)
   └─> Apply policy rubric
   └─> Generate scoring report
   └─> Provide recommendations
```
