# Academic Personnel Validator

Validates filled Excel files against the `AcademicPersonel_update.xlsx` template rules.

## Setup

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run validator_app.py
```

Then open http://localhost:8501 in your browser.

## How it works

1. Upload a client-filled `.xlsx` file from the sidebar
2. Tick/untick which columns you want to validate
3. Click **Run Validation**
4. Review the summary dashboard, color-coded table, and per-row error details

## Validation rules enforced

| Column | Rule |
|--------|------|
| A–C, G–J, O | Required text fields |
| D | F or M only |
| E | Integer, age 18–80 |
| F | Diploma / Bachelor / Master / Director / Doctorate / Professional |
| I | 33 allowed subjects (from Sheet2 dropdown) |
| K | Yes or No |
| L | Valid date |
| M | Career ladder values |
| N | Private / Public / Government / Unknown |
| P | School level values |
| Q | One of 11 Addis Ababa sub-cities |
| R | Licence type (optional) |
| S | Valid date (optional) |
| T | 9-digit mobile number |
| U | Disability category |
| V | Activity type |
| W | Non-negative number (optional) |