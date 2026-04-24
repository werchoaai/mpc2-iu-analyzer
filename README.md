# MPC² I-U Review Workstation (Projekt 3)

Streamlit app for triaging and reviewing I-U (Stromdichte-Potenzial) corrosion
measurement files. Classifies curve families, suggests OCP / RPP / E_pit,
and persists human overrides with audit trail.

**Status:** v0.1 — Review Workstation (triage + override + Excel export).
**Not yet built:** auto-report DOCX writer (deferred to v2).

## Run locally
```
pip install -r requirements.txt
streamlit run app.py
```

## Deploy
Streamlit Cloud. Python pinned via `runtime.txt` (3.11).
Set app secret `password` for the login gate.
