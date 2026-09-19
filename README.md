# Weekly Reporting Automation Demo

A small Streamlit proof of concept for a weekly reporting workflow.

The demo is designed to show how a Monday reporting pack could be prepared automatically after a Saturday reporting cycle closes. It uses synthetic data only.

## What the demo does

- Keeps preloaded weekly Excel files in the app. The user does not upload anything.
- Lets the user select a reporting week.
- Validates each weekly file.
- Combines historical weeks.
- Finds the previous available reporting week automatically.
- Calculates KPIs and week-on-week changes.
- Generates a short management summary.
- Uses OpenRouter for the wording when an API key is configured.
- Falls back to a deterministic summary if the API is unavailable.
- Generates a real `.pptx` presentation.
- Explains how the proof of concept maps to SharePoint, Power BI, Power Automate, Copilot, PowerPoint and Teams.

## Project structure

```text
weekly_reporting_automation_demo/
├── app.py
├── reporting.py
├── requirements.txt
├── .gitignore
├── .streamlit/
│   └── secrets.toml.example
└── data/
    ├── StoreActivity_2026-08-29.xlsx
    ├── StoreActivity_2026-09-05.xlsx
    ├── StoreActivity_2026-09-12.xlsx
    └── StoreActivity_2026-09-19.xlsx
```

## Run locally

Create a virtual environment if you want one, then install dependencies:

```bash
pip install -r requirements.txt
```

Start the app:

```bash
streamlit run app.py
```

## Enable OpenRouter summaries

The app works without an API key. Without one, it uses a built-in summary so the demo cannot fail during an interview.

To enable OpenRouter locally:

1. Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml`.
2. Put your real key in the new file.
3. Do not commit `secrets.toml`.

Example:

```toml
OPENROUTER_API_KEY = "your-key-here"
OPENROUTER_MODEL = "openai/gpt-4o-mini"
```

You can change `OPENROUTER_MODEL` to any model available to your OpenRouter account.

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload the project files.
3. Deploy `app.py` on Streamlit Community Cloud.
4. Add the OpenRouter values in the app's Secrets settings instead of committing the real key to GitHub.

## Suggested interview explanation

> I built this proof of concept outside the Microsoft environment because I do not have access to the company's tenant. The preloaded Excel files represent a SharePoint reporting folder. The app preserves historical weekly data, validates the selected reporting period, calculates KPIs, creates a short management summary and generates a PowerPoint. In the production environment, I would map those steps to SharePoint, Power Query and Power BI, Power Automate, Copilot, PowerPoint and Teams.

## Production mapping

| Demo | Microsoft environment |
|---|---|
| Preloaded Excel folder | SharePoint document library |
| Python data loading | Power Query |
| KPI calculations | Power BI semantic model / DAX |
| Validation logic | Power Automate + Power Query |
| Dashboard | Power BI |
| OpenRouter summary | Copilot |
| PPTX generation | Power BI / Power Automate / PowerPoint |
| Demo status messages | Teams notifications |

## Important design principle

The language model is used only for wording. All KPIs and comparisons are calculated first from validated data.
