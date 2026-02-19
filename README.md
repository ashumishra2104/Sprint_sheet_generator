# 🚀 Sprint Report Generator

A Streamlit web app that takes a **Jira CSV export** and generates a fully formatted **Excel Sprint Report** — complete with KPIs, status breakdowns, and a full Epic → Story/Task/Bug → Sub-task hierarchy.

---

## ✨ Features

- 3-step wizard: Sprint Details → Upload CSV → Download Report
- Auto-calculates Total Days, Days Left, Sprint End Date
- Parses full Jira hierarchy (Epics, Stories, Tasks, Bugs, Sub-tasks)
- Handles **external Epics** (parent not in CSV export) — no items lost
- Handles **standalone/unlinked items** — grouped at the bottom
- Colour-coded status cells and priority in the Excel output
- Clickable Jira hyperlinks in the generated Excel
- KPI summary block: Pending %, Production %, status counts

---

## 🗂️ Project Structure

```
sprint-report-generator/
├── app.py                  # Main Streamlit app (3-step UI)
├── requirements.txt        # Python dependencies
├── modules/
│   ├── __init__.py
│   ├── parser.py           # Jira CSV → hierarchy + KPI logic
│   └── excel_generator.py  # openpyxl Excel builder
```

---

## 🚀 Running Locally

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/sprint-report-generator.git
cd sprint-report-generator
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run
```bash
streamlit run app.py
```

App will open at **http://localhost:8501**

---

## ☁️ Deploy on Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select your repo → set `app.py` as the main file
4. Click **Deploy** — done!

---

## 📋 Jira CSV Requirements

Export your Jira board via: **Board → Export Issues → CSV (all fields)**

**Required columns:**
| Column | Description |
|--------|-------------|
| `Issue key` | Unique ticket ID (e.g. PB-1234) |
| `Issue Type` | Epic / Story / Task / Bug / Sub-task |
| `Summary` | Ticket title |
| `Status` | Current status |
| `Parent key` | Parent ticket's Issue key |

**Recommended columns:** `Priority`, `Assignee`, `Custom field (Target start)`, `Custom field (Target end)`

---

## 🎨 Status Colour Mapping

| Status | Colour |
|--------|--------|
| Not Initiated / To Do / Open | 🟠 Orange |
| In Progress | 🔵 Blue |
| Staging | 🟡 Yellow |
| QA Review / In Review | 🟡 Amber |
| QA Deployed | 🟢 Light Green |
| QA Approved | 🟢 Green |
| Done / Production / Released | 🟢 Dark Green |
| On Hold / Blocked | ⚪ Grey |
| To Be Picked In Another Sprint | 🟣 Purple |
