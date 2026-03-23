# Automate Google Workspace with AI | Powered by Cortex Code
# CoCo Google Productiviy Suite Skill

Automate Google Workspace (Slides, Sheets, Docs, Drive, Forms) programmatically using AI. Generate professional presentations, reports, spreadsheets, and documents with natural language prompts.

## What This Skill Does

This skill enables Cortex Code to:

- **Create presentations** — Generate polished Google Slides with themes, charts, and layouts
- **Build spreadsheets** — Create, read, and write Google Sheets with formatting and formulas
- **Write documents** — Generate Google Docs with structured content and formatting
- **Manage files** — Upload, download, search, and share files in Google Drive
- **Create forms** — Build surveys and quizzes with Google Forms, linked to Sheets

## Business Outcomes

| Outcome | How It's Achieved |
|---------|-------------------|
| **Save hours on presentations** | Generate complete slide decks from bullet points or data |
| **Automate reporting** | Pull data from databases → format in Sheets → create slides |
| **Standardize branding** | Apply consistent themes across all presentations |
| **Eliminate manual work** | Batch create documents, forms, and spreadsheets |
| **Version control content** | Track all changes with HISTORY.md for rollback |

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                    Google Productivity Automation                                │
│                                                                                  │
│  ┌──────────────┐                                                               │
│  │   Cortex     │   "Create a presentation about Q1 results"                    │
│  │    Code      │────────────────────────────────────────────────────────────┐  │
│  └──────────────┘                                                            │  │
│         │                                                                    │  │
│         │ Reads SKILL.md                                                     │  │
│         ▼                                                                    │  │
│  ┌──────────────┐                                                            │  │
│  │   Skill      │   • API reference                                          │  │
│  │  Knowledge   │   • Bundled themes (Corporate, Professional)               │  │
│  │              │   • Helper library source                                   │  │
│  └──────────────┘   • Working examples                                       │  │
│         │                                                                    │  │
│         │ Generates Python script                                            │  │
│         ▼                                                                    │  │
│  ┌──────────────┐     ┌──────────────┐     ┌──────────────────────────────┐ │  │
│  │   pixi run   │────▶│   Google     │────▶│  Google Workspace            │ │  │
│  │   python     │     │   APIs       │     │                              │ │  │
│  │   script.py  │     │   (REST)     │     │  📊 Slides  📑 Sheets        │ │  │
│  └──────────────┘     └──────────────┘     │  📝 Docs    📁 Drive         │ │  │
│                                            │  📋 Forms                     │ │  │
│                                            └──────────────────────────────┘ │  │
└─────────────────────────────────────────────────────────────────────────────────┘
```

## Quick Start

### 1. Prerequisites

```bash
# Install pixi (package manager)
curl -fsSL https://pixi.sh/install.sh | bash

# Install gcloud CLI
brew install --cask google-cloud-sdk  # macOS

# Authenticate
gcloud auth application-default login

# Set quota project (replace with your GCP project)
gcloud auth application-default set-quota-project your-gcp-project
```

### 2. Create a Project Folder

```bash
mkdir my_presentation && cd my_presentation
pixi init --format pyproject
pixi add python=3.11
pixi add --pypi google-api-python-client google-auth google-auth-oauthlib
pixi install
```

### 3. Ask Cortex Code

```
"Create a presentation about Q1 sales results with 5 slides"

"Make a spreadsheet tracking project milestones with conditional formatting"

"Generate a customer feedback form with rating questions"

"Create a technical document explaining our API architecture"
```

## Supported Google APIs

| API | Capabilities |
|-----|--------------|
| **Google Slides** | Create presentations, add slides, shapes, text boxes, images, charts, tables |
| **Google Sheets** | Create/read/write spreadsheets, format cells, create charts, use formulas |
| **Google Docs** | Create documents, insert text, format paragraphs, add tables, images |
| **Google Drive** | List, search, upload, download, share files and folders |
| **Google Forms** | Create forms/quizzes, add questions (multiple choice, text, scale), get responses |

## Project Structure

```
my_presentation/
├── pixi.toml              # Dependencies
├── pixi.lock              # Locked versions
├── HISTORY.md             # Change log for rollback
├── scripts/
│   ├── 001_create_presentation.py
│   ├── 002_add_content_slides.py
│   └── 003_add_charts.py
└── assets/
    └── logo.png
```

## Example Prompts

### Presentations

```
"Create a professional presentation with title slide and 3 content slides about data architecture"

"Add a comparison table slide to presentation ID xyz123 comparing AWS vs Azure vs GCP"

"Insert a bar chart on slide 4 showing quarterly revenue data"
```

### Spreadsheets

```
"Create a budget tracking spreadsheet with monthly columns and category rows"

"Read data from the Sales sheet and create a summary pivot table"

"Add conditional formatting to highlight cells over $10,000 in red"
```

### Documents

```
"Create a technical specification document for our REST API"

"Generate meeting notes template with sections for attendees, agenda, action items"

"Create a project proposal document from these bullet points: ..."
```

### Forms

```
"Create a customer satisfaction survey with 5 rating questions"

"Build a quiz about product fundamentals with 10 multiple choice questions"

"Create an event registration form collecting name, email, and dietary preferences"
```

## Bundled Themes

The skill includes pre-built themes for professional presentations:

| Theme | Description |
|-------|-------------|
| **Corporate Clean** | Professional gray/blue theme for business presentations |
| **Technical Dark** | Dark theme optimized for code and diagrams |
| **Modern Light** | Clean white theme with accent colors |

## File Structure

```
google_productivity_skill/
├── README.md                           # This file
├── google_productivity_skill/
│   └── SKILL.md                        # Full skill knowledge base (~2800 lines)
└── demo/
    ├── 01_create_presentation.py       # Create a themed presentation
    ├── 02_create_spreadsheet.py        # Create and populate a spreadsheet
    ├── 03_create_document.py           # Create a formatted document
    ├── 04_create_form.py               # Create a survey form
    └── 05_data_to_slides.py            # Pull database data → Slides
```

## Authentication

### Setup

1. Create a GCP project at https://console.cloud.google.com/projectcreate
2. Enable APIs: Slides, Sheets, Docs, Drive, Forms
   ```bash
   gcloud services enable slides.googleapis.com sheets.googleapis.com docs.googleapis.com drive.googleapis.com forms.googleapis.com --project=YOUR_PROJECT_ID
   ```
3. Authenticate and set quota project:
   ```bash
   gcloud auth application-default login
   gcloud auth application-default set-quota-project YOUR_PROJECT_ID
   ```

## Common Errors

| Error | Solution |
|-------|----------|
| `403 "API has not been used in project 764086051850"` | Set a quota project: `gcloud auth application-default set-quota-project PROJECT_ID` |
| `403 "The caller does not have permission"` | Ask project owner to grant `roles/serviceusage.serviceUsageConsumer` |
| `403 "API has not been enabled"` | Enable the API: `gcloud services enable slides.googleapis.com --project=PROJECT_ID` |
| `404 "Requested entity was not found"` | Check file/presentation ID is correct and you have access |

## Requirements

- Python 3.11+
- pixi (package manager)
- gcloud CLI (Google Cloud SDK)
- Google account with Workspace access
- GCP project with APIs enabled

## Resources

- [Google Slides API Reference](https://developers.google.com/slides/api/reference/rest)
- [Google Sheets API Reference](https://developers.google.com/sheets/api/reference/rest)
- [Google Docs API Reference](https://developers.google.com/docs/api/reference/rest)
- [Google Drive API Reference](https://developers.google.com/drive/api/reference/rest/v3)
- [Google Forms API Reference](https://developers.google.com/forms/api/reference/rest)
- [pixi Documentation](https://pixi.sh/)

## License

MIT License - See LICENSE file for details.

---

**Automate Google Workspace with AI** | Powered by Cortex Code
