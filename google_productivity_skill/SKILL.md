# Google Productivity Skill for Cortex Code

Automate Google Workspace (Slides, Sheets, Docs, Drive, Forms) programmatically
via the Google Workspace REST APIs and Python. This single file contains everything
Cortex Code needs to generate working scripts: prerequisites, API reference,
bundled themes, the full helper library source, and real-world examples.

**Supported APIs:**
- **Google Slides** — Create/edit presentations, add slides, shapes, text, charts
- **Google Sheets** — Create/read/write spreadsheets, format cells, create charts
- **Google Docs** — Create/edit documents, format text, insert content
- **Google Drive** — List/search/upload/download/share files and folders
- **Google Forms** — Create forms/quizzes, add questions, retrieve responses, link to Sheets

**Workflow:** Cortex Code reads this skill, generates a self-contained Python
script (embedding the helper functions and theme inline), writes it to the
project folder, and runs it with `pixi run python <script.py>`.

---

## Prerequisites

### 1. Install pixi (Python + dependency management)

pixi is a fast, cross-platform package manager for Python projects.

**macOS/Linux:**
```bash
curl -fsSL https://pixi.sh/install.sh | bash
```

**macOS (Homebrew):**
```bash
brew install pixi
```

**Windows (PowerShell):**
```powershell
iwr -useb https://pixi.sh/install.ps1 | iex
```

After installation, restart your terminal or run `source ~/.bashrc` (or equivalent).

Verify installation:
```bash
pixi --version
```

### 2. Install Google Cloud CLI (gcloud)

gcloud is required for authentication with Google APIs.

**macOS (Homebrew):**
```bash
brew install --cask google-cloud-sdk
```

**macOS/Linux (official installer):**
```bash
curl https://sdk.cloud.google.com | bash
exec -l $SHELL  # Restart shell
gcloud init
```

**Windows:**
Download and run the installer from: https://cloud.google.com/sdk/docs/install

Verify installation:
```bash
gcloud --version
```

### 3. Google Cloud Application Default Credentials (ADC)

```bash
gcloud auth application-default login
```

Creates `~/.config/gcloud/application_default_credentials.json`.

**Note:** Authentication does NOT require a GCP project. You're authenticating as
a user via OAuth, not as a service account.

### 4. Quota Project (REQUIRED)

**⚠️ IMPORTANT:** A Google Cloud project with enabled APIs is ALWAYS required to use
Google Workspace APIs. Without it, all API calls will fail with 403 errors.

Google needs a project to:
- Track API usage and quotas
- Enable specific APIs (Slides, Sheets, Drive, Docs, Forms)
- Bill for usage (free tier is generous, typically $0 for normal use)

#### For Organization Users

Use `your-gcp-project` — it's pre-configured with all required APIs enabled and
accessible to all @yourcompany.com users:

```bash
gcloud auth application-default set-quota-project your-gcp-project
```

Or in Python:
```python
creds, _ = google.auth.default(scopes=[...])
creds = creds.with_quota_project("your-gcp-project")
```

#### For Other Users (No Existing Project)

**Option A: Cortex Code creates a project for you**

If no accessible project exists, Cortex Code will attempt to create one:

```bash
# Create a new GCP project
gcloud projects create USER_PROJECT_NAME --name="Google Workspace APIs"

# Enable all required APIs
gcloud services enable slides.googleapis.com --project=USER_PROJECT_NAME
gcloud services enable sheets.googleapis.com --project=USER_PROJECT_NAME
gcloud services enable docs.googleapis.com --project=USER_PROJECT_NAME
gcloud services enable drive.googleapis.com --project=USER_PROJECT_NAME
gcloud services enable forms.googleapis.com --project=USER_PROJECT_NAME

# Set as quota project for ADC
gcloud auth application-default set-quota-project USER_PROJECT_NAME
```

To make the project accessible to other users in your organization:
```bash
# Grant Service Usage Consumer role to a group or domain
gcloud projects add-iam-policy-binding USER_PROJECT_NAME \
  --member="domain:yourcompany.com" \
  --role="roles/serviceusage.serviceUsageConsumer"

# Or grant to specific users
gcloud projects add-iam-policy-binding USER_PROJECT_NAME \
  --member="user:colleague@yourcompany.com" \
  --role="roles/serviceusage.serviceUsageConsumer"
```

**Option B: Ask the user to create one**

If Cortex Code cannot create a project (permissions, billing account required, etc.),
ask the user to:

1. Go to https://console.cloud.google.com/projectcreate
2. Create a new project (any name)
3. Enable APIs at: https://console.cloud.google.com/apis/library
   - Google Slides API
   - Google Sheets API
   - Google Drive API
   - Google Docs API
   - Google Forms API
4. Run: `gcloud auth application-default set-quota-project YOUR_PROJECT_ID`

#### Verifying Project Access

```bash
# Check current quota project
gcloud auth application-default print-access-token 2>&1 | head -5

# Test API access (should return empty list, not error)
curl -s "https://www.googleapis.com/drive/v3/files?pageSize=1" \
  -H "Authorization: Bearer $(gcloud auth application-default print-access-token)" \
  | head -20

# List enabled APIs in a project
gcloud services list --enabled --project=PROJECT_ID
```

#### Common Errors

**`HttpError 403 "Google Slides API has not been used in project 764086051850"`**
— ADC is using Google's default OAuth project (764086051850) which doesn't have
APIs enabled. Fix: Set a quota project with `gcloud auth application-default set-quota-project YOUR_PROJECT`.

**`HttpError 403 "The caller does not have permission"`**
— Your user doesn't have `serviceusage.services.use` permission on the project.
Fix: Ask project owner to grant you `roles/serviceusage.serviceUsageConsumer`.

**`HttpError 403 "API has not been enabled"`**
— The specific API isn't enabled in the quota project.
Fix: `gcloud services enable APINAME.googleapis.com --project=PROJECT_ID`

---

## Project-Based Workflow

Each presentation should have its own project folder with dependencies, scripts,
and change history. This enables reproducibility, collaboration, and rollback.

### Creating a New Presentation Project

```bash
# Create project folder
mkdir my_presentation && cd my_presentation

# Initialize pixi project
pixi init --format pyproject

# Add dependencies
pixi add python=3.11
pixi add --pypi google-api-python-client google-auth google-auth-oauthlib

# Install dependencies
pixi install
```

### Project Folder Structure

```
my_presentation/
├── pixi.toml              # Dependencies and project config
├── pixi.lock              # Locked dependency versions
├── HISTORY.md             # Change log with timestamps and descriptions
├── scripts/
│   ├── 001_create_presentation.py
│   ├── 002_add_title_slide.py
│   ├── 003_add_content_slides.py
│   └── ...
└── assets/                # Images, logos, etc.
    └── logo.png
```

### History Tracking (HISTORY.md)

Maintain a `HISTORY.md` file to track all changes for context and rollback:

```markdown
# Presentation History

**Presentation ID:** 17F6drLMbF3LD6Tp23CdTKOoog940ZBV4pXmcIBxg9UU
**URL:** https://docs.google.com/presentation/d/17F6drLMbF3LD6Tp23CdTKOoog940ZBV4pXmcIBxg9UU/edit
**Created:** 2026-03-21

---

## Change Log

### [003] 2026-03-21 14:30 — Added comparison slides
**Script:** `scripts/003_add_comparison.py`
**Changes:**
- Added slide 4: "Feature Comparison Table"
- Added slide 5: "Performance Benchmarks"
**Slide IDs:** id_a1b2c3d4e5f6, id_g7h8i9j0k1l2
**Rollback:** Delete slides 4-5 or run rollback script

### [002] 2026-03-21 13:15 — Added content slides
**Script:** `scripts/002_add_content.py`
**Changes:**
- Added slide 2: "Problem Statement"
- Added slide 3: "Our Solution"
**Slide IDs:** id_m3n4o5p6q7r8, id_s9t0u1v2w3x4

### [001] 2026-03-21 12:00 — Created presentation
**Script:** `scripts/001_create_presentation.py`
**Changes:**
- Created new presentation with title slide
- Applied corporate theme
**Presentation ID:** 17F6drLMbF3LD6Tp23CdTKOoog940ZBV4pXmcIBxg9UU
```

### Running Scripts

Always run scripts from the project folder using pixi:

```bash
cd my_presentation
pixi run python scripts/001_create_presentation.py
```

### Rollback Strategy

To rollback changes:
1. Check `HISTORY.md` for the slide IDs created in the change
2. Use the `deleteObject` API request to remove specific elements
3. Or restore from Google Slides version history (File → Version history)

Example rollback script:
```python
"""Rollback: Remove slides added in change [003]."""
from googleapiclient.discovery import build
import google.auth

PRES_ID = "17F6drLMbF3LD6Tp23CdTKOoog940ZBV4pXmcIBxg9UU"
SLIDE_IDS_TO_DELETE = ["id_a1b2c3d4e5f6", "id_g7h8i9j0k1l2"]

def main():
    creds, _ = google.auth.default(scopes=["https://www.googleapis.com/auth/presentations"])
    creds = creds.with_quota_project("your-gcp-project")
    svc = build("slides", "v1", credentials=creds)
    
    requests = [{"deleteObject": {"objectId": sid}} for sid in SLIDE_IDS_TO_DELETE]
    svc.presentations().batchUpdate(presentationId=PRES_ID, body={"requests": requests}).execute()
    print(f"Deleted {len(SLIDE_IDS_TO_DELETE)} slides")

if __name__ == "__main__":
    main()
```

---

## Best Practices for Amazing Presentations

### 1. Story First, Slides Second

The foundation of any great presentation is its **narrative arc**:
- **Problem** → What challenge exists?
- **Insight** → What did we discover?
- **Solution** → How do we solve it?
- **Impact** → What are the results?

**Don't** overwhelm stakeholders with dozens of slides. Focus on a structure that
guides the audience smoothly from challenge to resolution.

### 2. Design is Not Optional

Good design signals reliability and builds trust:
- Use consistent colors, fonts, and layouts throughout
- Maintain brand consistency (use the theme system)
- Leave white space — crowded slides lose the audience
- Use high-quality images and graphics

### 3. Data Storytelling

Transform raw numbers into insights:
- Use clear, readable charts (avoid 3D effects)
- Highlight the key metric — don't make them search
- Add subtle animations to direct attention
- Show trends, not just numbers

### 4. The Rule of Three

Audiences remember three things well:
- Three key points per slide maximum
- Three main sections in your presentation
- Three supporting stats for major claims

### 5. Visual Hierarchy

Guide the viewer's eye:
- Largest text = most important
- Use color to highlight, not decorate
- Consistent alignment creates professionalism
- Group related items together

### 6. Know Your Audience

Tailor content to who's watching:
- **Executives:** Focus on outcomes, costs, timelines
- **Engineers:** Technical depth, architecture, tradeoffs
- **Investors:** Market size, growth, competitive advantage
- **Customers:** Benefits, ease of use, support

### 7. One Idea Per Slide

Each slide should answer ONE question:
- If a slide needs extensive explanation, split it
- Use the slide title to state the key takeaway
- Content should support the title, not contradict it

### 8. Use Presenter Notes

Add detailed notes (not visible to audience):
- Key talking points
- Data sources and citations
- Backup information for Q&A
- Timing cues

---

## Software Architecture Slide Best Practices

Creating effective architecture diagrams requires balancing detail with clarity.

### Architecture Diagram Types

**1. C4 Model (Recommended)**
Hierarchical approach with four levels:
- **Context:** System + external actors (users, other systems)
- **Container:** Applications, databases, services
- **Component:** Internal modules within a container
- **Code:** Classes and functions (rarely needed in slides)

**2. UML Diagrams**
Formal notation for technical audiences:
- Class diagrams for data models
- Sequence diagrams for interactions
- Deployment diagrams for infrastructure

**3. Cloud Architecture Diagrams**
Use official icons from AWS, Azure, GCP:
- Show VPCs, subnets, and network topology
- Include data flow direction
- Number steps for complex flows

### Architecture Slide Design Principles

**Think in Layers:**
1. **User layer:** What users interact with (frontends, apps)
2. **Processing layer:** Services, APIs, business logic
3. **Data layer:** Databases, caches, storage

**Reduce Components Aggressively:**
- If you can't explain why it's on the slide in one sentence, remove it
- Group similar services under one label
- Abstract infrastructure into simple terms

**Labels Should Teach, Not Document:**
- ❌ "Auth Service" (internal name)
- ✅ "User Authentication & Access Control" (explains purpose)
- ❌ "Kafka Cluster"
- ✅ "Event Streaming Layer"

**Arrows Are Not Decoration:**
- Every arrow should answer: "What is flowing here?"
- Be consistent with direction (left-to-right for requests)
- Fewer intentional arrows > many confusing arrows

**Color Should Signal, Not Decorate:**
- Assign meaning: external=blue, internal=gray, data=green
- Keep palette small (3-4 colors max)
- If everything is colorful, nothing is meaningful

**White Space Matters:**
- Leave breathing room between layers
- Dense diagrams feel chaotic
- Space signals organization and intentionality

### Architecture Slide Patterns in Google Slides

**Pattern: Layered Architecture**
```
┌─────────────────────────────────────────────────────────────┐
│  [Users/Clients]  Web App    Mobile App    Partner API      │ ← User Layer
├─────────────────────────────────────────────────────────────┤
│  [Services]       Auth       API Gateway    Analytics       │ ← Processing Layer
├─────────────────────────────────────────────────────────────┤
│  [Data]           PostgreSQL    Redis    S3 Storage         │ ← Data Layer
└─────────────────────────────────────────────────────────────┘
```

**Implementation in Google Slides:**
1. Create horizontal bands (RECTANGLE) for each layer
2. Use different shades of theme colors for layers
3. Add TEXT_BOX elements for each component
4. Use small arrows (ARROW shape) to show data flow
5. Label arrows with what flows (data, requests, events)

**Pattern: Numbered Flow Diagram**
Show step-by-step process with numbered callouts:
1. Create numbered circles (ELLIPSE with fill)
2. Add arrows connecting the flow
3. Include a legend explaining each step

**Pattern: Before/After Comparison**
Split slide into two columns:
- Left: "Current State" with pain points
- Right: "Proposed Architecture" with improvements
- Use red/green colors sparingly to highlight problems/solutions

---

## Bundled Themes

Themes define colors, fonts, and layout defaults. Embed the relevant theme dict
directly in generated scripts — no external JSON files needed.

### corporate_2026

Corporate brand theme — blue/white with Montserrat + Open Sans.

```python
THEME_CORPORATE_2026 = {
    "name": "corporate_2026",
    "color_scheme": {
        "DARK1": "#11567F", "LIGHT1": "#FFFFFF",
        "DARK2": "#333333", "LIGHT2": "#F0F4F7",
        "ACCENT1": "#29B5E8", "ACCENT2": "#1DA1A8",
        "ACCENT3": "#4CAF50", "ACCENT4": "#FF6F00",
        "ACCENT5": "#7E57C2", "ACCENT6": "#78909C",
        "HYPERLINK": "#29B5E8", "FOLLOWED_HYPERLINK": "#11567F",
    },
    "fonts": {"title": "Montserrat", "body": "Open Sans", "code": "Roboto Mono"},
    "colors": {
        "primary": "#29B5E8", "primary_dark": "#11567F",
        "white": "#FFFFFF", "light_gray": "#F0F4F7",
        "dark_gray": "#333333", "medium_gray": "#555555",
        "accent_teal": "#1DA1A8", "accent_green": "#4CAF50",
    },
    "slide_defaults": {
        "title_font_size_pt": 28, "body_font_size_pt": 12,
        "bullet_font_size_pt": 11, "caption_font_size_pt": 9,
        "header_bar_height_inches": 0.95, "margin_inches": 0.5,
    },
}
```

### minimal_dark

Dark background minimal theme — clean and modern.

```python
THEME_MINIMAL_DARK = {
    "name": "minimal_dark",
    "color_scheme": {
        "DARK1": "#1A1A2E", "LIGHT1": "#E0E0E0",
        "DARK2": "#16213E", "LIGHT2": "#2A2A4A",
        "ACCENT1": "#0F3460", "ACCENT2": "#E94560",
        "ACCENT3": "#53D8FB", "ACCENT4": "#F5A623",
        "ACCENT5": "#7B68EE", "ACCENT6": "#6C757D",
        "HYPERLINK": "#53D8FB", "FOLLOWED_HYPERLINK": "#7B68EE",
    },
    "fonts": {"title": "Poppins", "body": "Inter", "code": "Fira Code"},
    "colors": {
        "primary": "#0F3460", "primary_dark": "#1A1A2E",
        "white": "#E0E0E0", "light_gray": "#2A2A4A",
        "dark_gray": "#1A1A2E", "medium_gray": "#A0A0B0",
        "accent_red": "#E94560", "accent_cyan": "#53D8FB",
    },
    "slide_defaults": {
        "title_font_size_pt": 30, "body_font_size_pt": 13,
        "bullet_font_size_pt": 12, "caption_font_size_pt": 9,
        "header_bar_height_inches": 0.85, "margin_inches": 0.6,
    },
}
```

---

## Core Concepts

### Units: EMU (English Metric Units)

Google Slides uses EMU for all positioning and sizing.

| Conversion | Value |
|---|---|
| 1 inch | 914,400 EMU |
| 1 point | 12,700 EMU |
| 1 cm | 360,000 EMU |
| Slide width (16:9) | 9,144,000 EMU (10 in) |
| Slide height (16:9) | 5,143,500 EMU (5.625 in) |

Helper: `emu = lambda inches: int(inches * 914400)`

### Object IDs

- Must match `[a-zA-Z0-9_][a-zA-Z0-9_\-:]{4,49}` (5–50 chars).
- Generate with: `"id_" + uuid.uuid4().hex[:20]`
- Or omit and let the API generate one (returned in response).

### The batchUpdate Pattern

**All mutations** go through `presentations.batchUpdate`. Build a list of
request dicts, send them in one call. This is atomic and saves quota.

```python
body = {"requests": [req1, req2, ...]}
response = service.presentations().batchUpdate(
    presentationId=PRES_ID, body=body
).execute()
```

Each `batchUpdate` counts as 1 write request regardless of sub-request count.

### Predefined Layouts

| Layout | Description |
|---|---|
| `BLANK` | No placeholders — full control |
| `TITLE` | Title + subtitle |
| `TITLE_AND_BODY` | Title + body text |
| `TITLE_AND_TWO_COLUMNS` | Title + two body columns |
| `TITLE_ONLY` | Just a title placeholder |
| `SECTION_HEADER` | Section divider |
| `ONE_COLUMN_TEXT` | Single column with title |
| `MAIN_POINT` | Large centered text |
| `BIG_NUMBER` | Big number heading |
| `CAPTION_ONLY` | Caption at bottom |

**Recommendation:** Use `BLANK` and build your own layout with shapes/text boxes.
Predefined layouts depend on the master theme and may not match expectations.

---

## API Request Reference

### Create a Slide

```python
{"createSlide": {
    "objectId": slide_id,
    "insertionIndex": 3,
    "slideLayoutReference": {"predefinedLayout": "BLANK"},
}}
```

### Create a Shape / Text Box

```python
{"createShape": {
    "objectId": obj_id,
    "shapeType": "TEXT_BOX",  # or RECTANGLE, ELLIPSE, etc.
    "elementProperties": {
        "pageObjectId": slide_id,
        "size": {
            "width": {"magnitude": width_emu, "unit": "EMU"},
            "height": {"magnitude": height_emu, "unit": "EMU"},
        },
        "transform": {
            "scaleX": 1, "scaleY": 1,
            "translateX": x_emu, "translateY": y_emu,
            "unit": "EMU",
        },
    },
}}
```

### Create Lines and Arrows

```python
{"createLine": {
    "objectId": line_id,
    "lineCategory": "STRAIGHT",  # or BENT, CURVED
    "elementProperties": {
        "pageObjectId": slide_id,
        "size": {
            "width": {"magnitude": width_emu, "unit": "EMU"},
            "height": {"magnitude": height_emu, "unit": "EMU"},
        },
        "transform": {
            "scaleX": 1, "scaleY": 1,
            "translateX": x_emu, "translateY": y_emu,
            "unit": "EMU",
        },
    },
}}

# Add arrow head
{"updateLineProperties": {
    "objectId": line_id,
    "lineProperties": {
        "startArrow": "NONE",
        "endArrow": "OPEN_ARROW",  # or STEALTH_ARROW, FILL_ARROW
    },
    "fields": "startArrow,endArrow",
}}
```

### Insert Text

```python
{"insertText": {"objectId": obj_id, "text": "Hello", "insertionIndex": 0}}
```

### Style Text

```python
{"updateTextStyle": {
    "objectId": obj_id,
    "textRange": {"type": "ALL"},  # or FIXED_RANGE with startIndex/endIndex
    "style": {
        "fontFamily": "Montserrat",
        "fontSize": {"magnitude": 24, "unit": "PT"},
        "bold": True,
        "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.16, "green": 0.34, "blue": 0.5}}},
    },
    "fields": "fontFamily,fontSize,bold,foregroundColor",
}}
```

**Critical:** The `fields` parameter is a field mask. Only listed fields are
changed. Always include it.

### Set Shape Background

```python
{"updateShapeProperties": {
    "objectId": obj_id,
    "fields": "shapeBackgroundFill.solidFill.color",
    "shapeProperties": {
        "shapeBackgroundFill": {
            "solidFill": {"color": {"rgbColor": {"red": 0.16, "green": 0.71, "blue": 0.91}}}
        }
    },
}}
```

### Remove Shape Outline

```python
{"updateShapeProperties": {
    "objectId": obj_id,
    "fields": "outline.propertyState",
    "shapeProperties": {"outline": {"propertyState": "NOT_RENDERED"}},
}}
```

### Set Slide Background

```python
{"updatePageProperties": {
    "objectId": slide_id,
    "pageProperties": {
        "pageBackgroundFill": {
            "solidFill": {"color": {"rgbColor": {"red": 0.07, "green": 0.34, "blue": 0.5}}}
        }
    },
    "fields": "pageBackgroundFill.solidFill.color",
}}
```

### Set Paragraph Style

```python
{"updateParagraphStyle": {
    "objectId": obj_id,
    "textRange": {"type": "ALL"},
    "style": {
        "alignment": "CENTER",  # START, CENTER, END, JUSTIFIED
        "spaceAbove": {"magnitude": 6, "unit": "PT"},
        "spaceBelow": {"magnitude": 6, "unit": "PT"},
    },
    "fields": "alignment,spaceAbove,spaceBelow",
}}
```

### Apply Theme Color Scheme to Master

The 12 required `ThemeColorType` entries **must be in this order**:
`DARK1, LIGHT1, DARK2, LIGHT2, ACCENT1–ACCENT6, HYPERLINK, FOLLOWED_HYPERLINK`

```python
pres = service.presentations().get(presentationId=PRES_ID).execute()
master_id = pres["masters"][0]["objectId"]

color_scheme = {"colors": [
    {"type": t, "color": hex_to_rgb(theme["color_scheme"][t])}
    for t in ["DARK1","LIGHT1","DARK2","LIGHT2",
              "ACCENT1","ACCENT2","ACCENT3","ACCENT4","ACCENT5","ACCENT6",
              "HYPERLINK","FOLLOWED_HYPERLINK"]
]}

body = {"requests": [{"updatePageProperties": {
    "objectId": master_id,
    "pageProperties": {"colorScheme": color_scheme},
    "fields": "colorScheme.colors",
}}]}
service.presentations().batchUpdate(presentationId=PRES_ID, body=body).execute()
```

### Reading a Presentation

```python
pres = service.presentations().get(presentationId=PRES_ID).execute()
print("Title:", pres["title"])
print("Slides:", len(pres["slides"]))
for slide in pres["slides"]:
    for el in slide.get("pageElements", []):
        if "shape" in el and "text" in el["shape"]:
            for te in el["shape"]["text"].get("textElements", []):
                if "textRun" in te:
                    print(f"  {te['textRun']['content'][:60]}")
```

---

## Slide Layout Patterns

### Pattern: Header Bar + Content

1. Full-width `RECTANGLE` at top (header bar) — `0.85–0.95 in` tall, dark brand color
2. `TEXT_BOX` overlaid on the header — white text, bold title font
3. Thin `RECTANGLE` below header — `0.04 in` tall, accent color
4. Content area below — text boxes, shapes, cards

### Pattern: Stat Cards

1. `RECTANGLE` with brand color fill
2. Overlay `TEXT_BOX` with large bold number (28pt)
3. Overlay second `TEXT_BOX` with label (10pt)
4. Stack vertically with `0.1–0.15 in` gaps

### Pattern: Two-Column

1. Left column: `x=0.5in, width=4.5in` — bullet points
2. Right column: `x=5.4in, width=4.1in` — stats, images, cards
3. Bottom callout bar: full width at `y=4.6in`

### Pattern: Architecture Diagram

1. Create layer bands (RECTANGLE) with subtle colors
2. Add component boxes within each layer
3. Use createLine with arrow properties for connections
4. Add numbered callouts (ELLIPSE with text) for flow steps
5. Include a legend explaining colors/symbols

---

## Bundled Helper Library

Embed this code at the top of generated scripts. It provides all the helper
functions needed to build slides. **No external imports beyond the Google API
client are required.**

```python
# ── slides_helper (embed in generated scripts) ─────────────────────────
import uuid
import google.auth
from googleapiclient.discovery import build

EMU_PER_INCH = 914_400
EMU_PER_PT = 12_700
SLIDE_WIDTH = 10 * EMU_PER_INCH   # 9144000
SLIDE_HEIGHT = int(5.625 * EMU_PER_INCH)  # 5143500

def emu(inches: float) -> int:
    return int(inches * EMU_PER_INCH)

def pt(points: float) -> int:
    return int(points * EMU_PER_PT)

def new_id() -> str:
    return "id_" + uuid.uuid4().hex[:20]

def hex_to_rgb(hex_color: str) -> dict:
    h = hex_color.lstrip("#")
    return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255, "blue": int(h[4:6], 16) / 255}

def get_service(quota_project: str | None = None):
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/presentations",
                "https://www.googleapis.com/auth/drive"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    return build("slides", "v1", credentials=creds)

def batch_update(service, presentation_id: str, requests: list) -> dict:
    return service.presentations().batchUpdate(
        presentationId=presentation_id, body={"requests": requests}).execute()

def get_presentation(service, presentation_id: str) -> dict:
    return service.presentations().get(presentationId=presentation_id).execute()

def create_blank_slide(slide_id: str, insertion_index: int | None = None) -> dict:
    req = {"createSlide": {
        "objectId": slide_id,
        "slideLayoutReference": {"predefinedLayout": "BLANK"},
    }}
    if insertion_index is not None:
        req["createSlide"]["insertionIndex"] = insertion_index
    return req

def create_shape(page_id, obj_id, x, y, width, height, shape_type="RECTANGLE"):
    return {"createShape": {
        "objectId": obj_id, "shapeType": shape_type,
        "elementProperties": {
            "pageObjectId": page_id,
            "size": {"width": {"magnitude": width, "unit": "EMU"},
                     "height": {"magnitude": height, "unit": "EMU"}},
            "transform": {"scaleX": 1, "scaleY": 1,
                          "translateX": x, "translateY": y, "unit": "EMU"},
        },
    }}

def create_textbox(page_id, obj_id, x, y, width, height):
    return create_shape(page_id, obj_id, x, y, width, height, "TEXT_BOX")

def create_line(page_id, obj_id, x, y, width, height, category="STRAIGHT"):
    return {"createLine": {
        "objectId": obj_id, "lineCategory": category,
        "elementProperties": {
            "pageObjectId": page_id,
            "size": {"width": {"magnitude": width, "unit": "EMU"},
                     "height": {"magnitude": height, "unit": "EMU"}},
            "transform": {"scaleX": 1, "scaleY": 1,
                          "translateX": x, "translateY": y, "unit": "EMU"},
        },
    }}

def set_line_arrow(obj_id, end_arrow="OPEN_ARROW", start_arrow="NONE"):
    return {"updateLineProperties": {
        "objectId": obj_id,
        "lineProperties": {"startArrow": start_arrow, "endArrow": end_arrow},
        "fields": "startArrow,endArrow",
    }}

def insert_text(obj_id, text, index=0):
    return {"insertText": {"objectId": obj_id, "text": text, "insertionIndex": index}}

def style_text(obj_id, style, fields, start=None, end=None):
    text_range = {"type": "ALL"}
    if start is not None and end is not None:
        text_range = {"type": "FIXED_RANGE", "startIndex": start, "endIndex": end}
    return {"updateTextStyle": {
        "objectId": obj_id, "style": style, "fields": fields, "textRange": text_range}}

def set_shape_fill(obj_id, hex_color):
    return {"updateShapeProperties": {
        "objectId": obj_id, "fields": "shapeBackgroundFill.solidFill.color",
        "shapeProperties": {"shapeBackgroundFill": {
            "solidFill": {"color": {"rgbColor": hex_to_rgb(hex_color)}}}}}}

def remove_shape_outline(obj_id):
    return {"updateShapeProperties": {
        "objectId": obj_id, "fields": "outline.propertyState",
        "shapeProperties": {"outline": {"propertyState": "NOT_RENDERED"}}}}

def set_page_background(page_id, hex_color):
    return {"updatePageProperties": {
        "objectId": page_id,
        "pageProperties": {"pageBackgroundFill": {
            "solidFill": {"color": {"rgbColor": hex_to_rgb(hex_color)}}}},
        "fields": "pageBackgroundFill.solidFill.color"}}

def set_paragraph_style(obj_id, alignment="START", space_above=0, space_below=0):
    return {"updateParagraphStyle": {
        "objectId": obj_id, "textRange": {"type": "ALL"},
        "style": {"alignment": alignment,
                  "spaceAbove": {"magnitude": space_above, "unit": "PT"},
                  "spaceBelow": {"magnitude": space_below, "unit": "PT"}},
        "fields": "alignment,spaceAbove,spaceBelow"}}

def delete_text(obj_id):
    return {"deleteText": {"objectId": obj_id, "textRange": {"type": "ALL"}}}

def delete_object(obj_id):
    return {"deleteObject": {"objectId": obj_id}}
# ── end slides_helper ───────────────────────────────────────────────────
```

---

## Complete Working Example

This is the script that was used to add a "Why We Are More Cost Effective"
slide to an existing presentation. It demonstrates the header-bar + two-column +
stat-cards + callout-bar pattern with the corporate theme.

```python
"""Add a 'Why We Are More Cost Effective' slide."""
import uuid
import google.auth
from googleapiclient.discovery import build

# ── Inline helpers (from Bundled Helper Library above) ──────────────────
EMU_PER_INCH = 914_400
SLIDE_WIDTH = 10 * EMU_PER_INCH

def emu(inches): return int(inches * EMU_PER_INCH)
def new_id(): return "id_" + uuid.uuid4().hex[:20]
def hex_to_rgb(h):
    h = h.lstrip("#")
    return {"red": int(h[0:2],16)/255, "green": int(h[2:4],16)/255, "blue": int(h[4:6],16)/255}

def get_service():
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/presentations",
                "https://www.googleapis.com/auth/drive"])
    return build("slides", "v1", credentials=creds)

def create_shape(pid, oid, x, y, w, h, st="RECTANGLE"):
    return {"createShape": {"objectId": oid, "shapeType": st,
        "elementProperties": {"pageObjectId": pid,
            "size": {"width": {"magnitude": w, "unit": "EMU"},
                     "height": {"magnitude": h, "unit": "EMU"}},
            "transform": {"scaleX":1,"scaleY":1,"translateX":x,"translateY":y,"unit":"EMU"}}}}

def create_textbox(pid, oid, x, y, w, h):
    return create_shape(pid, oid, x, y, w, h, "TEXT_BOX")

def insert_text(oid, text, idx=0):
    return {"insertText": {"objectId": oid, "text": text, "insertionIndex": idx}}

def style_text(oid, style, fields, start=None, end=None):
    tr = {"type": "ALL"} if start is None else {"type":"FIXED_RANGE","startIndex":start,"endIndex":end}
    return {"updateTextStyle": {"objectId": oid, "style": style, "fields": fields, "textRange": tr}}

def set_fill(oid, color):
    return {"updateShapeProperties": {"objectId": oid,
        "fields": "shapeBackgroundFill.solidFill.color",
        "shapeProperties": {"shapeBackgroundFill": {"solidFill": {"color": {"rgbColor": hex_to_rgb(color)}}}}}}

def no_outline(oid):
    return {"updateShapeProperties": {"objectId": oid,
        "fields": "outline.propertyState",
        "shapeProperties": {"outline": {"propertyState": "NOT_RENDERED"}}}}

def para_style(oid, align="START"):
    return {"updateParagraphStyle": {"objectId": oid, "textRange": {"type":"ALL"},
        "style": {"alignment": align}, "fields": "alignment"}}
# ── End helpers ─────────────────────────────────────────────────────────

# ── Theme ───────────────────────────────────────────────────────────────
SF_BLUE, SF_DARK_BLUE, SF_WHITE = "#29B5E8", "#11567F", "#FFFFFF"
SF_LIGHT_GRAY, SF_DARK_GRAY, SF_MED_GRAY = "#F0F4F7", "#333333", "#555555"
SF_TEAL = "#1DA1A8"
FONT_TITLE, FONT_BODY = "Montserrat", "Open Sans"
# ── End theme ───────────────────────────────────────────────────────────

PRES_ID = "17F6drLMbF3LD6Tp23CdTKOoog940ZBV4pXmcIBxg9UU"

def main():
    svc = get_service()
    sid = new_id()
    reqs = []

    # Blank slide at index 3
    reqs.append({"createSlide": {"objectId": sid, "insertionIndex": 3,
        "slideLayoutReference": {"predefinedLayout": "BLANK"}}})

    # Header bar
    hid = new_id()
    reqs += [create_shape(sid, hid, 0, 0, SLIDE_WIDTH, emu(0.95)),
             set_fill(hid, SF_DARK_BLUE), no_outline(hid)]

    # Title
    tid = new_id()
    reqs += [create_textbox(sid, tid, emu(0.5), emu(0.12), emu(9.0), emu(0.7)),
             insert_text(tid, "Why We Are More Cost Effective"),
             style_text(tid, {"fontFamily": FONT_TITLE,
                 "fontSize": {"magnitude": 28, "unit": "PT"},
                 "foregroundColor": {"opaqueColor": {"rgbColor": hex_to_rgb(SF_WHITE)}},
                 "bold": True}, "fontFamily,fontSize,foregroundColor,bold")]

    # Accent line
    lid = new_id()
    reqs += [create_shape(sid, lid, 0, emu(0.95), SLIDE_WIDTH, emu(0.04)),
             set_fill(lid, SF_BLUE), no_outline(lid)]

    # Left column — bullets
    bid = new_id()
    reqs.append(create_textbox(sid, bid, emu(0.5), emu(1.2), emu(4.5), emu(3.8)))
    bullets = [
        ("Per-Second Billing", "Pay only for active compute. Auto-suspend when idle."),
        ("Separated Storage & Compute", "Scale each independently. No over-provisioning."),
        ("Near-Zero Admin Overhead", "Fully managed — no index tuning or vacuuming."),
        ("Up to 7x Compression", "Billed on compressed size, not raw volume."),
    ]
    text = ""
    for t, d in bullets:
        text += f"● {t}\n{d}\n\n"
    text = text.rstrip("\n")
    reqs += [insert_text(bid, text),
             style_text(bid, {"fontFamily": FONT_BODY,
                 "fontSize": {"magnitude": 11, "unit": "PT"},
                 "foregroundColor": {"opaqueColor": {"rgbColor": hex_to_rgb(SF_MED_GRAY)}}},
                 "fontFamily,fontSize,foregroundColor")]
    # Bold bullet headers
    pos = 0
    for t, d in bullets:
        s, e = pos + 2, pos + 2 + len(t)
        reqs.append(style_text(bid, {"bold": True,
            "foregroundColor": {"opaqueColor": {"rgbColor": hex_to_rgb(SF_DARK_BLUE)}},
            "fontSize": {"magnitude": 12, "unit": "PT"}},
            "bold,foregroundColor,fontSize", start=s, end=e))
        pos += 2 + len(t) + 1 + len(d) + 2

    # Right column — stat cards
    stats = [("58%", "faster than Databricks\nServerless", SF_BLUE),
             ("3–5×", "lower cost vs. BigQuery\non complex queries", SF_DARK_BLUE),
             ("28%", "cheaper than Databricks\nfor standard analytics", SF_TEAL)]
    cy = emu(1.2)
    for num, lbl, clr in stats:
        cid, nid, lid2 = new_id(), new_id(), new_id()
        reqs += [create_shape(sid, cid, emu(5.4), cy, emu(4.1), emu(1.05)),
                 set_fill(cid, clr), no_outline(cid),
                 create_textbox(sid, nid, emu(5.6), cy+emu(0.08), emu(1.4), emu(0.89)),
                 insert_text(nid, num),
                 style_text(nid, {"fontFamily": FONT_TITLE,
                     "fontSize": {"magnitude": 28, "unit": "PT"},
                     "foregroundColor": {"opaqueColor": {"rgbColor": hex_to_rgb(SF_WHITE)}},
                     "bold": True}, "fontFamily,fontSize,foregroundColor,bold"),
                 create_textbox(sid, lid2, emu(7.0), cy+emu(0.12), emu(2.3), emu(0.85)),
                 insert_text(lid2, lbl),
                 style_text(lid2, {"fontFamily": FONT_BODY,
                     "fontSize": {"magnitude": 10, "unit": "PT"},
                     "foregroundColor": {"opaqueColor": {"rgbColor": hex_to_rgb(SF_WHITE)}}},
                     "fontFamily,fontSize,foregroundColor")]
        cy += emu(1.2)

    # Bottom callout
    cbg, ctxt = new_id(), new_id()
    reqs += [create_shape(sid, cbg, emu(0.5), emu(4.65), emu(9.0), emu(0.55)),
             set_fill(cbg, SF_LIGHT_GRAY), no_outline(cbg),
             create_textbox(sid, ctxt, emu(0.7), emu(4.7), emu(8.6), emu(0.45)),
             insert_text(ctxt, "Zero ingestion fees  •  10% free cloud services  •  Multi-cloud portability"),
             style_text(ctxt, {"fontFamily": FONT_BODY,
                 "fontSize": {"magnitude": 9, "unit": "PT"},
                 "foregroundColor": {"opaqueColor": {"rgbColor": hex_to_rgb(SF_DARK_BLUE)}}},
                 "fontFamily,fontSize,foregroundColor"),
             para_style(ctxt, "CENTER")]

    resp = svc.presentations().batchUpdate(presentationId=PRES_ID, body={"requests": reqs}).execute()
    print(f"Done — {len(resp.get('replies', []))} operations")
    print(f"https://docs.google.com/presentation/d/{PRES_ID}/edit")

if __name__ == "__main__":
    main()
```

---

## Google Sheets API

Google Sheets uses the same authentication pattern as Slides. The API supports
creating spreadsheets, reading/writing values, formatting cells, and more.

### OAuth Scopes for Sheets

```python
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",  # Full access to Sheets
    "https://www.googleapis.com/auth/drive.file",     # Access files created by the app
]
```

### Creating a Spreadsheet

```python
"""Create a new Google Spreadsheet."""
import google.auth
from googleapiclient.discovery import build

def create_spreadsheet(title: str, sheet_names: list[str] = None, quota_project: str = None) -> str:
    """Create a new spreadsheet and return its URL.
    
    Args:
        title: Spreadsheet title
        sheet_names: Optional list of sheet names (default: ["Sheet1"])
        quota_project: GCP project for quota
    
    Returns:
        Edit URL for the spreadsheet
    """
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive.file"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    sheets_svc = build("sheets", "v4", credentials=creds)
    
    # Build spreadsheet structure
    spreadsheet_body = {
        "properties": {"title": title},
        "sheets": []
    }
    
    for i, name in enumerate(sheet_names or ["Sheet1"]):
        spreadsheet_body["sheets"].append({
            "properties": {"title": name, "index": i}
        })
    
    spreadsheet = sheets_svc.spreadsheets().create(body=spreadsheet_body).execute()
    spreadsheet_id = spreadsheet["spreadsheetId"]
    
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"

# Usage
url = create_spreadsheet("Q1 Sales Report", ["Summary", "Raw Data", "Charts"], "your-gcp-project")
print(url)
```

### Reading Cell Values

```python
"""Read values from a spreadsheet."""
def read_sheet_values(spreadsheet_id: str, range_name: str, quota_project: str = None) -> list:
    """Read values from a spreadsheet range.
    
    Args:
        spreadsheet_id: The spreadsheet ID from URL
        range_name: A1 notation range (e.g., "Sheet1!A1:D10" or "A1:D10")
        quota_project: GCP project for quota
    
    Returns:
        2D list of cell values
    """
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    sheets_svc = build("sheets", "v4", credentials=creds)
    
    result = sheets_svc.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    
    return result.get("values", [])

# Usage
values = read_sheet_values("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms", "Sheet1!A1:E5", "your-gcp-project")
for row in values:
    print(row)
```

### Writing Cell Values

```python
"""Write values to a spreadsheet."""
def write_sheet_values(spreadsheet_id: str, range_name: str, values: list, quota_project: str = None):
    """Write values to a spreadsheet range.
    
    Args:
        spreadsheet_id: The spreadsheet ID
        range_name: A1 notation range (e.g., "Sheet1!A1")
        values: 2D list of values to write
        quota_project: GCP project for quota
    """
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/spreadsheets"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    sheets_svc = build("sheets", "v4", credentials=creds)
    
    body = {"values": values}
    
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",  # or "RAW" for literal values
        body=body
    ).execute()

# Usage - Write headers and data
data = [
    ["Name", "Department", "Salary"],
    ["Alice", "Engineering", 120000],
    ["Bob", "Sales", 95000],
    ["Carol", "Marketing", 88000],
]
write_sheet_values("YOUR_SPREADSHEET_ID", "Sheet1!A1", data, "your-gcp-project")
```

### Appending Rows

```python
"""Append rows to the end of a table."""
def append_rows(spreadsheet_id: str, range_name: str, values: list, quota_project: str = None):
    """Append rows to a sheet (finds the last row automatically).
    
    Args:
        spreadsheet_id: The spreadsheet ID
        range_name: Range to search for table (e.g., "Sheet1!A:D")
        values: 2D list of rows to append
        quota_project: GCP project for quota
    """
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/spreadsheets"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    sheets_svc = build("sheets", "v4", credentials=creds)
    
    body = {"values": values}
    
    sheets_svc.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body=body
    ).execute()

# Usage - Add new employees
new_employees = [
    ["David", "Engineering", 110000],
    ["Eve", "HR", 75000],
]
append_rows("YOUR_SPREADSHEET_ID", "Sheet1!A:C", new_employees, "your-gcp-project")
```

### Sheets batchUpdate Pattern

Like Slides, Sheets has a `batchUpdate` for formatting and structural changes:

```python
"""Format cells using batchUpdate."""
def format_header_row(spreadsheet_id: str, sheet_id: int = 0, quota_project: str = None):
    """Make the first row bold with a background color.
    
    Args:
        spreadsheet_id: The spreadsheet ID
        sheet_id: The sheet's numeric ID (usually 0 for first sheet)
        quota_project: GCP project for quota
    """
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/spreadsheets"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    sheets_svc = build("sheets", "v4", credentials=creds)
    
    requests = [
        # Bold the header row
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"bold": True},
                        "backgroundColor": {"red": 0.16, "green": 0.71, "blue": 0.91}
                    }
                },
                "fields": "userEnteredFormat(textFormat,backgroundColor)"
            }
        },
        # Freeze the header row
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 1}
                },
                "fields": "gridProperties.frozenRowCount"
            }
        },
        # Auto-resize columns to fit content
        {
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 10
                }
            }
        }
    ]
    
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()
```

### Common Sheets batchUpdate Requests

| Request | Purpose |
|---------|---------|
| `repeatCell` | Apply formatting to a range |
| `updateCells` | Set values and formats directly |
| `appendCells` | Append cells to the end of a sheet |
| `insertDimension` | Insert rows or columns |
| `deleteDimension` | Delete rows or columns |
| `updateSheetProperties` | Rename sheet, freeze rows/cols |
| `addSheet` | Add a new sheet |
| `deleteSheet` | Delete a sheet |
| `mergeCells` | Merge a range of cells |
| `autoResizeDimensions` | Auto-fit column/row sizes |
| `addChart` | Add a chart to the sheet |
| `updateChartSpec` | Modify an existing chart |

### Complete Sheets Example

```python
"""Create a formatted sales report spreadsheet."""
import google.auth
from googleapiclient.discovery import build

QUOTA_PROJECT = "your-gcp-project"

def create_sales_report():
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive.file"])
    creds = creds.with_quota_project(QUOTA_PROJECT)
    
    sheets_svc = build("sheets", "v4", credentials=creds)
    drive_svc = build("drive", "v3", credentials=creds)
    
    # Create spreadsheet
    spreadsheet = sheets_svc.spreadsheets().create(body={
        "properties": {"title": "Q1 2026 Sales Report"},
        "sheets": [{"properties": {"title": "Sales Data"}}]
    }).execute()
    
    spreadsheet_id = spreadsheet["spreadsheetId"]
    sheet_id = spreadsheet["sheets"][0]["properties"]["sheetId"]
    
    # Write data
    data = [
        ["Region", "Product", "Q1 Sales", "Q1 Target", "% of Target"],
        ["North", "Widget A", 150000, 140000, "=C2/D2"],
        ["North", "Widget B", 95000, 100000, "=C3/D3"],
        ["South", "Widget A", 180000, 170000, "=C4/D4"],
        ["South", "Widget B", 120000, 130000, "=C5/D5"],
        ["East", "Widget A", 200000, 190000, "=C6/D6"],
        ["East", "Widget B", 85000, 90000, "=C7/D7"],
    ]
    
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Sales Data!A1",
        valueInputOption="USER_ENTERED",
        body={"values": data}
    ).execute()
    
    # Format with batchUpdate
    requests = [
        # Header formatting
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                        "backgroundColor": {"red": 0.07, "green": 0.34, "blue": 0.5}
                    }
                },
                "fields": "userEnteredFormat(textFormat,backgroundColor)"
            }
        },
        # Format currency columns
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "startColumnIndex": 2, "endColumnIndex": 4},
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0"}
                    }
                },
                "fields": "userEnteredFormat.numberFormat"
            }
        },
        # Format percentage column
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "startColumnIndex": 4, "endColumnIndex": 5},
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "PERCENT", "pattern": "0.0%"}
                    }
                },
                "fields": "userEnteredFormat.numberFormat"
            }
        },
        # Freeze header
        {
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount"
            }
        },
        # Auto-resize columns
        {
            "autoResizeDimensions": {
                "dimensions": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 5}
            }
        }
    ]
    
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()
    
    # Share with anyone who has the link
    drive_svc.permissions().create(
        fileId=spreadsheet_id,
        body={"type": "anyone", "role": "reader"}
    ).execute()
    
    print(f"✅ Sales report created!")
    print(f"🔗 https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit")

if __name__ == "__main__":
    create_sales_report()
```

### Sheets Gotchas

**Range notation:** Use A1 notation (`Sheet1!A1:D10`) or R1C1 notation.
Sheet name is optional if there's only one sheet.

**valueInputOption:** 
- `USER_ENTERED`: Parses input like a user typing (formulas work)
- `RAW`: Stores values exactly as provided (formulas stored as text)

**Sheet ID vs Sheet Name:** `batchUpdate` requests use numeric `sheetId`, 
while `values()` methods use sheet names in range strings.

**Rate limits:** 60 read/60 write requests per minute per user.

---

## Google Drive API

The Drive API manages files and folders across Google Workspace. Use it to
list, search, upload, download, and share files.

### OAuth Scopes for Drive

```python
# Common scope combinations
SCOPES_READ_ONLY = ["https://www.googleapis.com/auth/drive.metadata.readonly"]
SCOPES_READ_WRITE = ["https://www.googleapis.com/auth/drive"]
SCOPES_APP_FILES = ["https://www.googleapis.com/auth/drive.file"]  # Only files created by this app
```

### List Files and Folders

```python
"""List files in Google Drive."""
import google.auth
from googleapiclient.discovery import build

def list_files(max_results: int = 20, query: str = None, quota_project: str = None) -> list:
    """List files in Google Drive.
    
    Args:
        max_results: Maximum number of files to return
        query: Search query (see Drive API query syntax)
        quota_project: GCP project for quota
    
    Returns:
        List of file metadata dicts
    """
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/drive.metadata.readonly"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    drive_svc = build("drive", "v3", credentials=creds)
    
    results = drive_svc.files().list(
        pageSize=max_results,
        fields="files(id, name, mimeType, size, modifiedTime, parents)",
        q=query
    ).execute()
    
    return results.get("files", [])

# Usage
files = list_files(10, quota_project="your-gcp-project")
for f in files:
    print(f"{f['name']} ({f['mimeType']})")
```

### Search for Files

Drive API supports powerful query syntax:

```python
# Search examples
query_recent_docs = "mimeType='application/vnd.google-apps.document' and modifiedTime > '2026-01-01'"
query_spreadsheets = "mimeType='application/vnd.google-apps.spreadsheet'"
query_by_name = "name contains 'Q1 Report'"
query_in_folder = "'FOLDER_ID' in parents"
query_shared_with_me = "sharedWithMe=true"
query_starred = "starred=true"
query_images = "mimeType contains 'image/'"

# Common MIME types
MIME_TYPES = {
    "folder": "application/vnd.google-apps.folder",
    "document": "application/vnd.google-apps.document",
    "spreadsheet": "application/vnd.google-apps.spreadsheet",
    "presentation": "application/vnd.google-apps.presentation",
    "pdf": "application/pdf",
}

# Search for presentations containing "Sales"
files = list_files(query="mimeType='application/vnd.google-apps.presentation' and name contains 'Sales'")
```

### Upload Files

```python
"""Upload a file to Google Drive."""
from googleapiclient.http import MediaFileUpload

def upload_file(local_path: str, drive_filename: str = None, 
                folder_id: str = None, quota_project: str = None) -> dict:
    """Upload a local file to Google Drive.
    
    Args:
        local_path: Path to local file
        drive_filename: Name in Drive (default: same as local)
        folder_id: Optional parent folder ID
        quota_project: GCP project for quota
    
    Returns:
        File metadata dict with id and webViewLink
    """
    import os
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/drive.file"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    drive_svc = build("drive", "v3", credentials=creds)
    
    file_metadata = {"name": drive_filename or os.path.basename(local_path)}
    if folder_id:
        file_metadata["parents"] = [folder_id]
    
    media = MediaFileUpload(local_path, resumable=True)
    
    file = drive_svc.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name, webViewLink"
    ).execute()
    
    return file

# Usage
result = upload_file("/path/to/report.pdf", "Q1 Report.pdf", quota_project="your-gcp-project")
print(f"Uploaded: {result['webViewLink']}")
```

### Download Files

```python
"""Download a file from Google Drive."""
import io
from googleapiclient.http import MediaIoBaseDownload

def download_file(file_id: str, local_path: str, quota_project: str = None):
    """Download a file from Google Drive.
    
    Args:
        file_id: The Drive file ID
        local_path: Where to save the file locally
        quota_project: GCP project for quota
    """
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/drive.readonly"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    drive_svc = build("drive", "v3", credentials=creds)
    
    request = drive_svc.files().get_media(fileId=file_id)
    
    with open(local_path, "wb") as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            print(f"Download {int(status.progress() * 100)}%")

# Usage
download_file("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms", "/tmp/downloaded.xlsx")
```

### Export Google Workspace Files

Google Docs/Sheets/Slides must be exported to a specific format:

```python
"""Export Google Workspace files to different formats."""
def export_google_file(file_id: str, mime_type: str, local_path: str, quota_project: str = None):
    """Export a Google Workspace file to a specific format.
    
    Args:
        file_id: The Drive file ID
        mime_type: Export format MIME type
        local_path: Where to save the file
        quota_project: GCP project for quota
    """
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/drive.readonly"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    drive_svc = build("drive", "v3", credentials=creds)
    
    request = drive_svc.files().export_media(fileId=file_id, mimeType=mime_type)
    
    with open(local_path, "wb") as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()

# Export formats
EXPORT_FORMATS = {
    "document": {
        "pdf": "application/pdf",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "txt": "text/plain",
        "html": "text/html",
    },
    "spreadsheet": {
        "pdf": "application/pdf",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "csv": "text/csv",
    },
    "presentation": {
        "pdf": "application/pdf",
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    },
}

# Usage - Export a Google Doc as PDF
export_google_file("DOC_ID", "application/pdf", "/tmp/document.pdf", "your-gcp-project")
```

### Create Folders

```python
"""Create a folder in Google Drive."""
def create_folder(name: str, parent_id: str = None, quota_project: str = None) -> str:
    """Create a folder and return its ID.
    
    Args:
        name: Folder name
        parent_id: Optional parent folder ID
        quota_project: GCP project for quota
    
    Returns:
        The new folder's ID
    """
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/drive.file"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    drive_svc = build("drive", "v3", credentials=creds)
    
    file_metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder"
    }
    if parent_id:
        file_metadata["parents"] = [parent_id]
    
    folder = drive_svc.files().create(body=file_metadata, fields="id").execute()
    return folder["id"]

# Usage
folder_id = create_folder("Q1 2026 Reports", quota_project="your-gcp-project")
print(f"Created folder: {folder_id}")
```

### Share Files

```python
"""Share a file with users or make it public."""
def share_file(file_id: str, share_type: str, role: str = "reader", 
               email: str = None, quota_project: str = None):
    """Share a file.
    
    Args:
        file_id: The file ID to share
        share_type: "anyone", "user", "group", or "domain"
        role: "reader", "writer", "commenter"
        email: Email address (required for user/group types)
        quota_project: GCP project for quota
    """
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/drive"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    drive_svc = build("drive", "v3", credentials=creds)
    
    permission = {"type": share_type, "role": role}
    if email:
        permission["emailAddress"] = email
    
    drive_svc.permissions().create(
        fileId=file_id,
        body=permission,
        sendNotificationEmail=False
    ).execute()

# Usage examples
share_file("FILE_ID", "anyone", "reader")  # Anyone with link can view
share_file("FILE_ID", "user", "writer", "colleague@company.com")  # Share with specific user
share_file("FILE_ID", "domain", "reader", "example.com")  # Share with domain
```

### Drive API Gotchas

**File IDs:** Extract from URLs:
- Docs: `https://docs.google.com/document/d/{FILE_ID}/edit`
- Sheets: `https://docs.google.com/spreadsheets/d/{FILE_ID}/edit`
- Slides: `https://docs.google.com/presentation/d/{FILE_ID}/edit`
- Drive: `https://drive.google.com/file/d/{FILE_ID}/view`

**Pagination:** Large result sets use `nextPageToken`. Loop until it's empty:
```python
files = []
page_token = None
while True:
    response = drive_svc.files().list(pageToken=page_token, ...).execute()
    files.extend(response.get("files", []))
    page_token = response.get("nextPageToken")
    if not page_token:
        break
```

**Trash vs Delete:** `files().delete()` permanently deletes. Use `files().update(fileId, body={"trashed": True})` to move to trash.

**Shared Drive (Team Drive):** Add `supportsAllDrives=True` and `includeItemsFromAllDrives=True` to access shared drives.

---

## Google Docs API

### Prerequisites

Enable the Docs API (in addition to Drive API):
```bash
gcloud services enable docs.googleapis.com --project=YOUR_PROJECT_ID
gcloud services enable drive.googleapis.com --project=YOUR_PROJECT_ID
```

### Creating a Google Doc

Use the **Drive API** to create docs with initial content. The Docs API `create`
method has race conditions and doesn't reliably allow immediate `batchUpdate`.

```python
"""Create a Google Doc with content using Drive API."""
import google.auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaInMemoryUpload

def create_google_doc(title: str, content: str, share_with: list[str] = None, 
                      anyone_can_view: bool = False, quota_project: str = None) -> str:
    """Create a Google Doc and return the edit URL.
    
    Args:
        title: Document title
        content: Plain text content (will be converted to Doc format)
        share_with: List of email addresses to share with (editor access)
        anyone_can_view: If True, anyone with link can view
        quota_project: GCP project for quota (optional)
    
    Returns:
        Edit URL for the created document
    """
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/drive.file"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    drive_svc = build("drive", "v3", credentials=creds)
    
    # Create doc via Drive API (converts text to Google Doc format)
    file_metadata = {
        "name": title,
        "mimeType": "application/vnd.google-apps.document"
    }
    media = MediaInMemoryUpload(content.encode("utf-8"), mimetype="text/plain")
    
    file = drive_svc.files().create(
        body=file_metadata,
        media_body=media,
        fields="id,webViewLink"
    ).execute()
    
    doc_id = file["id"]
    
    # Share with specific users
    if share_with:
        for email in share_with:
            drive_svc.permissions().create(
                fileId=doc_id,
                body={"type": "user", "role": "writer", "emailAddress": email},
                sendNotificationEmail=False
            ).execute()
    
    # Make viewable by anyone with link
    if anyone_can_view:
        drive_svc.permissions().create(
            fileId=doc_id,
            body={"type": "anyone", "role": "reader"}
        ).execute()
    
    return file.get("webViewLink", f"https://docs.google.com/document/d/{doc_id}/edit")
```

### Complete Google Docs Example

```python
"""Create a formatted meeting notes document."""
import google.auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaInMemoryUpload

QUOTA_PROJECT = "your-project-id"  # Replace with your project

CONTENT = """Project Update – March 2026

Hi Team,

Here's our weekly status update.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🚀 COMPLETED THIS WEEK

• Feature A shipped to production
• Bug fixes for module B
• Documentation updated

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📋 ACTION ITEMS

☐ Review PR #123
☐ Schedule design review
☐ Update roadmap

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Best regards,
[Your Name]
"""

def main():
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/drive.file"])
    creds = creds.with_quota_project(QUOTA_PROJECT)
    
    drive_svc = build("drive", "v3", credentials=creds)
    
    file_metadata = {
        "name": "Project Update – March 2026",
        "mimeType": "application/vnd.google-apps.document"
    }
    media = MediaInMemoryUpload(CONTENT.encode("utf-8"), mimetype="text/plain")
    
    file = drive_svc.files().create(
        body=file_metadata,
        media_body=media,
        fields="id,webViewLink"
    ).execute()
    
    doc_id = file["id"]
    
    # Optional: Share with anyone who has the link
    drive_svc.permissions().create(
        fileId=doc_id,
        body={"type": "anyone", "role": "reader"}
    ).execute()
    
    print(f"✅ Document created!")
    print(f"🔗 {file.get('webViewLink', f'https://docs.google.com/document/d/{doc_id}/edit')}")

if __name__ == "__main__":
    main()
```

### Editing Existing Google Docs

For editing existing docs, use the Docs API `batchUpdate`:

```python
"""Edit an existing Google Doc."""
import google.auth
from googleapiclient.discovery import build

def get_docs_service(quota_project: str = None):
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/documents"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    return build("docs", "v1", credentials=creds)

def append_text(doc_id: str, text: str, quota_project: str = None):
    """Append text to the end of a document."""
    svc = get_docs_service(quota_project)
    
    # Get current document to find end index
    doc = svc.documents().get(documentId=doc_id).execute()
    end_index = doc["body"]["content"][-1]["endIndex"] - 1
    
    requests = [
        {"insertText": {"location": {"index": end_index}, "text": text}}
    ]
    
    svc.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()

def replace_text(doc_id: str, find: str, replace: str, quota_project: str = None):
    """Replace all occurrences of text in a document."""
    svc = get_docs_service(quota_project)
    
    requests = [
        {"replaceAllText": {
            "containsText": {"text": find, "matchCase": True},
            "replaceText": replace
        }}
    ]
    
    svc.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()
```

### Google Docs Gotchas

**Creating docs:** Use Drive API with `MediaInMemoryUpload`, NOT Docs API `create`.
The Docs API create has race conditions where `batchUpdate` fails with 404 immediately after creation.

**Sharing:** Documents created via API are private by default. Use `permissions().create()` to share:
- `{"type": "anyone", "role": "reader"}` — Anyone with link can view
- `{"type": "anyone", "role": "writer"}` — Anyone with link can edit  
- `{"type": "user", "role": "writer", "emailAddress": "..."}` — Share with specific user

**Text indices:** Document indices are 1-based, and the body starts at index 1.
The last element's `endIndex - 1` is where you can safely insert.

---

## Google Forms API

### Prerequisites

Enable the Forms API:
```bash
gcloud services enable forms.googleapis.com --project=YOUR_PROJECT_ID
```

**OAuth Scopes:**
- `https://www.googleapis.com/auth/forms.body` — Create and edit forms
- `https://www.googleapis.com/auth/forms.responses.readonly` — Read form responses
- `https://www.googleapis.com/auth/drive.file` — Access forms created by the app

### Form Structure

A Form contains:
- **Info** — Title and description
- **Settings** — Quiz mode, email collection
- **Items** — Questions, text, images, videos, page breaks
- **PublishSettings** — Published state, accepting responses

**Question Types:**
| Type | Description |
|------|-------------|
| `choiceQuestion` (RADIO) | Single choice (radio buttons) |
| `choiceQuestion` (CHECKBOX) | Multiple choice (checkboxes) |
| `choiceQuestion` (DROP_DOWN) | Dropdown selection |
| `textQuestion` | Short answer or paragraph |
| `scaleQuestion` | Linear scale (1-5, 1-10, etc.) |
| `dateQuestion` | Date picker |
| `timeQuestion` | Time picker |
| `fileUploadQuestion` | File upload (API read-only) |
| `ratingQuestion` | Star/heart/thumb rating |

### Creating a Form

```python
"""Create a Google Form with questions."""
import google.auth
from googleapiclient.discovery import build

QUOTA_PROJECT = "your-gcp-project"  # Replace with your project

def get_forms_service(quota_project: str = None):
    creds, _ = google.auth.default(
        scopes=[
            "https://www.googleapis.com/auth/forms.body",
            "https://www.googleapis.com/auth/drive.file"
        ])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    return build("forms", "v1", credentials=creds)

def create_form(title: str, description: str = None, quota_project: str = None) -> dict:
    """Create a new form and return form info."""
    svc = get_forms_service(quota_project)
    
    form = {
        "info": {
            "title": title,
            "documentTitle": title
        }
    }
    
    result = svc.forms().create(body=form).execute()
    form_id = result["formId"]
    
    # Add description if provided
    if description:
        svc.forms().batchUpdate(formId=form_id, body={
            "requests": [{
                "updateFormInfo": {
                    "info": {"description": description},
                    "updateMask": "description"
                }
            }]
        }).execute()
    
    return result

# Usage
form = create_form(
    "Customer Feedback Survey",
    "Please share your experience with our service.",
    QUOTA_PROJECT
)
print(f"Form ID: {form['formId']}")
print(f"Edit URL: https://docs.google.com/forms/d/{form['formId']}/edit")
print(f"Responder URL: {form.get('responderUri', 'N/A')}")
```

### Adding Questions

```python
"""Add various question types to a form."""

def add_questions(form_id: str, questions: list[dict], quota_project: str = None):
    """Add multiple questions to a form.
    
    Args:
        form_id: The form ID
        questions: List of question items
        quota_project: GCP project for quota
    """
    svc = get_forms_service(quota_project)
    
    requests = []
    for i, q in enumerate(questions):
        requests.append({
            "createItem": {
                "item": q,
                "location": {"index": i}
            }
        })
    
    svc.forms().batchUpdate(formId=form_id, body={"requests": requests}).execute()

# Multiple choice question (radio)
radio_question = {
    "title": "How satisfied are you with our service?",
    "questionItem": {
        "question": {
            "required": True,
            "choiceQuestion": {
                "type": "RADIO",
                "options": [
                    {"value": "Very satisfied"},
                    {"value": "Satisfied"},
                    {"value": "Neutral"},
                    {"value": "Dissatisfied"},
                    {"value": "Very dissatisfied"}
                ]
            }
        }
    }
}

# Checkbox question (multi-select)
checkbox_question = {
    "title": "Which features do you use? (Select all that apply)",
    "questionItem": {
        "question": {
            "choiceQuestion": {
                "type": "CHECKBOX",
                "options": [
                    {"value": "Dashboard"},
                    {"value": "Reports"},
                    {"value": "Analytics"},
                    {"value": "Integrations"},
                    {"value": "API"}
                ]
            }
        }
    }
}

# Dropdown question
dropdown_question = {
    "title": "How did you hear about us?",
    "questionItem": {
        "question": {
            "choiceQuestion": {
                "type": "DROP_DOWN",
                "options": [
                    {"value": "Search engine"},
                    {"value": "Social media"},
                    {"value": "Friend or colleague"},
                    {"value": "Advertisement"},
                    {"value": "Other"}
                ]
            }
        }
    }
}

# Short text question
text_question = {
    "title": "What is your email address?",
    "questionItem": {
        "question": {
            "required": True,
            "textQuestion": {"paragraph": False}  # False = short answer
        }
    }
}

# Paragraph text question
paragraph_question = {
    "title": "Any additional feedback?",
    "description": "Please share your thoughts in detail.",
    "questionItem": {
        "question": {
            "textQuestion": {"paragraph": True}  # True = paragraph
        }
    }
}

# Linear scale question
scale_question = {
    "title": "How likely are you to recommend us?",
    "description": "0 = Not likely, 10 = Very likely",
    "questionItem": {
        "question": {
            "scaleQuestion": {
                "low": 0,
                "high": 10,
                "lowLabel": "Not likely",
                "highLabel": "Very likely"
            }
        }
    }
}

# Date question
date_question = {
    "title": "When did you start using our service?",
    "questionItem": {
        "question": {
            "dateQuestion": {
                "includeYear": True,
                "includeTime": False
            }
        }
    }
}

# Usage
add_questions(form_id, [
    radio_question,
    checkbox_question,
    text_question,
    paragraph_question,
    scale_question
], QUOTA_PROJECT)
```

### Creating a Quiz with Grading

```python
"""Create a graded quiz."""

def create_quiz(form_id: str, quota_project: str = None):
    """Convert a form to a quiz."""
    svc = get_forms_service(quota_project)
    
    svc.forms().batchUpdate(formId=form_id, body={
        "requests": [{
            "updateSettings": {
                "settings": {"quizSettings": {"isQuiz": True}},
                "updateMask": "quizSettings.isQuiz"
            }
        }]
    }).execute()

def add_graded_question(form_id: str, quota_project: str = None):
    """Add a graded multiple choice question."""
    svc = get_forms_service(quota_project)
    
    graded_question = {
        "title": "What is the capital of France?",
        "questionItem": {
            "question": {
                "required": True,
                "grading": {
                    "pointValue": 10,
                    "correctAnswers": {
                        "answers": [{"value": "Paris"}]
                    },
                    "whenRight": {"text": "Correct! Paris is the capital of France."},
                    "whenWrong": {"text": "Incorrect. The correct answer is Paris."}
                },
                "choiceQuestion": {
                    "type": "RADIO",
                    "options": [
                        {"value": "London"},
                        {"value": "Paris"},
                        {"value": "Berlin"},
                        {"value": "Madrid"}
                    ]
                }
            }
        }
    }
    
    svc.forms().batchUpdate(formId=form_id, body={
        "requests": [{
            "createItem": {
                "item": graded_question,
                "location": {"index": 0}
            }
        }]
    }).execute()

# Usage
create_quiz(form_id, QUOTA_PROJECT)
add_graded_question(form_id, QUOTA_PROJECT)
```

### Retrieving Form Responses

```python
"""Retrieve responses from a form."""

def get_forms_responses_service(quota_project: str = None):
    creds, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/forms.responses.readonly"])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    return build("forms", "v1", credentials=creds)

def get_responses(form_id: str, quota_project: str = None) -> list[dict]:
    """Get all responses for a form."""
    svc = get_forms_responses_service(quota_project)
    
    responses = []
    page_token = None
    
    while True:
        result = svc.forms().responses().list(
            formId=form_id,
            pageToken=page_token
        ).execute()
        
        responses.extend(result.get("responses", []))
        page_token = result.get("nextPageToken")
        
        if not page_token:
            break
    
    return responses

def get_response_by_id(form_id: str, response_id: str, quota_project: str = None) -> dict:
    """Get a specific response by ID."""
    svc = get_forms_responses_service(quota_project)
    return svc.forms().responses().get(formId=form_id, responseId=response_id).execute()

# Usage
responses = get_responses(form_id, QUOTA_PROJECT)
for r in responses:
    print(f"Response ID: {r['responseId']}")
    print(f"Submitted: {r.get('createTime')}")
    for qid, answer in r.get("answers", {}).items():
        print(f"  Question {qid}: {answer.get('textAnswers', {}).get('answers', [])}")
```

### Forms + Sheets Integration (Linked Responses)

When you link a form to a Google Sheet, responses are automatically written to the sheet.
**Note:** The Forms API doesn't directly support linking to Sheets programmatically.
Use one of these approaches:

#### Approach 1: Manual Link via UI
1. Open the form in Google Forms
2. Go to Responses tab → Click Sheets icon → Create or select spreadsheet

#### Approach 2: Read Linked Sheet ID
If already linked, the `linkedSheetId` field shows which Sheet receives responses:

```python
"""Check if form is linked to a Sheet."""

def get_linked_sheet(form_id: str, quota_project: str = None) -> str | None:
    """Get the linked Sheet ID if one exists."""
    svc = get_forms_service(quota_project)
    form = svc.forms().get(formId=form_id).execute()
    return form.get("linkedSheetId")

# Usage
sheet_id = get_linked_sheet(form_id, QUOTA_PROJECT)
if sheet_id:
    print(f"Responses go to: https://docs.google.com/spreadsheets/d/{sheet_id}/edit")
else:
    print("Form is not linked to a Sheet")
```

#### Approach 3: Poll Responses and Write to Sheet

Create your own integration by polling responses and writing to Sheets:

```python
"""Poll form responses and write to a Sheet."""
import google.auth
from googleapiclient.discovery import build
from datetime import datetime

def sync_form_to_sheet(form_id: str, sheet_id: str, quota_project: str = None):
    """Sync form responses to a Google Sheet.
    
    Args:
        form_id: The Google Form ID
        sheet_id: The Google Sheet ID to write responses to
        quota_project: GCP project for quota
    """
    creds, _ = google.auth.default(scopes=[
        "https://www.googleapis.com/auth/forms.responses.readonly",
        "https://www.googleapis.com/auth/spreadsheets"
    ])
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    
    forms_svc = build("forms", "v1", credentials=creds)
    sheets_svc = build("sheets", "v4", credentials=creds)
    
    # Get form structure for question titles
    form = forms_svc.forms().get(formId=form_id).execute()
    question_map = {}
    for item in form.get("items", []):
        if "questionItem" in item:
            qid = item["questionItem"]["question"].get("questionId")
            if qid:
                question_map[qid] = item.get("title", "Untitled")
    
    # Get responses
    responses = forms_svc.forms().responses().list(formId=form_id).execute()
    
    if not responses.get("responses"):
        print("No responses to sync")
        return
    
    # Build header row
    headers = ["Timestamp", "Response ID"] + list(question_map.values())
    
    # Build data rows
    rows = [headers]
    for resp in responses["responses"]:
        row = [
            resp.get("createTime", ""),
            resp.get("responseId", "")
        ]
        for qid, title in question_map.items():
            answer = resp.get("answers", {}).get(qid, {})
            text_answers = answer.get("textAnswers", {}).get("answers", [])
            row.append(", ".join([a.get("value", "") for a in text_answers]))
        rows.append(row)
    
    # Write to Sheet
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range="A1",
        valueInputOption="USER_ENTERED",
        body={"values": rows}
    ).execute()
    
    print(f"Synced {len(rows) - 1} responses to Sheet")

# Usage
sync_form_to_sheet(form_id, sheet_id, QUOTA_PROJECT)
```

### Publishing a Form

Forms created via API after June 30, 2026 are unpublished by default:

```python
"""Publish a form to accept responses."""

def publish_form(form_id: str, accepting_responses: bool = True, quota_project: str = None):
    """Publish a form and optionally start accepting responses."""
    svc = get_forms_service(quota_project)
    
    svc.forms().setPublishSettings(formId=form_id, body={
        "publishSettings": {
            "publishState": {
                "isPublished": True,
                "isAcceptingResponses": accepting_responses
            }
        }
    }).execute()

def unpublish_form(form_id: str, quota_project: str = None):
    """Unpublish a form (stops accepting responses)."""
    svc = get_forms_service(quota_project)
    
    svc.forms().setPublishSettings(formId=form_id, body={
        "publishSettings": {
            "publishState": {
                "isPublished": False,
                "isAcceptingResponses": False
            }
        }
    }).execute()

# Usage
publish_form(form_id, True, QUOTA_PROJECT)  # Publish and accept responses
```

### Complete Survey Example with Sheets Backend

```python
"""Create a complete customer survey with automatic Sheets integration."""
import google.auth
from googleapiclient.discovery import build

QUOTA_PROJECT = "your-gcp-project"  # Replace with your project

def create_survey_with_sheet_backend():
    """Create a survey form and a Sheet to collect responses."""
    creds, _ = google.auth.default(scopes=[
        "https://www.googleapis.com/auth/forms.body",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/spreadsheets"
    ])
    creds = creds.with_quota_project(QUOTA_PROJECT)
    
    forms_svc = build("forms", "v1", credentials=creds)
    sheets_svc = build("sheets", "v4", credentials=creds)
    drive_svc = build("drive", "v3", credentials=creds)
    
    # 1. Create the form
    form = forms_svc.forms().create(body={
        "info": {
            "title": "Customer Satisfaction Survey",
            "documentTitle": "Customer Satisfaction Survey"
        }
    }).execute()
    form_id = form["formId"]
    
    # 2. Add description
    forms_svc.forms().batchUpdate(formId=form_id, body={
        "requests": [{
            "updateFormInfo": {
                "info": {"description": "We value your feedback! Please take a moment to share your experience."},
                "updateMask": "description"
            }
        }]
    }).execute()
    
    # 3. Add questions
    questions = [
        {
            "title": "How satisfied are you with our product?",
            "questionItem": {
                "question": {
                    "required": True,
                    "choiceQuestion": {
                        "type": "RADIO",
                        "options": [
                            {"value": "Very satisfied"},
                            {"value": "Satisfied"},
                            {"value": "Neutral"},
                            {"value": "Dissatisfied"},
                            {"value": "Very dissatisfied"}
                        ]
                    }
                }
            }
        },
        {
            "title": "What features do you use most?",
            "questionItem": {
                "question": {
                    "choiceQuestion": {
                        "type": "CHECKBOX",
                        "options": [
                            {"value": "Reporting"},
                            {"value": "Analytics"},
                            {"value": "Integrations"},
                            {"value": "API"},
                            {"value": "Mobile app"}
                        ]
                    }
                }
            }
        },
        {
            "title": "How likely are you to recommend us? (0-10)",
            "questionItem": {
                "question": {
                    "required": True,
                    "scaleQuestion": {
                        "low": 0,
                        "high": 10,
                        "lowLabel": "Not likely",
                        "highLabel": "Extremely likely"
                    }
                }
            }
        },
        {
            "title": "Your email (optional)",
            "questionItem": {
                "question": {
                    "textQuestion": {"paragraph": False}
                }
            }
        },
        {
            "title": "Additional comments",
            "questionItem": {
                "question": {
                    "textQuestion": {"paragraph": True}
                }
            }
        }
    ]
    
    requests = []
    for i, q in enumerate(questions):
        requests.append({
            "createItem": {
                "item": q,
                "location": {"index": i}
            }
        })
    
    forms_svc.forms().batchUpdate(formId=form_id, body={"requests": requests}).execute()
    
    # 4. Create a Sheet to store responses
    sheet = sheets_svc.spreadsheets().create(body={
        "properties": {"title": "Survey Responses - Customer Satisfaction"},
        "sheets": [{"properties": {"title": "Responses"}}]
    }).execute()
    sheet_id = sheet["spreadsheetId"]
    
    # 5. Add headers to Sheet
    headers = [["Timestamp", "Satisfaction", "Features Used", "NPS Score", "Email", "Comments"]]
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range="Responses!A1",
        valueInputOption="USER_ENTERED",
        body={"values": headers}
    ).execute()
    
    # 6. Format header row
    sheets_svc.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body={
        "requests": [{
            "repeatCell": {
                "range": {"sheetId": 0, "startRowIndex": 0, "endRowIndex": 1},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.2, "green": 0.4, "blue": 0.6},
                        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)"
            }
        }]
    }).execute()
    
    # 7. Make Sheet publicly viewable (optional)
    drive_svc.permissions().create(
        fileId=sheet_id,
        body={"type": "anyone", "role": "reader"}
    ).execute()
    
    # 8. Publish the form
    forms_svc.forms().setPublishSettings(formId=form_id, body={
        "publishSettings": {
            "publishState": {
                "isPublished": True,
                "isAcceptingResponses": True
            }
        }
    }).execute()
    
    # Get responder URL
    form_info = forms_svc.forms().get(formId=form_id).execute()
    
    print("=" * 60)
    print("✅ Survey Created Successfully!")
    print("=" * 60)
    print(f"📝 Form Edit URL: https://docs.google.com/forms/d/{form_id}/edit")
    print(f"📤 Form Share URL: {form_info.get('responderUri', f'https://docs.google.com/forms/d/{form_id}/viewform')}")
    print(f"📊 Responses Sheet: https://docs.google.com/spreadsheets/d/{sheet_id}/edit")
    print("=" * 60)
    print("\nNOTE: To automatically link form responses to Sheet, open the form")
    print("in Google Forms UI → Responses tab → Click Sheets icon → Select your sheet")
    
    return {"form_id": form_id, "sheet_id": sheet_id}

if __name__ == "__main__":
    create_survey_with_sheet_backend()
```

### Forms API Gotchas

**API vs UI linking:** The Forms API cannot programmatically link a form to a Sheet for automatic response collection. You must either:
1. Link manually in the Forms UI, or
2. Build your own sync using `forms.responses.list()` + Sheets API

**Publish settings:** Forms created via API after June 30, 2026 are unpublished by default. Call `setPublishSettings` to publish.

**File uploads:** The `fileUploadQuestion` type is read-only in the API. You cannot create file upload questions programmatically.

**Question order:** When adding multiple questions with `batchUpdate`, request order matters. Requests are processed sequentially, so use increasing indices.

**Form ID from URL:** Extract from `https://docs.google.com/forms/d/{FORM_ID}/edit`

**Response format:** Text answers are in `answers[questionId].textAnswers.answers[].value`. Checkbox responses are returned as multiple values in the array.

---

## Gotchas and Troubleshooting

### Quota Project Errors
**Symptom:** `HttpError 403` mentioning project `764086051850`.
**Fix:** `gcloud auth application-default set-quota-project YOUR_PROJECT_ID`

### `serviceusage.services.use` Permission
Grant **Service Usage Consumer** role:
```bash
gcloud projects add-iam-policy-binding YOUR_PROJECT \
  --member="user:you@example.com" \
  --role="roles/serviceusage.serviceUsageConsumer"
```

### Field Masks
Always use `fields` in update requests. Without them you risk overwriting
unrelated properties or getting rejected.

### Text Index Tracking
- Newlines count as 1 character
- `"● "` = 2 characters (bullet + space)
- Track `pos` carefully: `pos += prefix_len + title_len + 1 + desc_len + 2`
- Use `textRange.type: "ALL"` when styling uniformly

### Object ID Conflicts
- IDs must be unique across the entire presentation
- Always use `uuid.uuid4().hex` with a prefix

### Rate Limits
- 60 read/60 write requests per minute per user
- Each `batchUpdate` = 1 write request (batch everything)

### Font Availability
Google Slides uses Google Fonts. Unavailable fonts silently fall back to Arial.
Safe choices:
- Titles: Montserrat, Poppins, Roboto, Lato
- Body: Open Sans, Inter, Roboto, Source Sans Pro
- Code: Roboto Mono, Fira Code, Source Code Pro

### Presentation ID
Extract from the URL: `https://docs.google.com/presentation/d/{PRESENTATION_ID}/edit`

---

## gcloud CLI Tips

### Check Current Authentication

```bash
# See who you're authenticated as
gcloud auth list

# See current ADC configuration
gcloud auth application-default print-access-token
```

### Manage Multiple Accounts

```bash
# Login with a different account
gcloud auth login --no-launch-browser

# Switch active account
gcloud config set account you@example.com

# For ADC specifically
gcloud auth application-default login
```

### API Management

```bash
# List enabled APIs
gcloud services list --enabled --project=YOUR_PROJECT

# Enable an API
gcloud services enable sheets.googleapis.com --project=YOUR_PROJECT

# Check if API is enabled
gcloud services list --enabled --project=YOUR_PROJECT --filter="name:sheets"
```

### Troubleshooting Authentication

```bash
# Revoke and re-authenticate ADC
gcloud auth application-default revoke
gcloud auth application-default login

# Set quota project (fixes 764086051850 errors)
gcloud auth application-default set-quota-project YOUR_PROJECT

# Check ADC file location
echo $GOOGLE_APPLICATION_CREDENTIALS
# Default: ~/.config/gcloud/application_default_credentials.json
```

### Using curl with ADC

```bash
# Get access token for manual API calls
TOKEN=$(gcloud auth application-default print-access-token)

# Test Sheets API
curl "https://sheets.googleapis.com/v4/spreadsheets/SPREADSHEET_ID" \
  -H "Authorization: Bearer $TOKEN"

# Test Drive API
curl "https://www.googleapis.com/drive/v3/files?pageSize=5" \
  -H "Authorization: Bearer $TOKEN"
```

### Service Account Alternative

For server-to-server auth without user interaction:

```bash
# Create service account
gcloud iam service-accounts create my-automation \
  --display-name="Automation Account" \
  --project=YOUR_PROJECT

# Create key
gcloud iam service-accounts keys create key.json \
  --iam-account=my-automation@YOUR_PROJECT.iam.gserviceaccount.com

# Use in Python
export GOOGLE_APPLICATION_CREDENTIALS="/path/to/key.json"
```

**Note:** Service accounts can't access user Google Drive files by default.
You must explicitly share files with the service account email.
