#!/usr/bin/env python3
"""
=============================================================================
Demo 01: Create a Themed Presentation
=============================================================================
Creates a new Google Slides presentation with a professional theme,
including a title slide and several content slides.

Run: pixi run python demo/01_create_presentation.py
=============================================================================
"""

from googleapiclient.discovery import build
import google.auth
import uuid

# =============================================================================
# CONFIGURATION
# =============================================================================

QUOTA_PROJECT = "your-gcp-project"  # Replace with your GCP project ID
PRESENTATION_TITLE = "Company Overview Presentation"

# Professional Theme Colors (RGB 0-1 scale)
COLORS = {
    "primary": {"red": 0.11, "green": 0.53, "blue": 0.93},      # #1C87ED
    "accent": {"red": 0.0, "green": 0.84, "blue": 0.88},         # #00D6E0
    "dark": {"red": 0.09, "green": 0.11, "blue": 0.18},          # #171C2E
    "white": {"red": 1.0, "green": 1.0, "blue": 1.0},
    "light_gray": {"red": 0.96, "green": 0.97, "blue": 0.98},     # #F5F7FA
}

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def get_slides_service():
    """Get authenticated Slides API service."""
    scopes = ["https://www.googleapis.com/auth/presentations"]
    creds, _ = google.auth.default(scopes=scopes)
    creds = creds.with_quota_project(QUOTA_PROJECT)
    return build("slides", "v1", credentials=creds)

def gen_id():
    """Generate a unique object ID."""
    return f"id_{uuid.uuid4().hex[:12]}"

def create_text_box(slide_id, text, x, y, width, height, font_size=18, bold=False, color=None):
    """Create a text box with specified properties."""
    box_id = gen_id()
    color = color or COLORS["dark"]
    
    return [
        {
            "createShape": {
                "objectId": box_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {
                        "width": {"magnitude": width, "unit": "PT"},
                        "height": {"magnitude": height, "unit": "PT"}
                    },
                    "transform": {
                        "scaleX": 1, "scaleY": 1,
                        "translateX": x, "translateY": y,
                        "unit": "PT"
                    }
                }
            }
        },
        {
            "insertText": {
                "objectId": box_id,
                "text": text,
                "insertionIndex": 0
            }
        },
        {
            "updateTextStyle": {
                "objectId": box_id,
                "style": {
                    "fontSize": {"magnitude": font_size, "unit": "PT"},
                    "fontFamily": "Open Sans",
                    "bold": bold,
                    "foregroundColor": {"opaqueColor": {"rgbColor": color}}
                },
                "textRange": {"type": "ALL"},
                "fields": "fontSize,fontFamily,bold,foregroundColor"
            }
        }
    ]

def set_slide_background(slide_id, color):
    """Set solid background color for a slide."""
    return {
        "updatePageProperties": {
            "objectId": slide_id,
            "pageProperties": {
                "pageBackgroundFill": {
                    "solidFill": {"color": {"rgbColor": color}}
                }
            },
            "fields": "pageBackgroundFill"
        }
    }

# =============================================================================
# SLIDE CONTENT
# =============================================================================

SLIDES_CONTENT = [
    {
        "title": "Your Company Name",
        "subtitle": "Tagline Goes Here",
        "layout": "title",
        "background": COLORS["dark"]
    },
    {
        "title": "About Us",
        "bullets": [
            "Brief introduction to your company",
            "Key value proposition number one",
            "Key value proposition number two",
            "What makes you different",
            "Call to action or summary"
        ],
        "layout": "content",
        "background": COLORS["white"]
    },
    {
        "title": "Key Capabilities",
        "bullets": [
            "Capability 1 — Description of first capability",
            "Capability 2 — Description of second capability",
            "Capability 3 — Description of third capability",
            "Capability 4 — Description of fourth capability",
            "Capability 5 — Description of fifth capability"
        ],
        "layout": "content",
        "background": COLORS["white"]
    },
    {
        "title": "Why Choose Us?",
        "bullets": [
            "Benefit 1 — Explanation of first benefit",
            "Benefit 2 — Explanation of second benefit",
            "Benefit 3 — Explanation of third benefit",
            "Benefit 4 — Explanation of fourth benefit",
            "Benefit 5 — Explanation of fifth benefit"
        ],
        "layout": "content",
        "background": COLORS["white"]
    },
    {
        "title": "Get Started Today",
        "subtitle": "yourcompany.com/contact",
        "layout": "title",
        "background": COLORS["primary"]
    }
]

# =============================================================================
# MAIN
# =============================================================================

def main():
    print("🎨 Creating presentation...")
    
    # Get API service
    service = get_slides_service()
    
    # Create presentation
    presentation = service.presentations().create(
        body={"title": PRESENTATION_TITLE}
    ).execute()
    
    pres_id = presentation["presentationId"]
    print(f"✅ Created presentation: {pres_id}")
    
    # Get the default slide (we'll replace it)
    default_slide_id = presentation["slides"][0]["objectId"]
    
    # Build all requests
    requests = []
    slide_ids = []
    
    # Create slides (skip first since we have default)
    for i, slide in enumerate(SLIDES_CONTENT):
        if i == 0:
            slide_id = default_slide_id
        else:
            slide_id = gen_id()
            requests.append({
                "createSlide": {
                    "objectId": slide_id,
                    "insertionIndex": i
                }
            })
        slide_ids.append(slide_id)
    
    # Execute slide creation first
    if requests:
        service.presentations().batchUpdate(
            presentationId=pres_id,
            body={"requests": requests}
        ).execute()
        print(f"✅ Created {len(SLIDES_CONTENT)} slides")
    
    # Now add content to each slide
    content_requests = []
    
    for i, (slide_id, slide) in enumerate(zip(slide_ids, SLIDES_CONTENT)):
        # Set background
        content_requests.append(set_slide_background(slide_id, slide["background"]))
        
        # Determine text color based on background
        is_dark_bg = slide["background"] in [COLORS["dark"], COLORS["primary"]]
        text_color = COLORS["white"] if is_dark_bg else COLORS["dark"]
        accent_color = COLORS["accent"] if is_dark_bg else COLORS["primary"]
        
        if slide["layout"] == "title":
            # Title slide layout
            content_requests.extend(create_text_box(
                slide_id, slide["title"],
                x=50, y=180, width=620, height=80,
                font_size=48, bold=True, color=text_color
            ))
            if "subtitle" in slide:
                content_requests.extend(create_text_box(
                    slide_id, slide["subtitle"],
                    x=50, y=270, width=620, height=50,
                    font_size=24, bold=False, color=accent_color
                ))
        else:
            # Content slide layout
            content_requests.extend(create_text_box(
                slide_id, slide["title"],
                x=50, y=30, width=620, height=50,
                font_size=32, bold=True, color=COLORS["primary"]
            ))
            
            # Add bullet points
            if "bullets" in slide:
                bullet_text = "\n".join(f"• {b}" for b in slide["bullets"])
                content_requests.extend(create_text_box(
                    slide_id, bullet_text,
                    x=50, y=100, width=620, height=280,
                    font_size=18, bold=False, color=text_color
                ))
    
    # Execute content updates
    service.presentations().batchUpdate(
        presentationId=pres_id,
        body={"requests": content_requests}
    ).execute()
    
    print(f"✅ Added content to all slides")
    print(f"\n🔗 Open your presentation:")
    print(f"   https://docs.google.com/presentation/d/{pres_id}/edit")
    
    # Return info for HISTORY.md
    return {
        "presentation_id": pres_id,
        "url": f"https://docs.google.com/presentation/d/{pres_id}/edit",
        "slide_ids": slide_ids,
        "slide_count": len(SLIDES_CONTENT)
    }

if __name__ == "__main__":
    result = main()
    print(f"\n📝 Add to HISTORY.md:")
    print(f"   Presentation ID: {result['presentation_id']}")
    print(f"   Slides created: {result['slide_count']}")
