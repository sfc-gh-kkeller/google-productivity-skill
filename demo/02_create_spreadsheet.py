#!/usr/bin/env python3
"""
=============================================================================
Demo 02: Create and Populate a Spreadsheet
=============================================================================
Creates a Google Sheet with sample data, formatting, and a chart.
Demonstrates: creating sheets, writing data, formatting, charts.

Run: pixi run python demo/02_create_spreadsheet.py
=============================================================================
"""

from googleapiclient.discovery import build
import google.auth

# =============================================================================
# CONFIGURATION
# =============================================================================

QUOTA_PROJECT = "your-gcp-project"  # Replace with your GCP project ID
SPREADSHEET_TITLE = "Q1 2026 Sales Dashboard"

# Sample data
SALES_DATA = [
    ["Region", "Jan", "Feb", "Mar", "Q1 Total"],
    ["North America", 125000, 142000, 158000, "=SUM(B2:D2)"],
    ["Europe", 98000, 105000, 118000, "=SUM(B3:D3)"],
    ["Asia Pacific", 87000, 95000, 112000, "=SUM(B4:D4)"],
    ["Latin America", 45000, 52000, 61000, "=SUM(B5:D5)"],
    ["Total", "=SUM(B2:B5)", "=SUM(C2:C5)", "=SUM(D2:D5)", "=SUM(E2:E5)"]
]

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def get_sheets_service():
    """Get authenticated Sheets API service."""
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds, _ = google.auth.default(scopes=scopes)
    creds = creds.with_quota_project(QUOTA_PROJECT)
    return build("sheets", "v4", credentials=creds)

def rgb_to_color(r, g, b):
    """Convert RGB (0-255) to Google Sheets color format (0-1)."""
    return {"red": r/255, "green": g/255, "blue": b/255}

# =============================================================================
# MAIN
# =============================================================================

def main():
    print("📊 Creating spreadsheet...")
    
    # Get API service
    service = get_sheets_service()
    
    # Create spreadsheet
    spreadsheet = service.spreadsheets().create(
        body={
            "properties": {"title": SPREADSHEET_TITLE},
            "sheets": [{"properties": {"title": "Sales Data"}}]
        }
    ).execute()
    
    sheet_id = spreadsheet["spreadsheetId"]
    first_sheet_id = spreadsheet["sheets"][0]["properties"]["sheetId"]
    
    print(f"✅ Created spreadsheet: {sheet_id}")
    
    # Write data
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range="Sales Data!A1",
        valueInputOption="USER_ENTERED",
        body={"values": SALES_DATA}
    ).execute()
    
    print("✅ Added sales data")
    
    # Format the spreadsheet
    requests = [
        # Format header row (bold, blue background)
        {
            "repeatCell": {
                "range": {
                    "sheetId": first_sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 5
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": rgb_to_color(28, 135, 237),  # Primary blue
                        "textFormat": {
                            "bold": True,
                            "foregroundColor": rgb_to_color(255, 255, 255)
                        },
                        "horizontalAlignment": "CENTER"
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
            }
        },
        # Format total row (bold, light gray background)
        {
            "repeatCell": {
                "range": {
                    "sheetId": first_sheet_id,
                    "startRowIndex": 5,
                    "endRowIndex": 6,
                    "startColumnIndex": 0,
                    "endColumnIndex": 5
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": rgb_to_color(240, 240, 240),
                        "textFormat": {"bold": True}
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)"
            }
        },
        # Format numbers as currency
        {
            "repeatCell": {
                "range": {
                    "sheetId": first_sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 6,
                    "startColumnIndex": 1,
                    "endColumnIndex": 5
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "CURRENCY",
                            "pattern": "$#,##0"
                        }
                    }
                },
                "fields": "userEnteredFormat.numberFormat"
            }
        },
        # Auto-resize columns
        {
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": first_sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 5
                }
            }
        },
        # Add conditional formatting (highlight high values in green)
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": first_sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": 5,
                        "startColumnIndex": 1,
                        "endColumnIndex": 4
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "NUMBER_GREATER_THAN_EQ",
                            "values": [{"userEnteredValue": "100000"}]
                        },
                        "format": {
                            "backgroundColor": rgb_to_color(198, 239, 206)  # Light green
                        }
                    }
                },
                "index": 0
            }
        },
        # Freeze header row
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": first_sheet_id,
                    "gridProperties": {"frozenRowCount": 1}
                },
                "fields": "gridProperties.frozenRowCount"
            }
        }
    ]
    
    # Add a chart
    chart_request = {
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Q1 Sales by Region",
                    "basicChart": {
                        "chartType": "COLUMN",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Region"},
                            {"position": "LEFT_AXIS", "title": "Sales ($)"}
                        ],
                        "domains": [{
                            "domain": {
                                "sourceRange": {
                                    "sources": [{
                                        "sheetId": first_sheet_id,
                                        "startRowIndex": 1,
                                        "endRowIndex": 5,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 1
                                    }]
                                }
                            }
                        }],
                        "series": [
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [{
                                            "sheetId": first_sheet_id,
                                            "startRowIndex": 1,
                                            "endRowIndex": 5,
                                            "startColumnIndex": 1,
                                            "endColumnIndex": 2
                                        }]
                                    }
                                },
                                "targetAxis": "LEFT_AXIS",
                                "color": rgb_to_color(28, 135, 237)
                            },
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [{
                                            "sheetId": first_sheet_id,
                                            "startRowIndex": 1,
                                            "endRowIndex": 5,
                                            "startColumnIndex": 2,
                                            "endColumnIndex": 3
                                        }]
                                    }
                                },
                                "targetAxis": "LEFT_AXIS",
                                "color": rgb_to_color(0, 214, 224)
                            },
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [{
                                            "sheetId": first_sheet_id,
                                            "startRowIndex": 1,
                                            "endRowIndex": 5,
                                            "startColumnIndex": 3,
                                            "endColumnIndex": 4
                                        }]
                                    }
                                },
                                "targetAxis": "LEFT_AXIS",
                                "color": rgb_to_color(23, 28, 46)
                            }
                        ],
                        "headerCount": 1
                    }
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": first_sheet_id,
                            "rowIndex": 8,
                            "columnIndex": 0
                        },
                        "offsetXPixels": 0,
                        "offsetYPixels": 0,
                        "widthPixels": 600,
                        "heightPixels": 350
                    }
                }
            }
        }
    }
    requests.append(chart_request)
    
    # Execute all formatting
    service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": requests}
    ).execute()
    
    print("✅ Applied formatting and added chart")
    print(f"\n🔗 Open your spreadsheet:")
    print(f"   https://docs.google.com/spreadsheets/d/{sheet_id}/edit")
    
    return {
        "spreadsheet_id": sheet_id,
        "url": f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    }

if __name__ == "__main__":
    result = main()
    print(f"\n📝 Spreadsheet ID: {result['spreadsheet_id']}")
