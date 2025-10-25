# PPTX Export MCP Server

A Model Context Protocol (MCP) server that provides PowerPoint (PPTX) export functionality. This server allows you to create presentations with text, tables, charts, images, and shapes.

## Features

- Create PowerPoint presentations programmatically
- Support for multiple slide layouts (16x9, 16x10, 4x3)
- Add various elements to slides:
  - **Text**: Formatted text with fonts, colors, sizes
  - **Tables**: Data tables with customizable styling
  - **Charts**: Bar, line, pie, area, scatter, bubble, doughnut, radar charts
  - **Images**: Local files or base64 data URIs
  - **Shapes**: Rectangles, ellipses, triangles, arrows, etc.
- Slide background colors
- Presentation metadata (author, title, subject)
- File size reporting
- UUID-based unique filenames
- Saves files to `/tmp/protex-intelligence-file-exports/`

## Installation

```bash
# Clone and install
git clone <repository-url>
cd pptx-export-mcp
uv pip install -e .
```

## Dependencies

This package uses `python-pptx` for PowerPoint file generation, which is a pure Python library with no system dependencies.

## Usage

The server provides one tool:

### `pptx_export`

Exports slide data to PowerPoint (PPTX) format.

**Parameters:**
- `slides` (required): Array of slide objects containing elements
- `filename` (optional): Filename for the exported file (without extension), defaults to "output"
- `description` (optional): Description of the file contents
- `options` (optional): Presentation options
  - `layout`: Slide layout - "16x9", "16x10", "4x3" (default: "16x9")
  - `author`: Presentation author name
  - `title`: Presentation title
  - `subject`: Presentation subject

**Slide Structure:**
```json
{
  "backgroundColor": "FF0000",
  "elements": [
    {
      "type": "text",
      "text": "Hello World",
      "options": {
        "x": 1, "y": 1, "w": 8, "h": 1,
        "fontSize": 24, "color": "000000", "bold": true
      }
    },
    {
      "type": "table",
      "rows": [
        ["Header 1", "Header 2"],
        ["Data 1", "Data 2"]
      ],
      "options": {"x": 1, "y": 3, "w": 8, "h": 3}
    },
    {
      "type": "chart",
      "chartType": "bar",
      "chartData": [
        {
          "name": "Series 1",
          "labels": ["A", "B", "C"],
          "values": [10, 20, 30]
        }
      ],
      "options": {"x": 1, "y": 1, "w": 8, "h": 5}
    }
  ]
}
```

**Element Types:**
- `text`: Text boxes with formatting options
- `table`: Data tables with rows and columns
- `chart`: Charts (bar, line, pie, area, scatter, bubble, doughnut, radar)
- `image`: Images from file paths or base64 data URIs
- `shape`: Geometric shapes (rectangle, ellipse, triangle, etc.)

**Common Options:**
- `x`, `y`: Position in inches from top-left
- `w`, `h`: Width and height in inches
- `fontSize`: Font size in points
- `color`: Text color (hex without #)
- `bold`, `italic`: Text formatting
- `fill`: Shape fill color (hex without #)
- `line`: Shape outline color (hex without #)

**Example:**
```json
{
  "slides": [
    {
      "backgroundColor": "F0F0F0",
      "elements": [
        {
          "type": "text",
          "text": "Sales Report Q4 2024",
          "options": {
            "x": 1, "y": 1, "w": 10, "h": 1,
            "fontSize": 32, "color": "2C3E50", "bold": true
          }
        }
      ]
    }
  ],
  "filename": "sales_presentation",
  "options": {
    "layout": "16x9",
    "author": "John Doe",
    "title": "Q4 Sales Results"
  }
}
```

## Running the Server

```bash
pptx-export-mcp
```

## Development

```bash
# Install in development mode
uv pip install -e .

# Run tests (if available)
python -m pytest
```

## License

MIT