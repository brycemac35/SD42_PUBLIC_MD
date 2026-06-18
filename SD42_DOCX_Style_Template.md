# SD42 Work Instruction — DOCX Style Template (v2.0)
### Hybrid Checklist + Appendix Format for use with GitHub Copilot, Microsoft 365 Copilot, or any AI code assistant

---

## Purpose

This document defines the exact visual style, colour palette, layout rules, and `docx` JavaScript code patterns used to generate **School District No. 42 (Maple Ridge & Pitt Meadows) Work Instructions (WI-) and How-To Guides (HT-)** as `.docx` files with a **hybrid structure**: a quick checklist procedure that references detailed step-by-step appendices.

### Document Structure Overview

**Hybrid Format:**
- **Main Procedure Section**: Quick checklist (2-column table) with step name + appendix reference
- **Appendices**: Detailed click-by-click instructions organized by procedure step
- **Title Page**: Separate cover sheet directing readers to start at the Procedure section
- **Backward Compatible**: All existing WI/HT styling, colours, and ISO 9001:2015 compliance maintained

---

## Toolchain

| Item | Detail |
|---|---|
| Language | JavaScript (Node.js) |
| Library | `docx` npm package (`npm install -g docx`) |
| Output | `.docx` file written via `Packer.toBuffer()` |
| Validation | Run `python scripts/office/validate.py output.docx` after generation |

**Required imports:**
```javascript
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak, LevelFormat, TabStopType, TabStopPosition, HyperlinkType
} = require('docx');
const fs = require('fs');
```

---

## Page Setup

```javascript
properties: {
  page: {
    size: { width: 12240, height: 15840 },          // Letter, portrait
    margin: { top: 1080, right: 1080, bottom: 1080, left: 1080, header: 708, footer: 708 }
  }
}
```

- All measurements in **DXA** (twentieths of a point). 1 inch = 1440 DXA.
- **Content width** = 12240 − 1080 − 1080 = **10080 DXA**

---

## Colour Palette

| Role | Hex | Usage |
|---|---|---|
| `NAVY` | `1F4E79` | Primary brand colour — header left cell, H1 background, table headers, borders |
| `NAVY_LIGHT` | `D6E4F0` | Header right cell background |
| `NAVY_MID` | `BDD7EE` | Header subtitle text colour |
| `GREY_ROW` | `F2F2F2` | Alternating table row fill (even rows) |
| `BORDER_CC` | `CCCCCC` | Data table cell borders |
| `TEXT_BODY` | `000000` | All body text |
| `TEXT_GREY` | `595959` | Footer text, secondary captions |
| `WARN_RED_BORDER` | `C00000` | WARNING callout left border |
| `WARN_RED_FILL` | `FDECEA` | WARNING callout background |
| `WARN_RED_TEXT` | `7B0000` | WARNING callout text |
| `WARN_YLW_BORDER` | `BF8F00` | NOTE callout left border |
| `WARN_YLW_FILL` | `FFF2CC` | NOTE callout background |
| `WARN_YLW_TEXT` | `5A3E00` | NOTE callout text |
| `WHITE` | `FFFFFF` | Table header text, general white |
| `HYPERLINK_BLUE` | `0563C1` | Hyperlink text colour |

---

## Typography

| Element | Font | Size (half-pts) | Weight | Colour |
|---|---|---|---|---|
| Body text | Arial | 22 (11pt) | Normal | `000000` |
| Heading 1 text | Arial | 26 (13pt) | Bold | `FFFFFF` on `NAVY` bg |
| Heading 2 text | Arial | 22 (11pt) | Bold | `1F4E79` |
| Heading 3 text | Arial | 20 (10pt) | Bold | `1F4E79` |
| Table header text | Arial | 18 (9pt) | Bold | `FFFFFF` |
| Table body text | Arial | 20 (10pt) | Normal | `000000` |
| Header title | Arial | 22 (11pt) | Bold | `FFFFFF` |
| Header subtitle | Arial | 18 (9pt) | Normal | `BDD7EE` |
| Header meta (right) | Arial | 16 (8pt) | Bold labels, normal values | `1F4E79` |
| Footer text | Arial | 16 (8pt) | Normal | `595959` |
| Callout title line | Arial | 22 (11pt) | Bold | varies by type |
| Callout body lines | Arial | 22 (11pt) | Normal | varies by type |
| Hyperlink text | Arial | varies | Normal | `0563C1` (underlined) |

---

## Document Styles Block

Define these in the `styles` key of the `Document` constructor:

```javascript
styles: {
  default: { document: { run: { font: "Arial", size: 22, color: "000000" } } },
  paragraphStyles: [
    {
      id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 26, bold: true, font: "Arial", color: "FFFFFF" },
      paragraph: {
        spacing: { before: 240, after: 120 }, outlineLevel: 0,
        shading: { fill: "1F4E79", type: ShadingType.CLEAR }
      }
    },
    {
      id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 22, bold: true, font: "Arial", color: "1F4E79" },
      paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 }
    },
    {
      id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 20, bold: true, font: "Arial", color: "1F4E79" },
      paragraph: { spacing: { before: 160, after: 60 }, outlineLevel: 2 }
    }
  ]
}
```

---

## Numbering Block

Define in the `numbering` key of the `Document` constructor:

```javascript
numbering: {
  config: [
    {
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 270 } } }
      }]
    },
    {
      reference: "steps",
      levels: [{
        level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 270 } } }
      }]
    },
    {
      reference: "appendix-steps",
      levels: [
        {
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 540, hanging: 270 } } }
        },
        {
          level: 1, format: LevelFormat.LOWER_LETTER, text: "%2.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1080, hanging: 270 } } }
        }
      ]
    }
  ]
}
```

---

## Title Page (NEW)

A separate cover sheet before the header. Styled centered with document metadata and a instruction line.

```javascript
function titlePage(docNo, title, author, date, scope) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 2400, after: 240 },
      children: [new TextRun({ text: title, font: "Arial", size: 32, bold: true, color: "1F4E79" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 480 },
      children: [new TextRun({ text: "School District No. 42 — Information Technology Services", font: "Arial", size: 18, color: "595959" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 60 },
      children: [new TextRun({ text: "Document Number:", font: "Arial", size: 22, bold: true, color: "1F4E79" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 160 },
      children: [new TextRun({ text: docNo, font: "Arial", size: 22, color: "000000" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 60 },
      children: [new TextRun({ text: "Author:", font: "Arial", size: 22, bold: true, color: "1F4E79" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 160 },
      children: [new TextRun({ text: author, font: "Arial", size: 22, color: "000000" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 60 },
      children: [new TextRun({ text: "Date:", font: "Arial", size: 22, bold: true, color: "1F4E79" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 160 },
      children: [new TextRun({ text: date, font: "Arial", size: 22, color: "000000" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 240 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: "1F4E79", space: 2 } },
      children: [new TextRun({ text: "Scope", font: "Arial", size: 22, bold: true, color: "1F4E79" })]
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 480 },
      indent: { left: 360, right: 360 },
      children: [new TextRun({ text: scope, font: "Arial", size: 22, color: "000000" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 600, after: 0 },
      children: [new TextRun({ text: "TO BEGIN:", font: "Arial", size: 22, bold: true, color: "C00000" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 0 },
      children: [new TextRun({ text: "See the Procedure section (Section 6) for quick steps.", font: "Arial", size: 22, color: "000000" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 0 },
      children: [new TextRun({ text: "Detailed instructions are in the Appendices.", font: "Arial", size: 22, color: "000000" })]
    }),
    new PageBreak()
  ];
}
```

---

## Header

The header is a **2-column borderless table** (no cell borders — `BorderStyle.NONE`). Runs on all pages after the title page.

| Column | Width | Fill | Content |
|---|---|---|---|
| Left | 6480 DXA | `1F4E79` (NAVY) | SD42 logo image (90×47px) + document title (white bold 22pt) + subtitle (18pt `BDD7EE`) |
| Right | 3600 DXA | `D6E4F0` (light blue) | Doc No, Author, Rev, Date — right-aligned, `1F4E79`, 16pt |

```javascript
new Header({
  children: [
    new Table({
      width: { size: 10080, type: WidthType.DXA },
      columnWidths: [6480, 3600],
      rows: [new TableRow({
        children: [
          // LEFT CELL
          new TableCell({
            borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
            shading: { fill: "1F4E79", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 160, right: 80 },
            width: { size: 6480, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                children: [
                  new ImageRun({ type: "png", data: logoData, transformation: { width: 90, height: 47 } }),
                  new TextRun({ text: "  YOUR DOCUMENT TITLE", font: "Arial", size: 22, bold: true, color: "FFFFFF" }),
                ]
              }),
              new Paragraph({
                children: [new TextRun({ text: "Subtitle / Scope Line", font: "Arial", size: 18, color: "BDD7EE" })]
              })
            ]
          }),
          // RIGHT CELL
          new TableCell({
            borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
            shading: { fill: "D6E4F0", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            width: { size: 3600, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              // Repeat pattern for: Doc No, Author, Rev, Date
              new Paragraph({ alignment: AlignmentType.RIGHT, children: [
                new TextRun({ text: "Doc No: ", font: "Arial", size: 16, bold: true, color: "1F4E79" }),
                new TextRun({ text: "WI-XXXX-001",         font: "Arial", size: 16, color: "1F4E79" }),
              ]}),
              new Paragraph({ alignment: AlignmentType.RIGHT, children: [
                new TextRun({ text: "Bryce M", font: "Arial", size: 16, bold: true, color: "1F4E79" }),
              ]}),
              new Paragraph({ alignment: AlignmentType.RIGHT, children: [
                new TextRun({ text: "Rev: ", font: "Arial", size: 16, bold: true, color: "1F4E79" }),
                new TextRun({ text: "1.0",  font: "Arial", size: 16, color: "1F4E79" }),
              ]}),
              new Paragraph({ alignment: AlignmentType.RIGHT, children: [
                new TextRun({ text: "Date: ", font: "Arial", size: 16, bold: true, color: "1F4E79" }),
                new TextRun({ text: "YYYY-MM-DD", font: "Arial", size: 16, color: "1F4E79" }),
              ]}),
            ]
          })
        ]
      })]
    })
  ]
})
```

---

## Footer

Single paragraph with a top border (`1F4E79`), left-aligned organization name, right-tab page number.

```javascript
new Footer({
  children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: "1F4E79", space: 1 } },
    tabStops: [{ type: TabStopType.RIGHT, position: 10080 }],
    children: [
      new TextRun({ text: "School District No. 42 — Information Technology Services", font: "Arial", size: 16, color: "595959" }),
      new TextRun({ text: "\t", font: "Arial" }),
      new TextRun({ text: "Page ", font: "Arial", size: 16, color: "595959" }),
      new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "595959" }),
      new TextRun({ text: " of ", font: "Arial", size: 16, color: "595959" }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 16, color: "595959" }),
    ]
  })]
})
```

---

## Content Elements

### Heading 1 — `h1(text)`
White bold text on solid NAVY background. Applied via paragraph `style: "Heading1"` with shading.

```javascript
function h1(text) {
  return new Paragraph({
    style: "Heading1",
    shading: { fill: "1F4E79", type: ShadingType.CLEAR },
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: "FFFFFF" })]
  });
}
```

### Heading 2 — `h2(text)`
NAVY bold text with a bottom border underline. No background fill.

```javascript
function h2(text) {
  return new Paragraph({
    style: "Heading2",
    spacing: { before: 200, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "1F4E79", space: 2 } },
    children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: "1F4E79" })]
  });
}
```

### Heading 3 — `h3(text)` (NEW for Appendix sections)
NAVY bold text, smaller than H2, no border. Used for appendix section headers.

```javascript
function h3(text) {
  return new Paragraph({
    style: "Heading3",
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20, bold: true, color: "1F4E79" })]
  });
}
```

### Body Paragraph — `body(text, opts)`

```javascript
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "000000", ...opts })]
  });
}
```

### Bullet Point — `bullet(text)`
Uses the `"bullets"` numbering reference.

```javascript
function bullet(text, bold = false) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "000000", bold })]
  });
}
```

### Numbered Step — `step(text)`
Uses the `"steps"` numbering reference. For main procedure steps.

```javascript
function step(text) {
  return new Paragraph({
    numbering: { reference: "steps", level: 0 },
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "000000" })]
  });
}
```

### Appendix Step with Sub-steps — `appendixStep(text)` and `appendixSubStep(text)` (NEW)
Nested numbering for detailed click-by-click instructions in appendices (1, 1.a, 1.b, etc.).

```javascript
function appendixStep(text) {
  return new Paragraph({
    numbering: { reference: "appendix-steps", level: 0 },
    spacing: { before: 100, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "000000", bold: true })]
  });
}

function appendixSubStep(text) {
  return new Paragraph({
    numbering: { reference: "appendix-steps", level: 1 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "000000" })]
  });
}
```

### Hyperlinked Appendix Reference (NEW)
Creates clickable "Appendix X.X" text that jumps to that section. Used in the Procedure checklist table.

```javascript
function appendixLink(appendixId, pageNum) {
  return new TextRun({
    text: `Appendix ${appendixId}`,
    font: "Arial", size: 20, color: "0563C1", underline: {},
    link: { type: HyperlinkType.INTERNAL, anchor: `appendix-${appendixId}` }
  });
}
// Usage in a table cell:
new Paragraph({
  children: [
    appendixLink("1.1", 8),
    new TextRun({ text: ", p. 8", font: "Arial", size: 20, color: "000000" })
  ]
})
```

### Spacer — `sp(pts)`

```javascript
function sp(pts = 120) {
  return new Paragraph({ spacing: { before: pts, after: 0 }, children: [new TextRun("")] });
}
```

### Page Break — `pb()`

```javascript
function pb() {
  return new Paragraph({ children: [new PageBreak()] });
}
```

---

## Callout Boxes (WARNING / NOTE)

Callouts are **paragraph-level** elements, not tables. Each line is its own `Paragraph` with a left border and background shading. Returns an **array of Paragraphs** — spread with `...` when inserting into `children`.

```javascript
function callout(lines, type = "note") {
  const borderColor = type === "danger" ? "C00000" : "BF8F00";
  const fillColor   = type === "danger" ? "FDECEA" : "FFF2CC";
  const textColor   = type === "danger" ? "7B0000" : "5A3E00";

  return lines.map((line, i) => new Paragraph({
    spacing: { before: 80, after: 80 },
    indent: { left: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 10, color: borderColor, space: 0 } },
    shading: { fill: fillColor, type: ShadingType.CLEAR },
    children: [new TextRun({
      text: line, font: "Arial", size: 22,
      color: textColor, bold: i === 0    // first line is bold (acts as title)
    })]
  }));
}
```

**Usage:**
```javascript
// WARNING (danger — red)
...callout([
  "WARNING: Power off both devices before beginning.",
  "Failure to do so may result in personal injury or hardware damage.",
], "danger"),

// NOTE (yellow)
...callout([
  "NOTE: Settings are starting points only.",
  "Always test on scrap material before a final job.",
], "note"),
```

---

## Procedure Checklist Table (NEW) — `procedureChecklist(steps)`

2-column table with step names and hyperlinked appendix references. Used in **Section 6: Procedure**.

```javascript
function procedureChecklist(steps) {
  // steps = array of { stepName, appendixId, pageNum }
  const colWidths = [6500, 3580];
  const headers = ["Step", "Detailed Instructions"];

  const headerRow = new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders: { top: navyBorder, bottom: navyBorder, left: navyBorder, right: navyBorder },
      shading: { fill: "1F4E79", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      width: { size: colWidths[i], type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        children: [new TextRun({ text: h, font: "Arial", size: 18, bold: true, color: "FFFFFF" })]
      })]
    }))
  });

  const dataRows = steps.map((step, ri) => new TableRow({
    children: [
      // Step name (left column)
      new TableCell({
        borders: { top: greyBorder, bottom: greyBorder, left: greyBorder, right: greyBorder },
        shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        width: { size: colWidths[0], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          children: [new TextRun({ text: step.stepName, font: "Arial", size: 20, color: "000000" })]
        })]
      }),
      // Appendix hyperlink (right column)
      new TableCell({
        borders: { top: greyBorder, bottom: greyBorder, left: greyBorder, right: greyBorder },
        shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        width: { size: colWidths[1], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [
            new TextRun({
              text: `Appendix ${step.appendixId}`,
              font: "Arial", size: 20, color: "0563C1", underline: {},
              link: { type: HyperlinkType.INTERNAL, anchor: `appendix-${step.appendixId}` }
            }),
            new TextRun({ text: `, p. ${step.pageNum}`, font: "Arial", size: 20, color: "000000" })
          ]
        })]
      })
    ]
  }));

  return new Table({
    width: { size: 10080, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows]
  });
}
```

**Example usage:**
```javascript
procedureChecklist([
  { stepName: "1. Verify device hash is in Intune", appendixId: "1.1", pageNum: 8 },
  { stepName: "2. Enroll device via Autopilot", appendixId: "1.2", pageNum: 9 },
  { stepName: "3. Assign device to employee group", appendixId: "1.3", pageNum: 10 },
])
```

---

## Data Tables

NAVY header row, alternating `FFFFFF` / `F2F2F2` rows, `CCCCCC` cell borders.

```javascript
function dataTable(headers, rows, colWidths) {
  const headerRow = new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders: { top: navyBorder, bottom: navyBorder, left: navyBorder, right: navyBorder },
      shading: { fill: "1F4E79", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      width: { size: colWidths[i], type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        children: [new TextRun({ text: h, font: "Arial", size: 18, bold: true, color: "FFFFFF" })]
      })]
    }))
  });

  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders: { top: greyBorder, bottom: greyBorder, left: greyBorder, right: greyBorder },
      shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 120, right: 120 },
      width: { size: colWidths[ci], type: WidthType.DXA },
      children: [new Paragraph({
        children: [new TextRun({ text: cell, font: "Arial", size: 20, color: "000000" })]
      })]
    }))
  }));

  return new Table({
    width: { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows]
  });
}
```

---

## Sign-Off / Document Control Table

4-column layout: label (NAVY) | value (white) | label (NAVY) | value (white). Two field pairs per row.

```javascript
function signoffTable(pairs) {
  // pairs = array of [label, value] — rendered 2 pairs per row
  const rows = [];
  for (let i = 0; i < pairs.length; i += 2) {
    const left  = pairs[i];
    const right = pairs[i + 1] || ["", ""];
    rows.push(new TableRow({
      children: [
        new TableCell({
          borders: greyBorders,
          shading: { fill: "1F4E79", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          width: { size: 2000, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: left[0], font: "Arial", size: 18, bold: true, color: "FFFFFF" })] })]
        }),
        new TableCell({
          borders: greyBorders, shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          width: { size: 3040, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: left[1], font: "Arial", size: 20 })] })]
        }),
        new TableCell({
          borders: greyBorders, shading: { fill: "1F4E79", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          width: { size: 2000, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: right[0], font: "Arial", size: 18, bold: true, color: "FFFFFF" })] })]
        }),
        new TableCell({
          borders: greyBorders, shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          width: { size: 3040, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: right[1], font: "Arial", size: 20 })] })]
        }),
      ]
    }));
  }
  return new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: [2000, 3040, 2000, 3040], rows });
}
```

---

## Verification / Checklist Table

4 columns: `#` | `Verification Check` | `Pass ☐` | `Fail ☐`. NAVY header, alternating rows.

```javascript
function verificationTable(checks) {
  const colWidths = [600, 7080, 1200, 1200];
  const headers   = ["#", "Verification Check", "Pass", "Fail"];

  const headerRow = new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders: navyBorders,
      shading: { fill: "1F4E79", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      width: { size: colWidths[i], type: WidthType.DXA },
      children: [new Paragraph({ children: [new TextRun({ text: h, font: "Arial", size: 18, bold: true, color: "FFFFFF" })] })]
    }))
  });

  const dataRows = checks.map((text, ri) => new TableRow({
    children: [
      // # column
      new TableCell({
        borders: greyBorders, shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 600, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: String(ri + 1), font: "Arial", size: 20 })] })]
      }),
      // check description
      new TableCell({
        borders: greyBorders, shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 7080, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 20 })] })]
      }),
      // Pass checkbox
      new TableCell({
        borders: greyBorders, shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 1200, type: WidthType.DXA },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "\u2610", font: "Arial", size: 22 })] })]
      }),
      // Fail checkbox
      new TableCell({
        borders: greyBorders, shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 1200, type: WidthType.DXA },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "\u2610", font: "Arial", size: 22 })] })]
      }),
    ]
  }));

  return new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths, rows: [headerRow, ...dataRows] });
}
```

---

## ISO 9001:2015 Required Sections

Every SD42 work instruction **must** include the following sections in this order:

| # | Section | ISO 9001 Clause | Notes |
|---|---|---|---|
| 1 | Document Control | 7.5.2 | Doc No, Revision, Author, Effective Date, Review Cycle, Classification |
| 2 | Purpose & Scope | 8.5.1 | State what the procedure covers and who it applies to |
| 3 | Referenced Documents | 7.5.3 | List all related manuals, standards, or procedures |
| 4 | Required Tools & Materials | 7.1.3 / 7.1.4 | Table format preferred |
| 5 | Safety & Precautions | 8.5.1 | Use `danger` callout for warnings; `note` callout for cautions |
| 6 | Procedure (Checklist) | 8.5.1 | **NEW**: 2-column table with step names + appendix references |
| 7 | Post-Job Verification | 8.6 | Use verification table with Pass/Fail checkboxes |
| 8 | Troubleshooting | 8.7 | Symptom / Likely Cause / Corrective Action table |
| 9 | Quality Record / Sign-Off | 7.5.3 | Must be completed and retained ≥ 1 year per SD42 retention policy |
| **Appendices** | Detailed Instructions | 8.5.1 (expanded) | Click-by-click steps with nested numbering (1.a, 1.b, 1.c, etc.) |

**Appendix Numbering:**
- Appendix 1.1, 1.2, 1.3, etc. — tied to main procedure step 1
- Appendix 2.1, 2.2, etc. — tied to main procedure step 2
- Each appendix uses H3 heading: "Appendix 1.1: [Description]"
- Nested numbering within appendices: 1, 1.a, 1.b, 1.c, etc.

**ISO 9001 QMS reference statement** — include in Section 2 (Purpose & Scope):
> *This document is maintained in accordance with the SD42 Quality Management System (QMS) and aligns with ISO 9001:2015 requirements for documented work instructions (Clause 8.5.1).*

**ISO 9001 records retention statement** — include in the Quality Record / Sign-Off section:
> *This section must be completed and retained per the SD42 QMS records control requirement (ISO 9001:2015 Clause 7.5.3). Retain this record for a minimum of one (1) year, or per current SD42 records retention policy, whichever is longer.*

---

## Closing Brand Block

Place at the very end of the document, after the sign-off table:

```javascript
new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 160 },
  border: { top: { style: BorderStyle.SINGLE, size: 4, color: "1F4E79", space: 4 } },
  children: [
    new TextRun({ text: "School District No. 42 (Maple Ridge & Pitt Meadows)  —  IT Services", font: "Arial", size: 18, bold: true, color: "1F4E79" }),
    new TextRun({ break: 1 }),
    new TextRun({ text: "Learning Today, Leading Tomorrow", font: "Arial", size: 18, italics: true, color: "595959" }),
    new TextRun({ break: 1 }),
    new TextRun({ text: "WI-XXXX-001 Rev 1.0  —  YYYY-MM-DD  —  For internal use only", font: "Arial", size: 16, color: "AAAAAA" }),
  ]
})
```

---

## Document Numbering Convention

### Format

All SD42 IT Services documents follow this naming structure:

**`[TYPE]-[CATEGORY]-[NNN]`**

Where:
- **[TYPE]** = Document type (`WI` = Work Instruction, `HT` = How-To Guide)
- **[CATEGORY]** = Technology/domain category (2-letter code)
- **[NNN]** = Sequential document number (always 3 digits, zero-padded: 001, 002, ..., 999)

### Category Codes

| Category Code | Full Name | Examples |
|---|---|---|
| `HW` | Hardware | Device swaps, WLAN card transfers, OptiPlex upgrades |
| `SW` | Software | Windows provisioning, app deployment, OS updates |
| `NET` | Network | VPN setup, IP configuration, network drive mapping |
| `SEC` | Security | Password management, MFA setup, encryption procedures |
| `GEN` | General / Cross-functional | Document management, user onboarding, general maintenance |
| `LIB` | Library Systems | iPad/device provisioning, library-specific workflows |
| `MOBILE` | Mobile Devices | iOS/Android management, MDM enrollment, app configuration |
| `COLLAB` | Collaboration Tools | Teams, SharePoint, OneDrive setup and administration |
| `PRINT` | Printing & Scanning | Print server setup, scanner configuration, troubleshooting |
| `PHONE` | Telephony | Phone provisioning, VoIP configuration, call routing |

*Note: New categories can be added as the documentation library expands. Coordinate with the IT Services manager before introducing a new category code.*

### Examples

| Document | Type | Category | Purpose | Example Doc No |
|---|---|---|---|---|
| OptiPlex device swap procedure | WI | Hardware | Quick reference for experienced techs | `WI-HW-001` |
| OptiPlex device swap — detailed | HT | Hardware | Learning guide with appendices for newer techs | `HT-HW-001` |
| Windows device provisioning via Intune | WI | Software | Fast provisioning steps for field techs | `WI-SW-001` |
| Windows device provisioning — detailed | HT | Software | In-depth guide with troubleshooting appendices | `HT-SW-001` |
| Network drive mapping (Windows & macOS) | WI | Network | Quick how-to for both platforms | `WI-NET-001` |
| iPad migration from Jamf to Intune | WI | Mobile | Streamlined iPad enrollment process | `WI-MOBILE-001` |

### When to Use WI vs. HT

**Use Work Instruction (WI-)** when:
- Procedure is straightforward and well-defined
- Audience is experienced technicians or field staff familiar with the domain
- Document should be concise (3–6 pages) and serve as quick reference
- Minimal contextual explanation is needed

**Use How-To Guide (HT-)** when:
- Procedure is complex with many decision points or variations
- Audience includes newer technicians, trainees, or staff learning the domain
- Document should provide detailed step-by-step instructions in appendices
- Click-by-click guidance is valuable for users unfamiliar with the system
- Troubleshooting scenarios and edge cases need explanation

### Versioning & Revision

- **First release**: Always `Rev 1.0` (in header meta and closing block)
- **Minor updates** (typos, clarifications, colour corrections): Increment 0.1 (e.g., 1.0 → 1.1)
- **Major updates** (procedure changes, new appendices, restructure): Increment 1.0 (e.g., 1.0 → 2.0)
- **Effective Date**: Format as `YYYY-MM-DD` (ISO 8601)
- **Review Cycle**: Typically `Annual`; adjust if the procedure changes frequently

### Document Numbering Authority

- **Author**: Proposes document number based on category and next available sequence
- **IT Services Manager**: Approves document number to avoid conflicts across the library
- **GitHub Repository**: `/brycemac35/SD42_PUBLIC_MD` is the single source of truth; all document numbers are logged in the repository README

---

## AI Prompt — How to Use This Template

### For Work Instructions (WI-)

When asking an AI assistant to generate a new SD42 Work Instruction, use this prompt:

```
You are generating a School District No. 42 (SD42) Work Instruction (WI-) as a .docx file
using the Node.js `docx` npm library.

Follow the SD42_DOCX_Style_Template.md v2.0 specification for Work Instructions EXACTLY:
- Hybrid structure: Title page + quick checklist procedure + brief sections (no appendices)
- Page size: Letter (12240 × 15840 DXA), margins 1080 DXA all sides
- Colours: NAVY=1F4E79, NAVY_LIGHT=D6E4F0, NAVY_MID=BDD7EE, GREY_ROW=F2F2F2
- Font: Arial throughout; body 22pt, H1 26pt, H2 22pt
- Title page: Separate cover with Doc No, Title, Author, Date, Scope + instruction to start at Procedure
- Header: 2-column borderless table — left NAVY (logo + title), right D6E4F0 (meta)
- Footer: NAVY top border, left org name, right tab page number
- Procedure (Section 6): Simple numbered steps (1, 2, 3) — NOT the checklist table format
- Include all ISO 9001:2015 required sections for WI (Document Control, Purpose & Scope, Referenced Documents,
  Tools & Materials, Safety, Procedure, Verification, Troubleshooting, Quality Record/Sign-Off)
- End with the SD42 closing brand block
- NO appendices (this is a quick-reference document)

The document to create is: [DESCRIBE YOUR WI DOCUMENT HERE]
Document number: [WI-CATEGORY-NNN]
Author: [NAME]
```

### For How-To Guides (HT-)

When asking an AI assistant to generate a new SD42 How-To Guide, use this prompt:

```
You are generating a School District No. 42 (SD42) How-To Guide (HT-) as a .docx file
using the Node.js `docx` npm library.

Follow the SD42_DOCX_Style_Template.md v2.0 specification for How-To Guides EXACTLY:
- Hybrid structure: Title page + checklist procedure (with appendix references) + detailed appendices
- Page size: Letter (12240 × 15840 DXA), margins 1080 DXA all sides
- Colours: NAVY=1F4E79, NAVY_LIGHT=D6E4F0, NAVY_MID=BDD7EE, GREY_ROW=F2F2F2
- Font: Arial throughout; body 22pt, H1 26pt, H2 22pt, H3 20pt
- Title page: Separate cover with Doc No, Title, Author, Date, Scope + instruction to start at Procedure
- Header: 2-column borderless table — left NAVY (logo + title), right D6E4F0 (meta)
- Footer: NAVY top border, left org name, right tab page number
- Procedure (Section 6): 2-column CHECKLIST TABLE format using procedureChecklist() function
  - Left column: "1. Step name", "2. Step name", etc.
  - Right column: Hyperlinked "Appendix 1.1, p. X", "Appendix 1.2, p. Y", etc.
  - Each step hyperlinks to its appendix section with internal anchor
- Appendices (after Sign-Off): Detailed click-by-click instructions
  - H3 headers: "Appendix 1.1: [Description]"
  - Nested numbering using appendixStep() and appendixSubStep()
  - Format: 1, 1.a, 1.b, 1.c, 2, 2.a, 2.b, etc.
- Include all ISO 9001:2015 required sections for HT (Document Control, Purpose & Scope, Referenced Documents,
  Tools & Materials, Safety, Procedure Checklist, Verification, Troubleshooting, Quality Record/Sign-Off)
- End with the SD42 closing brand block

The document to create is: [DESCRIBE YOUR HT DOCUMENT HERE]
Document number: [HT-CATEGORY-NNN]
Author: [NAME]
Audience: [Field technicians / new staff / etc.]
```

---

## Quick Reference — Element Summary

| Element | Function | Returns | WI | HT |
|---|---|---|---|---|
| `titlePage(docNo, title, author, date, scope)` | Cover sheet with instructions | `Paragraph[]` | ✓ | ✓ |
| `h1(text)` | Section heading (NAVY bg, white text) | `Paragraph` | ✓ | ✓ |
| `h2(text)` | Sub-section heading (NAVY text, underline) | `Paragraph` | ✓ | ✓ |
| `h3(text)` | Appendix section heading (NAVY text) | `Paragraph` | — | ✓ |
| `body(text, opts)` | Body paragraph | `Paragraph` | ✓ | ✓ |
| `bullet(text)` | Bullet list item | `Paragraph` | ✓ | ✓ |
| `step(text)` | Numbered procedure step | `Paragraph` | ✓ | — |
| `appendixStep(text)` | Main appendix step | `Paragraph` | — | ✓ |
| `appendixSubStep(text)` | Sub-step in appendix (1.a, 1.b) | `Paragraph` | — | ✓ |
| `sp(pts)` | Vertical spacer | `Paragraph` | ✓ | ✓ |
| `pb()` | Page break | `Paragraph` | ✓ | ✓ |
| `...callout(lines, "danger")` | Red WARNING block | `Paragraph[]` | ✓ | ✓ |
| `...callout(lines, "note")` | Yellow NOTE block | `Paragraph[]` | ✓ | ✓ |
| `procedureChecklist(steps)` | Checklist table with appendix links | `Table` | — | ✓ |
| `dataTable(headers, rows, colWidths)` | Data/reference table | `Table` | ✓ | ✓ |
| `signoffTable(pairs)` | Document control / sign-off | `Table` | ✓ | ✓ |
| `verificationTable(checks)` | Pass/Fail checklist | `Table` | ✓ | ✓ |

---

## Key Differences: WI- vs. HT-

| Aspect | WI (Work Instruction) | HT (How-To Guide) |
|---|---|---|
| **Audience** | Experienced technicians; quick reference | Newer/learning techs; detailed guidance |
| **Length** | Concise; 3–6 pages | Detailed; 8–15+ pages |
| **Procedure Format** | Simple numbered steps (1, 2, 3) | 2-column checklist table + appendix links |
| **Detail Level** | Minimal; assumes prior knowledge | Extensive; click-by-click in appendices |
| **Appendices** | None | Multiple (1.1, 1.2, 1.3, etc.) with nested steps |
| **Use Case** | "How do I do this quickly?" | "Help me understand and do this step-by-step" |
| **Title Page** | Yes, basic | Yes, basic |
| **Target User** | Experienced field techs | New hires, trainees, learning techs |

---

*SD42 IT Services — Internal use only. Template version 2.0 (Restructured with title page, hybrid checklist + appendix format)*
