# SD42 Work Instruction & Comprehensive Guide — DOCX Style Template
### For use with GitHub Copilot, Microsoft 365 Copilot, or any AI code assistant

---

## Purpose

This document defines the exact visual style, colour palette, layout rules, and `docx` JavaScript code patterns used to generate **School District No. 42 (Maple Ridge & Pitt Meadows) Work Instructions (WI-) and Comprehensive Guides (CG-)** as `.docx` files. Feed this file to an AI assistant and ask it to create a new work instruction or comprehensive guide. The AI should follow every specification here exactly.

### Document Types

- **Work Instruction (WI-)**: Concise, procedural documents for experienced technicians. Focus: efficiency and quick reference.
- **Comprehensive Guide (CG-)**: Detailed, contextual documents for field technicians learning complex procedures. Focus: understanding, troubleshooting, and cross-functional awareness.

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
  VerticalAlign, PageNumber, PageBreak, LevelFormat, TabStopType, TabStopPosition
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
      reference: "steps-detailed",
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
    },
    {
      reference: "steps-b",
      levels: [{
        level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 270 } } }
      }]
    }
  ]
}
```

---

## Header

The header is a **2-column borderless table** (no cell borders — `BorderStyle.NONE`).

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
                new TextRun({ text: "CG-XXXX-001",         font: "Arial", size: 16, color: "1F4E79" }),
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

### Heading 3 — `h3(text)` [NEW — for CG documents]
NAVY bold text, smaller than H2, no border. Used for sub-sections within detailed procedure steps.

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
Uses the `"steps"` numbering reference. For concise WI documents.

```javascript
function step(text) {
  return new Paragraph({
    numbering: { reference: "steps", level: 0 },
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "000000" })]
  });
}
```

### Detailed Step with Sub-steps — `stepDetailed(text)` and `subStep(text)` [NEW — for CG documents]
Nested numbering (1, 1.a, 1.b, etc.) for comprehensive guides with click-by-click instructions.

```javascript
function stepDetailed(text) {
  return new Paragraph({
    numbering: { reference: "steps-detailed", level: 0 },
    spacing: { before: 100, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "000000", bold: true })]
  });
}

function subStep(text) {
  return new Paragraph({
    numbering: { reference: "steps-detailed", level: 1 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "000000" })]
  });
}
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

**Example call:**
```javascript
dataTable(
  ["Material", "Operation", "Notes"],
  [
    ["Wood (3mm)", "Engrave & Cut", "Watch for MDF resin"],
    ["Cast Acrylic", "Engrave & Cut", "Frosted finish when engraved"],
  ],
  [3200, 2160, 4720]   // column widths — must sum to ≤ 10080
)
```

---

## Diagnostic Tree Table [NEW — for CG documents]

For troubleshooting guides: **Symptom** | **Likely Cause** | **Corrective Action**. Useful in conjunction with flowchart visuals.

```javascript
function diagnosticTable(diagnostics) {
  // diagnostics = array of { symptom, cause, action }
  const colWidths = [2500, 3500, 4080];
  const headers = ["Symptom", "Likely Cause", "Corrective Action"];

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

  const dataRows = diagnostics.map((diag, ri) => new TableRow({
    children: [
      new TableCell({
        borders: { top: greyBorder, bottom: greyBorder, left: greyBorder, right: greyBorder },
        shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        width: { size: colWidths[0], type: WidthType.DXA },
        verticalAlign: VerticalAlign.TOP,
        children: [new Paragraph({ children: [new TextRun({ text: diag.symptom, font: "Arial", size: 20 })] })]
      }),
      new TableCell({
        borders: { top: greyBorder, bottom: greyBorder, left: greyBorder, right: greyBorder },
        shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        width: { size: colWidths[1], type: WidthType.DXA },
        verticalAlign: VerticalAlign.TOP,
        children: [new Paragraph({ children: [new TextRun({ text: diag.cause, font: "Arial", size: 20 })] })]
      }),
      new TableCell({
        borders: { top: greyBorder, bottom: greyBorder, left: greyBorder, right: greyBorder },
        shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        width: { size: colWidths[2], type: WidthType.DXA },
        verticalAlign: VerticalAlign.TOP,
        children: [new Paragraph({ children: [new TextRun({ text: diag.action, font: "Arial", size: 20 })] })]
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

---

## FAQ Section [NEW — for CG documents]

Format: **Q: [Question]** (bold) followed by **A: [Answer]** (normal). Use for common questions about procedure nuances.

```javascript
function faqItem(question, answer) {
  return [
    new Paragraph({
      spacing: { before: 120, after: 60 },
      children: [new TextRun({ text: `Q: ${question}`, font: "Arial", size: 22, bold: true, color: "1F4E79" })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: `A: ${answer}`, font: "Arial", size: 22, color: "000000" })]
    })
  ];
}
```

**Usage:**
```javascript
...faqItem(
  "Can I skip the verification step?",
  "No. Post-job verification confirms the procedure succeeded and is required by ISO 9001:2015 Clause 8.6. Document all results."
),
...faqItem(
  "What if the device is already updated?",
  "Check the current version in Settings > System > About. If already at target version, proceed to Post-Job Verification only."
)
```

---

## Cross-References Section [NEW — for CG documents]

Links to related Work Instructions (WI-) or other Comprehensive Guides (CG-). Format as a simple reference table.

```javascript
function crossReferences(docs) {
  // docs = array of { docNo, title, relationship }
  const colWidths = [1800, 5280, 3000];
  const headers = ["Document No.", "Title", "Relationship"];

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

  const dataRows = docs.map((doc, ri) => new TableRow({
    children: [
      new TableCell({
        borders: { top: greyBorder, bottom: greyBorder, left: greyBorder, right: greyBorder },
        shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        width: { size: colWidths[0], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: doc.docNo, font: "Arial", size: 20, bold: true, color: "1F4E79" })] })]
      }),
      new TableCell({
        borders: { top: greyBorder, bottom: greyBorder, left: greyBorder, right: greyBorder },
        shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        width: { size: colWidths[1], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: doc.title, font: "Arial", size: 20 })] })]
      }),
      new TableCell({
        borders: { top: greyBorder, bottom: greyBorder, left: greyBorder, right: greyBorder },
        shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F2F2F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        width: { size: colWidths[2], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: doc.relationship, font: "Arial", size: 20 })] })]
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

**Example call:**
```javascript
signoffTable([
  ["Document No.", "CG-XXXX-001"],  ["Revision",      "1.0"],
  ["Category",     "Software"],     ["Effective Date", "YYYY-MM-DD"],
  ["Author",        "Bryce M"],     ["Review Cycle",   "Annual"],
])
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

### Work Instructions (WI-)

Every SD42 work instruction **must** include the following sections in this order:

| # | Section | ISO 9001 Clause | Notes |
|---|---|---|---|
| 1 | Document Control | 7.5.2 | Doc No, Revision, Author, Effective Date, Review Cycle, Classification |
| 2 | Purpose & Scope | 8.5.1 | State what the procedure covers and who it applies to |
| 3 | Referenced Documents | 7.5.3 | List all related manuals, standards, or procedures |
| 4 | Required Tools & Materials | 7.1.3 / 7.1.4 | Table format preferred |
| 5 | Safety & Precautions | 8.5.1 | Use `danger` callout for warnings; `note` callout for cautions |
| 6–N | Procedure Steps | 8.5.1 | Numbered steps using `step()` function; bullets for non-sequential items |
| N+1 | Post-Job Verification | 8.6 | Use verification table with Pass/Fail checkboxes |
| N+2 | Troubleshooting | 8.7 | Symptom / Likely Cause / Corrective Action table |
| N+3 | Quality Record / Sign-Off | 7.5.3 | Must be completed and retained ≥ 1 year per SD42 retention policy |

### Comprehensive Guides (CG-)

Every SD42 comprehensive guide **must** include the following sections in this order:

| # | Section | ISO 9001 Clause | Notes |
|---|---|---|---|
| 1 | Document Control | 7.5.2 | Doc No, Revision, Author, Effective Date, Review Cycle, Classification |
| 2 | Purpose & Scope | 8.5.1 | Describe audience (field techs, new staff), context, and learning objectives |
| 3 | Background & Context | 8.5.1 (expanded) | Why this procedure matters; system/software overview; common scenarios |
| 4 | Referenced Documents | 7.5.3 | List all related manuals, standards, WI-, and CG- documents |
| 5 | Required Tools & Materials | 7.1.3 / 7.1.4 | Table format preferred |
| 6 | Safety & Precautions | 8.5.1 | Use `danger` callout for warnings; `note` callout for cautions |
| 7–N | Detailed Procedure Steps | 8.5.1 | Nested numbering (1., 1.a, 1.b, 1.c, etc.) with click-by-click sub-steps; screenshots encouraged |
| N+1 | Cross-References | 7.5.3 (new) | Related WI- and CG- documents with descriptions |
| N+2 | Frequently Asked Questions | 8.5.1 (clarification) | Common questions about procedure nuances, edge cases, optional steps |
| N+3 | Expanded Troubleshooting | 8.7 | Diagnostic tables + decision flowcharts for complex problems |
| N+4 | Post-Job Verification | 8.6 | Use verification table with Pass/Fail checkboxes |
| N+5 | Quality Record / Sign-Off | 7.5.3 | Must be completed and retained ≥ 1 year per SD42 retention policy |

**ISO 9001 QMS reference statement** — include in Section 2 (Purpose & Scope) for both WI and CG:
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
    new TextRun({ text: "CG-XXXX-001 Rev 1.0  —  YYYY-MM-DD  —  For internal use only", font: "Arial", size: 16, color: "AAAAAA" }),
  ]
})
```

---

## Document Numbering Convention

### Format

All SD42 IT Services documents follow this naming structure:

**`[TYPE]-[CATEGORY]-[NNN]`**

Where:
- **[TYPE]** = Document type (`WI` = Work Instruction, `CG` = Comprehensive Guide)
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
| OptiPlex device swap — detailed guide | CG | Hardware | Step-by-step learning guide for new techs | `CG-HW-001` |
| Windows device provisioning via Intune | WI | Software | Fast provisioning steps for field techs | `WI-SW-001` |
| Windows device provisioning — comprehensive | CG | Software | In-depth troubleshooting and context for complex deployments | `CG-SW-001` |
| Network drive mapping (Windows & macOS) | WI | Network | Quick how-to for both platforms | `WI-NET-001` |
| iPad migration from Jamf to Intune | WI | Mobile | Streamlined iPad enrollment process | `WI-MOBILE-001` |
| VPN connectivity troubleshooting | CG | Network | Diagnostic trees and common issues for remote workers | `CG-NET-001` |

### When to Use WI vs. CG

**Use Work Instruction (WI-)** when:
- Procedure is straightforward and well-defined
- Audience is experienced technicians or field staff familiar with the domain
- Document should be concise (3–6 pages) and serve as quick reference
- Minimal contextual explanation is needed

**Use Comprehensive Guide (CG-)** when:
- Procedure is complex with many decision points
- Audience includes new technicians, trainees, or staff learning the domain
- Document should provide deep context, background, and learning material (10–20+ pages)
- Troubleshooting scenarios, FAQs, and cross-references are valuable
- Nested step-by-step sub-procedures are necessary ("click-by-click" guidance)

### Versioning & Revision

- **First release**: Always `Rev 1.0` (in header meta and closing block)
- **Minor updates** (typos, clarifications, colour corrections): Increment 0.1 (e.g., 1.0 → 1.1)
- **Major updates** (procedure changes, new sections, restructure): Increment 1.0 (e.g., 1.0 → 2.0)
- **Effective Date**: Format as `YYYY-MM-DD` (ISO 8601)
- **Review Cycle**: Typically `Annual`; adjust if the procedure changes frequently

### Naming Conflicts & Duplicates

If you create both a WI and a CG for the same topic:
- **Same category code is acceptable** — the TYPE prefix (`WI-` vs. `CG-`) distinguishes them
  - Example: `WI-SW-001` (quick provisioning steps) and `CG-SW-001` (detailed provisioning guide)
- **Different sequential numbers are required** if you have multiple related documents in the same category
  - Example: `WI-HW-001`, `WI-HW-002`, `WI-HW-003` for different hardware procedures

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

Follow the SD42_DOCX_Style_Template.md specification for Work Instructions EXACTLY:
- Page size: Letter (12240 × 15840 DXA), margins 1080 DXA all sides
- Colours: NAVY=1F4E79, NAVY_LIGHT=D6E4F0, NAVY_MID=BDD7EE, GREY_ROW=F2F2F2
- Font: Arial throughout; body 22pt, H1 26pt, H2 22pt
- Header: 2-column borderless table — left NAVY (logo + title), right D6E4F0 (meta)
- Footer: NAVY top border, left org name, right tab page number
- H1: white on NAVY background
- H2: NAVY bold text, NAVY underline
- Callouts: left-border blocks — red (C00000/FDECEA) for WARNING, yellow (BF8F00/FFF2CC) for NOTE
- Procedure steps: numbered using step() function; bullets for non-sequential items
- Include all ISO 9001:2015 required sections for WI (Document Control, Purpose & Scope, Referenced Documents,
  Tools & Materials, Safety, Procedure, Verification, Troubleshooting, Quality Record/Sign-Off)
- End with the SD42 closing brand block

The document to create is: [DESCRIBE YOUR WI DOCUMENT HERE]
Document number: [WI-CATEGORY-NNN]
Author: [NAME]
```

### For Comprehensive Guides (CG-)

When asking an AI assistant to generate a new SD42 Comprehensive Guide, use this prompt:

```
You are generating a School District No. 42 (SD42) Comprehensive Guide (CG-) as a .docx file
using the Node.js `docx` npm library.

Follow the SD42_DOCX_Style_Template.md specification for Comprehensive Guides EXACTLY:
- Page size: Letter (12240 × 15840 DXA), margins 1080 DXA all sides
- Colours: NAVY=1F4E79, NAVY_LIGHT=D6E4F0, NAVY_MID=BDD7EE, GREY_ROW=F2F2F2
- Font: Arial throughout; body 22pt, H1 26pt, H2 22pt, H3 20pt
- Header: 2-column borderless table — left NAVY (logo + title), right D6E4F0 (meta)
- Footer: NAVY top border, left org name, right tab page number
- H1: white on NAVY background
- H2: NAVY bold text, NAVY underline
- H3: NAVY bold text (no underline) — used for sub-sections in detailed steps
- Callouts: left-border blocks — red (C00000/FDECEA) for WARNING, yellow (BF8F00/FFF2CC) for NOTE
- Procedure steps: NESTED numbering using stepDetailed() for main steps and subStep() for click-by-click
  Format: 1. [Main step], 1.a [Sub-step], 1.b [Sub-step], etc.
- Include ALL ISO 9001:2015 required sections for CG (Document Control, Purpose & Scope, Background & Context,
  Referenced Documents, Tools & Materials, Safety, Detailed Procedure with nested steps, Cross-References,
  FAQs, Expanded Troubleshooting, Verification, Quality Record/Sign-Off)
- Use faqItem() for frequently asked questions
- Use diagnosticTable() for expanded troubleshooting with symptom/cause/action
- Use crossReferences() to link related WI- and CG- documents
- End with the SD42 closing brand block

The document to create is: [DESCRIBE YOUR CG DOCUMENT HERE — note target audience is field techs needing detailed context]
Document number: [CG-CATEGORY-NNN]
Author: [NAME]
Audience: Field technicians, new staff, or anyone needing detailed step-by-step guidance
```

---

## Quick Reference — Element Summary

| Element | Function | Returns | WI | CG |
|---|---|---|---|---|
| `h1(text)` | Section heading (NAVY bg, white text) | `Paragraph` | ✓ | ✓ |
| `h2(text)` | Sub-section heading (NAVY text, underline) | `Paragraph` | ✓ | ✓ |
| `h3(text)` | Sub-sub-section heading (NAVY text) | `Paragraph` | — | ✓ |
| `body(text, opts)` | Body paragraph | `Paragraph` | ✓ | ✓ |
| `bullet(text)` | Bullet list item | `Paragraph` | ✓ | ✓ |
| `step(text)` | Numbered procedure step | `Paragraph` | ✓ | — |
| `stepDetailed(text)` | Main step (nested numbering) | `Paragraph` | — | ✓ |
| `subStep(text)` | Sub-step (nested numbering 1.a, 1.b) | `Paragraph` | — | ✓ |
| `sp(pts)` | Vertical spacer | `Paragraph` | ✓ | ✓ |
| `pb()` | Page break | `Paragraph` | ✓ | ✓ |
| `...callout(lines, "danger")` | Red WARNING block | `Paragraph[]` | ✓ | ✓ |
| `...callout(lines, "note")` | Yellow NOTE block | `Paragraph[]` | ✓ | ✓ |
| `dataTable(headers, rows, colWidths)` | Data/reference table | `Table` | ✓ | ✓ |
| `diagnosticTable(diags)` | Symptom/Cause/Action troubleshooting | `Table` | — | ✓ |
| `...faqItem(q, a)` | FAQ question & answer pair | `Paragraph[]` | — | ✓ |
| `crossReferences(docs)` | Related document references | `Table` | — | ✓ |
| `signoffTable(pairs)` | Document control / sign-off | `Table` | ✓ | ✓ |
| `verificationTable(checks)` | Pass/Fail checklist | `Table` | ✓ | ✓ |

---

## Key Differences: WI- vs. CG-

| Aspect | WI (Work Instruction) | CG (Comprehensive Guide) |
|---|---|---|
| **Audience** | Experienced technicians; quick reference | Field techs; learners needing context |
| **Length** | Concise; typically 3–6 pages | Detailed; typically 10–20+ pages |
| **Step format** | Simple numbered steps (1, 2, 3, ...) | Nested numbered steps (1, 1.a, 1.b, 1.c, ...) |
| **Sections** | Essential 9 sections | Essential 11 sections (includes Background, FAQs, Cross-Refs) |
| **Troubleshooting** | Basic Symptom/Cause/Action table | Expanded diagnostic table + decision flowcharts |
| **Context** | Minimal; assumes prior knowledge | Substantial; explains the 'why' |
| **FAQs** | None | Dedicated section for common questions |
| **Cross-refs** | Inline mentions | Dedicated Cross-References table |
| **Use case** | "How do I...?" reference | "Help me understand and do..." learning guide |

---

*SD42 IT Services — Internal use only. Template version 2.0 (Updated with CG- support)*
