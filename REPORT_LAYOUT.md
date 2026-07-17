# Reusable Word Report Layout Specification

This specification documents the branded Word-report settings used by this
project. It is intended as a reusable reference for other report generators,
regardless of implementation language.

## Page and typography

| Setting | Value |
| --- | --- |
| Page size | US Letter portrait: 8.5 x 11 in (`12240 x 15840` twips) |
| Margins | Top: 0.75 in (`1080` twips); left, right, and bottom: 0.5 in (`720` twips) |
| Header/footer distance | 0.3 in (`432` twips) |
| Default body font | Calibri, 10 pt, `#222222` |
| Cover title | Calibri bold, 28 pt, `#005F9E` |
| Cover subtitle | Calibri, 16 pt, `#203647` |
| Cover date | Calibri, 13 pt |
| Heading 1 | Calibri bold, 16 pt, `#005F9E` |
| Heading 2 | Calibri bold, 13 pt, `#404040` |
| Table body and header | Calibri, 9 pt |
| Footer | Calibri, 8.5 pt |

Use 16 pt before and 8 pt after Heading 1, and 14 pt before and 7 pt after
Heading 2. Keep headings with the following content where the document library
supports that behavior.

## Header and footer

### Header

- Place the Winslow Tech Group logo in the upper-left at 105 x 32.25 pt
  (`1333500 x 409575` EMU).
- Place the report identifier in the upper-right in Calibri 8 pt, `#404040`.
- Use a blue bottom rule in `#005F9E`.

### Footer

- Left: italic `Confidential`.
- Center: `©2026 Winslow Tech Group. All Right Reserved`.
- Right: `Page X of Y`.
- Use center and right tab stops at `5400` and `10800` twips, respectively.
- Apply a blue `#005F9E` top rule with 4 pt spacing.
- All footer runs, including page-number fields, must be 8.5 pt.

## Tables

| Setting | Value |
| --- | --- |
| Table width | Full content width: `10800` twips |
| Layout | Automatic fit (`autofit`) |
| Borders | Single line, size `1`, `#CCCCCC` |
| Header fill | `#005F9E` |
| Header text | White, bold, 9 pt |
| Alternate body rows | `#F2F7FB` |
| Cell padding | Top/bottom: `60` twips (3 pt); left/right: `120` twips (6 pt) |
| Cell alignment | Vertically centered; align status columns deliberately |
| Page breaks | Repeat header rows and prevent all rows from splitting across pages |

After generating the document, perform a final DOCX/Open XML pass for every
body table. This is required even if the document library initially assigns
column widths.

1. Determine the maximum number of grid columns, respecting any cell spans.
2. Start each column at a minimum width of `600` twips.
3. For each cell, calculate a content weight from its longest text line:

   ```text
   content_width = clamp(480 + min(longest_line_length, 60) x 90, 600, 4800)
   ```

4. Divide the cell weight across its spanned columns, retaining the greatest
   weight observed for each column.
5. Scale the widths to the usable page width (`10800` twips), updating the
   table width, grid columns, and every cell preferred width.
6. Enforce `autofit` and 6-point left/right cell padding on every table cell.

## Logo handling

- Attempt to download the current WTG PNG logo from the official WTG website.
- Validate the downloaded file is non-empty and begins with the PNG signature:
  `89 50 4E 47 0D 0A 1A 0A`.
- Use `images/winslow-technology-group-logo.png` as the bundled fallback.
- Cache a successful download only for the active report generation, then
  delete the temporary file in all success and failure paths.

## Implementation notes

- Apply the final Open XML layout pass to both native Open XML reports and
  reports first created through Word automation.
- Keep all width values in twips in OOXML. One point is 20 twips.
- Use EMUs for DrawingML image dimensions. One inch is 914400 EMUs.
- Regenerate a representative sample report and inspect its XML after changing
  layout logic. Verify the footer run size, table grid totals, cell padding,
  borders, and row page-break settings.
