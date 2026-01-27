# Excel Template Processing Limitations

## Overview

When the report generator processes the `auswertungs_template.xlsx` file using openpyxl, certain Excel elements are preserved while others are not. This document explains what is and isn't preserved, and why.

## ✓ What IS Preserved

The following elements from your template WILL be transferred to the generated reports:

- **Cell values** (text, numbers, formulas)
- **Cell formatting** (fonts, colors, borders, alignment)
- **Cell comments** (right-click > Insert Comment)
- **Conditional formatting rules**
- **Data validation rules**
- **Merged cells**
- **Row heights and column widths**
- **Sheet protection settings**
- **Named ranges**
- **Charts** (when properly added via openpyxl)
- **Images** (when properly added via openpyxl)
- **Hyperlinks**

## ✗ What is NOT Preserved

The following elements will be **LOST** when the template is processed:

### Text Boxes (Shapes)
- **What they are**: Text boxes created via Insert > Text Box in Excel
- **Why they're lost**: Text boxes are VML (Vector Markup Language) drawing objects stored in the legacy drawing layer
- **Technical reason**: openpyxl does not support VML shapes/drawings
- **Impact**: Any text boxes in your template will disappear in the generated reports

### Other Drawing Objects
- Shapes (rectangles, arrows, circles, etc.)
- Form controls (buttons, checkboxes, etc.)
- ActiveX controls
- WordArt

### VBA Macros
- Macros can be preserved with `keep_vba=True` but are not editable
- Currently the code does not use this parameter

## Why This Happens

When openpyxl loads and saves an Excel file:
1. It reads the Excel file structure (which is actually a ZIP file containing XML files)
2. It reconstructs the file from its internal representation
3. It only saves elements it understands and supports
4. Unsupported elements (like text boxes) are silently dropped

This is a **fundamental limitation of the openpyxl library**, not a bug in your code.

## About Cell Comments

**Cell comments SHOULD be preserved!** My testing confirms that:
- Cell comments are fully supported by openpyxl
- They persist through load/save cycles
- Writing new values to cells does NOT remove their comments

**If you're losing cell comments, please check:**

1. **Are they in the template or the source file?**
   - ✓ Comments in `auswertungs_template.xlsx` → Will be preserved
   - ✗ Comments in source ORSA files → Will NOT be copied (the code doesn't copy anything from source files)

2. **Are they actual cell comments or text boxes?**
   - Right-click on a cell and look for "Edit Comment" → Real comment ✓
   - Click and drag to move it freely → Text box ✗

## Recommended Solutions

### For Text Boxes → Use Cell Comments Instead

Instead of text boxes, use cell comments for notes/instructions:

**How to add a cell comment:**
1. Right-click on a cell
2. Select "Insert Comment" (or "New Note" in newer Excel versions)
3. Type your text
4. The comment will show when you hover over the cell

**Advantages:**
- ✓ Fully preserved by openpyxl
- ✓ Attached to specific cells
- ✓ Can be shown/hidden
- ✓ Can be printed if needed

### For Static Text → Use Cells with Special Formatting

Instead of text boxes for labels/headers:
1. Use regular cells with text
2. Apply special formatting (background color, borders, bold text)
3. Merge cells if needed for larger text areas
4. These will be fully preserved

### For Complex Instructions → Use a Separate Instructions Sheet

If you need extensive documentation:
1. Create a separate sheet called "Instructions" or "Help"
2. Use formatted cells for all text
3. This sheet will be fully preserved
4. Users can refer to it as needed

## Alternative: Using win32com (Windows Only)

If text boxes are absolutely essential, you would need to use a different approach:

```python
# This requires Windows and Excel installed
import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(template_path)
# Copy/manipulate with full Excel COM support
# This preserves ALL Excel features including text boxes
wb.SaveAs(output_path)
wb.Close()
excel.Quit()
```

**Disadvantages:**
- Only works on Windows with Excel installed
- Much slower than openpyxl
- More complex error handling
- Requires Excel license on server

## Technical Deep Dive

### Why openpyxl Can't Support Text Boxes

Excel files (.xlsx) are ZIP archives containing XML files. Text boxes are stored in:
- `xl/drawings/vmlDrawing1.vml` (legacy VML format)
- `xl/worksheets/_rels/sheet1.xml.rels` (relationships)

The VML format is a Microsoft legacy format that openpyxl doesn't parse or write. When openpyxl saves a file, it only includes XML it can generate, so VML drawings are omitted.

### How Cell Comments Are Different

Cell comments are stored in:
- `xl/comments1.xml` (comment content)
- `xl/worksheets/sheet1.xml` (comment references)

This is a standard Office Open XML format that openpyxl fully supports.

## Summary

**DO use in your template:**
- ✓ Cell comments for notes
- ✓ Formatted cells for text
- ✓ Merged cells for headers
- ✓ Conditional formatting
- ✓ Data validation

**DON'T use in your template:**
- ✗ Text boxes (Insert > Text Box)
- ✗ Shapes/drawing objects
- ✗ Form controls
- ✗ WordArt

If you're currently using text boxes, please convert them to cell comments or formatted cells, and they will be preserved in the generated reports.
