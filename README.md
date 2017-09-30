# Excel Projects

This repository is nothing more than a collection of independent projects in Excel. To use or view 
any of these projects, just download the particular file you are interested in.

## Contents

* [ExternalLinkFinder](#externallinkfinder)
* [PasswordBreaker](#passwordbreaker)

### ExternalLinkFinder

If you've opened an Excel file a gotten one of these errors:

> This workbook contains links to other data sources

or

> This workbook contains one or more links that cannot be updated

then you're probably familiar with how frustrating it can sometimes be to figure out just what is 
causing it. This code searches every sheet for external links in:

- Cell formulas
- Cell conditional formatting formulas
- Cell data validation formulas
- Chart formulas
- Shape formulas
- Shape assigned macros
- Form Control input ranges
- Form Control linked cells
- Pivot Table data sources
- Regular Table data sources
- Named Ranges RefersTo formulas

When an external link is found, a new line is created in the summary sheet, detailing:
- Type     - Cell, Shape, Pivot Table, etc
- Name     - Name of Range, Pivot Table, etc
- Location - Sheet name and cell
- Offender - Was it the cell's formula? Conditional formatting? Validation?
- Value    - What the actual reference formula is

### PasswordBreaker

If a worksheet password in Excel 2010 or earlier is ever forgotten or unknown, this code will unlock
the spreadsheet. The file can be opened in a text editor, and its contents pasted into an Excel VBA 
module, or the .bas file can be imported directly into the Excel file.