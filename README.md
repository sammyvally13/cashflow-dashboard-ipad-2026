# Cashflow Dashboard Web App

## Files
- `index.html` - UI and layout
- `styles.css` - visual design and responsive/print styles
- `app.js` - all calculation logic, charts, dynamic rows, template-based Excel export (ExcelJS)
- `Cashflow Summary template.xlsx` - Excel template used during export
- `templateData.js` - embedded template payload to preserve formatting in browser exports

## Run
Open `index.html` directly in a browser, or serve the folder with a static server.

## Notes
- Core calculation logic mirrors the original dashboard (income, CPF, expenses/liabilities, insurance, savings, summary).
- Inputs are intentionally blank by default.
- Excel export fills the attached template layout and adds a `Comments` sheet with extra details and notes.
