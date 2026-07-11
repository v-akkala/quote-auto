# SKC Service Quote Generator

Desktop tool for filling the SKC service quote Excel template from user input.

## Run

Double-click:

```bat
run_quote_generator.bat
```

The launcher creates a local `.venv`, installs `openpyxl`, and opens the app.

## Workflow

1. Fill in the quote dates and customer details.
2. Add as many products as needed. Select a product from the list to edit, duplicate, or remove it.
3. Open the `Tests` tab.
4. Add one or more tests. Test names are editable and the dropdown filters as you type.
5. Click `Generate Excel Quote`.
6. The generated workbook is saved in `outputs`.

The blank Excel template lives in `templates/Template Quote.xlsx`. Generated quote numbers start at `SQ26270001` and increment after each successful generation. The generated `outputs` and local `.venv` are ignored by git so the repo is ready to share.
