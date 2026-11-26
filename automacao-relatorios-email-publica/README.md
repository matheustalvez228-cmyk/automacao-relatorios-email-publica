# Configurable Excel-to-Outlook reporting automation

A Windows/Python starter project that reads a master workbook, filters records by age and category, creates one formatted workbook per store and category, and opens a reviewable Outlook draft for stores with eligible records.

The sample configuration uses fictional category names and empty workbook templates. It contains no customer data, contact list, company mailbox, Outlook `.oft` file, or personal paths.

## What it does

- Reads separate category sheets from a master Excel workbook.
- Filters records by an inclusive age window (the sample selects days 6 and 7).
- Groups records by store and category.
- Creates the configured store/category folder structure and uses a formatted workbook template for each output.
- Leaves a template header when a store/category has no eligible rows.
- Skips email drafts for stores with zero eligible rows.
- Attaches generated workbooks and opens Outlook drafts with `.Display()` for review.
- Does not call `.Send()`.

## Requirements

- Windows with desktop Outlook configured, for draft creation.
- Python 3.10 or newer.
- Excel input files matching the column and sheet settings in `config.json`.

Install dependencies in a virtual environment:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Configure your files

1. Copy `config.example.json` to `config.json`.
2. Put your master workbook and contacts workbook in the local `data/` folder, or change their paths in `config.json`.
3. Edit the category entries so `sheet`, `store_column`, `date_column`, template, output folder, and output name match your workbook.
4. Make each workbook template's first row match the columns in its source sheet. The sample templates are empty and fictional.
5. Configure the contact sheet's store, email, and optional manager column names.
6. Set `sender_on_behalf` only if your Outlook account has permission to use that shared mailbox. Leave it blank otherwise.

Paths in `config.json` are relative to the project folder unless they are absolute paths. Store names in the contacts workbook determine the output folders created below `stores_root`.

## Run safely

Preview eligible stores and counts without writing reports or opening Outlook:

```powershell
python automacao.py --config config.json --dry-run
```

The sample counts age inclusively: a record date is day 1, so day 6 is `today - 5 days` and day 7 is `today - 6 days`. To preview a fixed date, add `--as-of YYYY-MM-DD`.

After reviewing the counts and paths, prepare files and Outlook drafts:

```powershell
python automacao.py --config config.json --prepare
```

`--prepare` overwrites the configured output workbooks with the eligible records and displays drafts for nonzero stores. Back up any existing output files first. Review each draft manually. The code never sends messages automatically.

## Expected workbook structure

Each configured category maps to one worksheet. The source sheet must include the configured store and date columns. Optional `type_column` and `type_values` settings can further filter a source sheet when one sheet contains more than one category. The contacts workbook maps each store to its email and optional manager name.

The sample files under `examples/workbook_templates/` are header-only, contain no real records, and demonstrate the expected relationship between source columns and output formatting.

## Privacy and GitHub checklist

- Keep `config.json`, the `data/` directory, generated reports, logs, and mailbox-specific files out of Git.
- Use `config.example.json` as the public configuration template.
- Review your repository diff before pushing. Do not commit real emails, customer data, account names, or internal paths.
- Outlook automation is Windows-specific; the workbook preparation code can be adapted for other mail providers.
