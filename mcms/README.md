# MCMS Invoice Manager

## Run in dev

**Double-click in Finder:**  
Open the folder **`Motti - run this for dev mode`** (in the repo root) and double-click **`run-dev.command`**. Terminal will open and start the app with live reload—no terminal commands needed. Keep the Terminal window open while using the app.

(Any other commands—builds, tests, etc.—are run by the agent; you don’t run terminal commands yourself.)

## Template formatting (line item description)

If the **second (or later) line item description** shows truncated or duplicated text (e.g. “ite item description 2” or the description twice), the cause is usually the description cell’s formatting in the Excel template:

1. **Use a single, plain run of text**  
   In the template row that will be repeated for line items, the cell that contains `{{item_description}}` should have **no mixed formatting** (no bold/italic/colour on only part of the cell). In Excel: select that cell → ensure the whole cell is one format (e.g. all plain, or all bold), and that the only content is `{{item_description}}`.

2. **Put only the token in the description cell**  
   The cell should contain **only** `{{item_description}}` (and optional spaces). If you need a label like “Description:”, put it in a **separate cell** (e.g. in the header row), not in the same cell as the token. Same-cell prefixes/suffixes can create multiple “runs” and trigger the bug.

3. **Cell format**  
   Set the cell format to **General** or **Text** (Home → Number → General or Text). Avoid merging that cell with others if you can; merged cells can behave like multiple runs.

After changing the template, save it and regenerate the invoice. The app also forces description cells to a single text run when filling; the above keeps the template in line with that.
