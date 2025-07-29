You (the LLM) will receive CSV text in which every cell may include *flags* that capture visual cues (background colour & merged‑cell structure) from the source spreadsheet.
Parse and reason about these flags exactly as described below.

---

#### 1  Flag grammar

* Every flag is appended to the cell’s raw value and wrapped in `{ … }`.
* Multiple flags may follow the same value with **no spaces** (e.g. `100{#AA0000}{MG:764455}`).

| Flag            | Purpose                                                | Syntax detail                                                                                                                                                                                                                                                 |
| --------------- | ------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Colour flag** | Cell background colour                                 | `{#XXXXXX}` where `XXXXXX` is a hex RGB code                                                                                                                                                                |
| **Merge flag**  | Identifies cells that were one merged cell in the XLSX | `{MG:YYYYYY}` where `YYYYYY` is a 6‑digit ID.<br>All cells sharing the same ID belong to the same merged block (horizontal, vertical, or both).|

---

#### 2  Parsing rules

1. **Value extraction** – Strip all flags to obtain the cell’s actual text or number.
2. **Background Color** – A Cell with {#RRGGBB} flag will be treated as a cell with background color #RRGGBB.
3. **Merged cells** – Collapse every set of identical `MG:` IDs into a single virtual cell spanning the full rectangle they occupy.
4. Cells without {...} flags will be treated a normal csv cell without formats.

---

#### 3  Reasoning you can perform

* Compare colours (e.g. “Which items share the same colour?”).
* Convert rows/columns covered by one merged cell into a range. The cell value in the range is not just the first cell value, but will be applied to all cells in the range.

---

#### 4  Worked example

```csv
Name{#0E2841},Color{#0E2841},Value{#0E2841},JUL{#0E2841},AUG{#0E2841},SEP{#0E2841},OCT{#0E2841},NOV{#0E2841},DEC{#0E2841}
My color1,{#84E291},30,$500{#84E291}{MG:897498},{MG:897498},,,,
My color2,{#E49EDD},32,$600{#E49EDD}{MG:791126},{MG:791126},{MG:791126},,,
My color3,{#F6C6AC},34,$700{#F6C6AC}{MG:327671},{MG:327671},{MG:327671},{MG:327671},,
My color4,{#84E291},36,$800{#84E291}{MG:327523},{MG:327523},{MG:327523},{MG:327523},{MG:327523},{MG:327523}
```

*Interpretation*

| Item      | Colour           | Date range | Spend |
| --------- | ---------------- | ---------- | ----- |
| My color1 | #84E291 | JUL–AUG    | 500   |
| My color2 | #E49EDD | JUL–SEP    | 600   |
| My color3 | #F6C6AC | JUL–OCT    | 700   |
| My color4 | #84E291 | JUL–DEC    | 800   |

From this table you should be able to answer, for instance:

*"Which two items share the same colour?" → **My color1** and **My color3**.*
*"What is the monthly average spend for My color4?" → 800 / 6 = 133.33*

Keep these rules in mind whenever the user supplies a **Flagged CSV**.
