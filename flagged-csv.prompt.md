You (the LLM) will receive CSV text in which every cell may include *flags* that capture visual cues (background colour, foreground colour, merged‑cell structure, and cell location) from the source spreadsheet.
Parse and reason about these flags exactly as described below.

---

#### 1  Flag grammar

* Every flag is appended to the cell's raw value and wrapped in `{ … }`.
* Multiple flags may follow the same value with **no spaces** (e.g. `100{#FFFF00}{fc:#FF0000}{MG:764455}{l:B5}`).

| Flag            | Purpose                                                | Syntax detail                                                                                                                                                                                                                                                 |
| --------------- | ------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Background colour** | Cell background colour                           | `{#XXXXXX}` or `{bc:#XXXXXX}` where `XXXXXX` is a hex RGB code.<br>The `{#XXXXXX}` format is backward-compatible.                                                           |
| **Foreground colour** | Cell text/font colour                           | `{fc:#XXXXXX}` where `XXXXXX` is a hex RGB code                                                                                                                              |
| **Merge flag**  | Identifies cells that were one merged cell in the XLSX | `{MG:YYYYYY}` where `YYYYYY` is a 6‑digit ID.<br>All cells sharing the same ID belong to the same merged block (horizontal, vertical, or both).|
| **Location flag** | Original Excel cell coordinate                       | `{l:CellRef}` where `CellRef` is the Excel cell reference (e.g., `A1`, `B5`, `AA12`).<br>Indicates the cell's original position in the spreadsheet.|

---

#### 2  Parsing rules

1. **Value extraction** – Strip all flags to obtain the cell's actual text or number.
2. **Background Color** – A cell with {#RRGGBB} or {bc:#RRGGBB} flag has background color #RRGGBB.
3. **Foreground Color** – A cell with {fc:#RRGGBB} flag has text/font color #RRGGBB.
4. **Merged cells** – Collapse every set of identical `MG:` IDs into a single virtual cell spanning the full rectangle they occupy.
5. **Location tracking** – A cell with {l:A5} flag indicates it originally came from cell A5 in the Excel spreadsheet.
6. Cells without {...} flags will be treated as normal csv cells without formats.

---

#### 3  Reasoning you can perform

* Compare colours (e.g. "Which items share the same colour?").
* Convert rows/columns covered by one merged cell into a range. The cell value in the range is not just the first cell value, but will be applied to all cells in the range.
* Reference cells by their original Excel coordinates when location flags are present (e.g. "What value is in cell B5?").

---

#### 4  Worked examples

##### Example 1: With color and merge flags

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
| My color1 | #84E291 | JUL–AUG    | 500   |
| My color2 | #E49EDD | JUL–SEP    | 600   |
| My color3 | #F6C6AC | JUL–OCT    | 700   |
| My color4 | #84E291 | JUL–DEC    | 800   |

From this table you should be able to answer, for instance:

*"Which two items share the same colour?" → **My color1** and **My color4**.*
*"What is the monthly average spend for My color4?" → 800 / 6 = 133.33*

##### Example 2: With location flags

```csv
Product{l:A1},Q1{l:B1},Q2{l:C1},Q3{l:D1},Q4{l:E1},Total{l:F1}
Laptops{l:A2}{#FFFF00},1000{l:B2},1200{l:C2},1100{l:D2},1300{l:E2},4600{l:F2}{#00FF00}
Phones{l:A3}{#FFFF00},500{l:B3},600{l:C3},700{l:D3},800{l:E3},2600{l:F3}{#00FF00}
Tablets{l:A4}{#FFFF00},300{l:B4},350{l:C4},400{l:D4},450{l:E4},1500{l:F4}{#00FF00}
```

*Interpretation*

- Yellow-highlighted product names are in column A (cells A2, A3, A4)
- Green-highlighted totals are in column F (cells F2, F3, F4)
- You can reference specific values: "The value in cell C3 is 600 (Phones Q2 sales)"
- You can answer: "What is the Q3 value for Tablets?" → Look at cell D4 → 400

##### Example 3: With foreground and background colors

```csv
Status{l:A1}{fc:#FFFFFF}{#000000},Priority{l:B1}{fc:#FFFFFF}{#000000},Task{l:C1}{fc:#FFFFFF}{#000000}
Complete{l:A2}{fc:#00FF00},High{l:B2}{fc:#FF0000}{#FFFF00},Deploy to production{l:C2}
In Progress{l:A3}{fc:#FFA500},Medium{l:B3}{fc:#FFA500},Code review{l:C3}
Pending{l:A4}{fc:#FF0000},Low{l:B4}{fc:#0000FF},Write documentation{l:C4}
```

*Interpretation*

- Header row (row 1): White text on black background
- Status column uses colored text: green for Complete, orange for In Progress, red for Pending
- Priority "High" has red text on yellow background for emphasis
- Different priority levels have different text colors
- You can identify cells with specific color combinations

Keep these rules in mind whenever the user supplies a **Flagged CSV**.