/* Original public edition. Source research is documented in SOURCES.md. */
window.EXCEL_TUTORIAL = [
  {
    "id": "s01",
    "title": "Excel Introduction",
    "html": "<p>Excel turns records into calculations and reports. Start with one cell, not a dashboard. You need no programming experience.</p><h2>Your working files</h2><p>The course works offline without an account. Real workbook practice needs a separately licensed Microsoft 365 desktop Excel installation on Windows. Browser demos are limited teaching rules, not a spreadsheet application.</p><p>Keep this page and its files together. Create a separate My-work folder for your saved XLSX files. In File Explorer choose View → Show → File name extensions. CSV is plain-text data, XLSX keeps sheets/formulas/formatting, and PDF is a fixed document.</p><p>Follow Previous and Next in order. Read, change an example, check the question, then practise in Excel. Save a progress backup each week and workbooks separately. An old local course's ticks are not imported as mastery into this distinct public edition.</p>",
    "example": {
      "cells": {
        "A1": "Quantity",
        "B1": "Unit cost (RM)",
        "A2": 2,
        "B2": 350
      },
      "formula": "=A2*B2",
      "output": "C2",
      "rule": {
        "op": "product",
        "refs": [
          "A2",
          "B2"
        ]
      },
      "editable": [
        "A2",
        "B2"
      ],
      "change": {
        "cell": "A2",
        "value": 3
      },
      "explain": "Two devices at RM350 each cost RM700. A formula connects the answer to the inputs. You will learn the notation next."
    },
    "quiz": {
      "question": "Which file keeps formulas, formatting and several worksheets?",
      "options": [
        "CSV",
        "XLSX",
        "PDF"
      ],
      "correct": 1,
      "explain": "XLSX is a workbook. CSV holds plain-text records; PDF is for reading or sharing a fixed layout."
    },
    "practice": [
      "Create My-work beside this page. Confirm file-name extensions are visible.",
      "Create and save My-learning.xlsx in My-work. Close and reopen it.",
      "Name a sheet Notes and write your current skill level plus one office-reporting goal."
    ],
    "module": "00",
    "minutes": 30
  },
  {
    "id": "s02",
    "title": "Excel Cells and Ranges",
    "html": "<p>A <strong>workbook</strong> is one file. A <strong>worksheet</strong> is a tab inside it. A <strong>cell</strong> is one box. Column letters and row numbers form its address: <code>B3</code> means column B, row 3.</p><h2>Select a cell</h2><p>Click the Name Box to the left of the formula bar, type <code>A1</code> and press Enter. A range such as <code>A1:A3</code> includes A1, A2 and A3. A colon means “through”.</p><h2>Your first formula</h2><p>Enter 2 in A1 and 3 in A2. In A3, type <code>=A1+A2</code> and press Enter. The equals sign starts a formula. Change A1 to 7: A3 becomes 10 without being rewritten. Select A3 to see the formula in the formula bar.</p><p>Add a sheet with the + beside the tabs and rename it My notes by double-clicking the tab. Use Ctrl+Z to undo a mistake and Ctrl+S to save. CSV cannot keep the extra sheet: save as XLSX first.</p>",
    "example": {
      "cells": {
        "A1": 2,
        "A2": 3
      },
      "formula": "=A1+A2",
      "output": "A3",
      "rule": {
        "op": "sum",
        "refs": [
          "A1",
          "A2"
        ]
      },
      "editable": [
        "A1",
        "A2"
      ],
      "change": {
        "cell": "A1",
        "value": 7
      },
      "explain": "The result is 5. Changing A1 to 7 makes it 10. The formula stays =A1+A2."
    },
    "quiz": {
      "question": "Which range includes A1, A2 and A3?",
      "options": [
        "A1+A3",
        "A1:A3",
        "A1,A3 only"
      ],
      "correct": 1,
      "explain": "The colon includes every cell between the two addresses."
    },
    "practice": [
      "Open Data/Starter-assets.csv and save an XLSX copy before editing.",
      "Add a sheet: A1=2, A2=3, A3 =A1+A2. Change A1 to 7.",
      "Reopen the saved XLSX and check both sheets and the formula remain."
    ],
    "module": "00",
    "minutes": 45
  },
  {
    "id": "s03",
    "title": "Excel Values and Data Types",
    "html": "<p>A cell can hold a number, text, a date/time value or a formula. A product ID such as <code>001</code> is an identifier, not an amount to add. Store IDs as Text when leading zeros matter.</p><h2>Value versus appearance</h2><p>Formatting changes how a value looks. It does not normally change the value used in calculations. A cell containing 12.345 can display 12.35 and still contribute 12.345 to a total.</p><p>Select the cell and inspect the formula bar. Temporarily choose Home → Number Format → General to inspect the underlying value. Use <code>ISNUMBER(A2)</code> when you need a type check: a numeric-looking text value is not the same as a number.</p><p class=\"note\">If a date or amount shows ##### instead of a value, widen its column first. Double-click the boundary between column letters. Do not “fix” the formula just to remove a width warning.</p>",
    "example": {
      "cells": {
        "A1": "Stored amount",
        "A2": 12.345
      },
      "formula": "=A2",
      "output": "B2",
      "rule": {
        "op": "identity",
        "ref": "A2"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 12.346
      },
      "explain": "The stored result stays 12.345. With two-decimal display it looks like 12.35. Formatting has not rounded the calculation.",
      "display": "decimal2"
    },
    "quiz": {
      "question": "An amount displays 12.35 after formatting. What must you inspect before trusting a total?",
      "options": [
        "Only the font colour",
        "The stored value and formula",
        "Whether the column is centred"
      ],
      "correct": 1,
      "explain": "Displayed decimal places can hide precision. Inspect the stored value, not only its appearance."
    },
    "practice": [
      "Enter 12.345; format to two decimals and inspect the formula bar.",
      "Format a new cell as Text before entering ID 001. Compare it with a numeric 1.",
      "Document which fields are amounts and which are identifiers."
    ],
    "module": "01",
    "minutes": 60
  },
  {
    "id": "s04",
    "title": "Excel Data Entry",
    "html": "<p>A useful dataset has one header row, one field per column and one complete record per row. Tab moves right, Enter moves down and Esc cancels an uncommitted edit. Avoid merged cells and blank rows inside a dataset.</p><h2>Enter three records</h2><p>In a blank sheet enter AssetID, Department and Quantity in A1:C1. Enter A001 / IT / 2 in row 2, A002 / HR / 1 in row 3, and A003 / IT / 3 in row 4. Each quantity must stay with its original ID.</p><h2>Create your first Excel Table</h2><ol><li>Select A1:C4, including every header and record.</li><li>Press Ctrl+T. Confirm the address and tick My table has headers, then OK.</li><li>Click inside the result. The Table Design tab confirms that it is an Excel Table, not just a coloured range.</li><li>On Table Design, change Table Name to EntryPractice. Save the XLSX.</li></ol><p>Tables keep records together for sorting/filtering and can expand. You will study calculated columns later. A CSV imported through Close &amp; Load already arrives as a Table; do not create a second Table on top.</p><h2>Count what is present</h2><p>COUNTA counts populated cells, including text and empty-text formulas. COUNT counts numbers. Three names give COUNTA=3 and COUNT=0. Widen columns by double-clicking their right boundary; use Ctrl+Z on a disposable edit and verify the whole record remains aligned.</p>",
    "example": {
      "cells": {
        "A1": "Owner",
        "A2": "Alex",
        "A3": "Bo",
        "A4": "Chen"
      },
      "formula": "=COUNTA(A2:A4)",
      "output": "B2",
      "rule": {
        "op": "counta",
        "range": "A2:A4"
      },
      "editable": [
        "A2",
        "A3",
        "A4"
      ],
      "change": {
        "cell": "A4",
        "value": ""
      },
      "explain": "Three typed names give 3. Clear one name and run again: this example now has two populated cells."
    },
    "quiz": {
      "question": "What should one row normally represent?",
      "options": [
        "One complete record",
        "All values with the same colour",
        "One field spread across unrelated records"
      ],
      "correct": 0,
      "explain": "Keep related fields together in one record. This makes sorting, filtering and validation possible."
    },
    "practice": [
      "Create headers AssetID, Department, Quantity. Enter three complete records.",
      "Practise Tab, Enter, Esc and Ctrl+Z on a disposable edit.",
      "Reopen your saved workbook and verify each quantity still belongs to the correct ID."
    ],
    "module": "01",
    "minutes": 60
  },
  {
    "id": "x01",
    "title": "Excel Freeze Panes and Find",
    "minutes": 35,
    "module": "01",
    "html": "<p>Keep context visible while moving through a sheet. Freeze Panes changes the view, not values or formulas.</p><h2>Freeze the right boundary</h2><p>On Office-requests data, choose A2 → View → Freeze Panes → Freeze Panes to keep row 1 visible. To keep row 1 and column A, select B2 instead. Rows above and columns left of the selected cell stay visible. View → Freeze Panes → Unfreeze Panes reverses it.</p><p>Press Ctrl+F to find a known RequestID. Use the Name Box to return to A1. The small sum below checks that changing the view does not change arithmetic; the demo does not implement frozen panes.</p>",
    "example": {
      "cells": {
        "A1": "Request amount",
        "A2": 120,
        "A3": 80
      },
      "formula": "=SUM(A2:A3)",
      "output": "B2",
      "rule": {
        "op": "sum",
        "range": "A2:A3"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 150
      },
      "explain": "120 + 80 = 200. Freezing the real worksheet must not change that result."
    },
    "quiz": {
      "question": "Which cell freezes both row 1 and column A?",
      "options": [
        "A1",
        "B2",
        "A2 only"
      ],
      "correct": 1,
      "explain": "Freeze Panes keeps rows above and columns left of the selected cell."
    },
    "practice": [
      "Open Office-requests.csv in an XLSX copy. Freeze the first row and first column from B2.",
      "Scroll down/right and verify headers/IDs remain visible. Unfreeze and compare.",
      "Find request R006 with Ctrl+F, return to A1 and confirm amounts are unchanged."
    ]
  },
  {
    "id": "x02",
    "title": "Excel Blanks and Empty Text",
    "minutes": 45,
    "module": "01",
    "html": "<p>A cell can look empty without being truly empty. This matters for completeness checks.</p><h2>Three different cells</h2><p>Leave A2 untouched, put <code>=\"\"</code> in A3, and type one ordinary space in A4. ISBLANK is TRUE only for A2. COUNTBLANK counts A2 and A3, but COUNTA counts A3 and A4. A formula returning empty text is still a populated cell for COUNTA.</p><p>Testing <code>A3=\"\"</code> returns TRUE although ISBLANK(A3) is FALSE. Do not treat the tests as interchangeable. The demo accepts literal text only, not formulas; verify the three real cases in Excel.</p>",
    "example": {
      "cells": {
        "A1": "Owner",
        "A2": "Alex",
        "A3": "",
        "A4": " "
      },
      "formula": "=COUNTBLANK(A2:A4)",
      "output": "B2",
      "rule": {
        "op": "blankCount",
        "range": "A2:A4"
      },
      "editable": [
        "A3"
      ],
      "change": {
        "cell": "A3",
        "value": "Bo"
      },
      "explain": "One literal blank is counted. The cell containing one space is not blank."
    },
    "quiz": {
      "question": "Which statement is correct for a cell containing =\"\"?",
      "options": [
        "ISBLANK is TRUE",
        "COUNTA counts it and COUNTBLANK also counts it",
        "It is always numeric zero"
      ],
      "correct": 1,
      "explain": "Empty-text formulas are not truly empty cells. Different functions apply different counting rules."
    },
    "practice": [
      "Build the untouched / =\"\" / single-space cases in Excel.",
      "Compare ISBLANK, COUNTBLANK, COUNTA and =\"\" for each case.",
      "Write which test matches missing Owner in your report and retain raw data."
    ]
  },
  {
    "id": "s05",
    "title": "Excel Formulas and References",
    "html": "<p>A reference tells a formula where to read a value. <code>A2</code> is relative: it moves when copied. <code>$B$2</code> fixes both column B and row 2.</p><h2>Add a range</h2><p><code>=SUM(A2:A3)</code> adds A2 and A3. A function name is followed by parentheses containing its inputs. Here the input is a range.</p><h2>Keep a denominator fixed</h2><p>If B2 holds a target of 600, <code>=SUM(A2:A3)/$B$2</code> divides the total by that target. Format the result as Percentage: 0.5 means 50%, not 0.5%.</p><h2>Copy a formula</h2><p>Select a formula cell. Drag the small square at its bottom-right corner, called the fill handle, down to copy it to more rows. Alternatively, select the formula and cells below it and press Ctrl+D. Inspect the first and last copied formula. Relative references move; fixed references stay fixed.</p><p>Press F4 while editing a reference to cycle its dollar signs; some laptop keyboards need Fn+F4. A spilled formula fills several cells from one top-left formula. Its destination must be empty.</p>",
    "example": {
      "cells": {
        "A1": "Sales",
        "B1": "Target",
        "A2": 100,
        "A3": 200,
        "B2": 600
      },
      "formula": "=SUM(A2:A3)/$B$2",
      "output": "C2",
      "rule": {
        "op": "ratioSum",
        "range": "A2:A3",
        "denominator": "B2"
      },
      "editable": [
        "A2",
        "A3",
        "B2"
      ],
      "change": {
        "cell": "A2",
        "value": 200
      },
      "explain": "300 divided by 600 is 0.5, displayed as 50%. The dollars keep the target reference fixed when copied.",
      "display": "percent"
    },
    "quiz": {
      "question": "Which reference fixes both the column and row?",
      "options": [
        "B2",
        "$B$2",
        "B$2 only"
      ],
      "correct": 1,
      "explain": "$B$2 fixes both parts. B$2 fixes the row but permits the column to move."
    },
    "practice": [
      "Reproduce the example and format the ratio as Percentage.",
      "Copy a relative-reference formula down two rows and inspect first/last references.",
      "Change only the target input and explain the change without retyping results."
    ],
    "module": "01",
    "minutes": 75
  },
  {
    "id": "x03",
    "title": "Excel Named Inputs",
    "minutes": 40,
    "module": "01",
    "html": "<p>A name can make an input easier to understand. It still refers to a real cell and needs the correct scope.</p><h2>Name one assumption</h2><p>Put 1000 in A2 and 5% in B2. Select B2, type ServiceRate in the Name Box and press Enter. In C2 use =A2*ServiceRate. Use Formulas → Name Manager to inspect its Refers to address and workbook scope.</p><p>The guided example uses the equivalent fixed reference. In Excel, change the named input and inspect the formula. Use descriptive names without spaces. Delete or repair stale names only after checking their users.</p>",
    "example": {
      "cells": {
        "A1": "Base",
        "B1": "Rate",
        "A2": 1000,
        "B2": 0.05
      },
      "formula": "=A2*$B$2",
      "output": "C2",
      "rule": {
        "op": "product",
        "refs": [
          "A2",
          "B2"
        ]
      },
      "editable": [
        "B2"
      ],
      "change": {
        "cell": "B2",
        "value": 0.08
      },
      "explain": "A 5% input gives 50; 8% gives 80. A descriptive name refers to that same input."
    },
    "quiz": {
      "question": "Where can you inspect a name’s exact target?",
      "options": [
        "Name Manager",
        "Font colour",
        "Print orientation"
      ],
      "correct": 0,
      "explain": "Name Manager shows scope and the referred range."
    },
    "practice": [
      "Create workbook-scoped ServiceRate for B2 and use it in a formula.",
      "Change B2 and compare the named and fixed-reference results.",
      "Reopen the workbook and inspect the name and its target in Name Manager."
    ]
  },
  {
    "id": "s06",
    "title": "Excel Lookup Preview and Reports",
    "html": "<p>A report connects source records, calculation rules and a question someone needs answered. Keep changing assumptions in labelled cells rather than hiding them inside long formulas.</p><h2>Find a result from a band</h2><p>The example maps a proportion to a grade band. <code>XLOOKUP</code> reads a lookup value, finds a match in one range and returns the corresponding entry from another. The final <code>-1</code> allows an exact match or the next smaller threshold.</p><p>The band table starts at 0, 0.6 and 0.8. A score of 0.75 belongs to the band starting at 0.6. You will study lookup modes in detail later.</p><h2>Make the result usable</h2><p>Add a title, identify units and inspect File → Print Preview. A report must be readable before you export it. You do not need to print on paper.</p>",
    "example": {
      "cells": {
        "A1": "Minimum",
        "B1": "Band",
        "A2": 0,
        "B2": "Review",
        "A3": 0.6,
        "B3": "Pass",
        "A4": 0.8,
        "B4": "Strong",
        "D1": "Score",
        "D2": 0.75
      },
      "formula": "=XLOOKUP(D2,A2:A4,B2:B4,\"No band\",-1)",
      "output": "E2",
      "rule": {
        "op": "lookup",
        "key": "D2",
        "keys": "A2:A4",
        "values": "B2:B4",
        "mode": "smaller",
        "missing": "No band"
      },
      "editable": [
        "D2"
      ],
      "change": {
        "cell": "D2",
        "value": 0.85
      },
      "explain": "0.75 selects Pass; 0.85 selects Strong. A label changes because its input changed, not because you retyped it."
    },
    "quiz": {
      "question": "Where should a frequently changing target normally live?",
      "options": [
        "In a labelled input cell",
        "Inside every copied formula",
        "Only in the report title"
      ],
      "correct": 0,
      "explain": "One labelled input is easier to inspect and update than many hidden constants."
    },
    "practice": [
      "Build the three-band table and score input from the example.",
      "Test 0.59, 0.6, 0.79 and 0.8 and inspect the returned band.",
      "Use Print Preview and add a title and units before saving."
    ],
    "module": "01",
    "minutes": 75
  },
  {
    "id": "s07",
    "title": "Excel Sales Report Exercise",
    "html": "<p>Build a first independent sales report. Keep inputs, calculation and explanation separate. A plausible number is not evidence that the formula works.</p><h2>Sales plus one bonus</h2><p>The example totals three sales amounts and adds a bonus once. Change only the bonus: increasing it by 25 must increase the total by 25. Inspect the formula bar to distinguish a formula from a typed answer.</p><h2>Independent task</h2><p>In a new sheet enter weekly sales of 180, 240 and 330, a bonus of 60, and a target of 900. Calculate total sales including bonus and the proportion of target. Label all inputs and units. Do not read the worked check until you finish.</p>",
    "example": {
      "cells": {
        "A1": "Weekly sales",
        "B1": "Bonus",
        "A2": 100,
        "A3": 200,
        "A4": 300,
        "B2": 50
      },
      "formula": "=SUM(A2:A4)+B2",
      "output": "C2",
      "rule": {
        "op": "sumPlus",
        "range": "A2:A4",
        "add": "B2"
      },
      "editable": [
        "B2"
      ],
      "change": {
        "cell": "B2",
        "value": 75
      },
      "explain": "Weekly sales total 600; the bonus makes 650. Changing the bonus to 75 makes 675."
    },
    "quiz": {
      "question": "The weekly total is 600 and the bonus is 50. Which formula models the total?",
      "options": [
        "=SUM(A2:A4)+B2",
        "=SUM(A2:A4)-B2",
        "=B2"
      ],
      "correct": 0,
      "explain": "Add the bonus once to the regular sales total."
    },
    "practice": [
      "Build the independent 180/240/330 sales task with bonus 60 and target 900.",
      "Calculate total and percentage of target with a fixed target reference.",
      "Change bonus to 90; predict the result first, then check and restore it."
    ],
    "module": "01",
    "minutes": 90
  },
  {
    "id": "s08",
    "title": "Excel COUNTIFS and SUMIFS",
    "html": "<p>Use a count to answer “how many records?” and a sum to answer “how much?”. The result must match the business question.</p><h2>Sum matching records</h2><p><code>=SUMIFS(B2:B5,A2:A5,D2)</code> adds amounts in B only when the corresponding department in A matches D2. The sum range comes first in SUMIFS.</p><p><code>=COUNTIFS(A2:A5,D2)</code> counts matching rows instead. Every criteria range must cover corresponding records; do not shift one range down a row.</p><h2>Check the source</h2><p>For IT, find the two matching records manually before trusting the formula. Labels and criteria are usually case-insensitive in these Excel functions, but extra spaces can still make records differ.</p>",
    "example": {
      "cells": {
        "A1": "Department",
        "B1": "Amount",
        "A2": "IT",
        "B2": 1200,
        "A3": "HR",
        "B3": 800,
        "A4": "IT",
        "B4": 450,
        "A5": "Finance",
        "B5": 1100,
        "D1": "Selected",
        "D2": "IT"
      },
      "formula": "=SUMIFS(B2:B5,A2:A5,D2)",
      "output": "E2",
      "rule": {
        "op": "sumifs",
        "sum": "B2:B5",
        "criteria": [
          {
            "range": "A2:A5",
            "cell": "D2"
          }
        ]
      },
      "editable": [
        "D2"
      ],
      "change": {
        "cell": "D2",
        "value": "HR"
      },
      "explain": "IT has amounts 1200 and 450, giving 1650. Selecting HR returns 800."
    },
    "quiz": {
      "question": "You need total acquisition cost, not the number of purchases. Which operation fits?",
      "options": [
        "COUNT",
        "SUMIFS",
        "COUNTA"
      ],
      "correct": 1,
      "explain": "SUMIFS adds the selected amounts. Counting rows answers a different question."
    },
    "practice": [
      "Reproduce SUMIFS and add COUNTIFS using the same department input.",
      "Change the input from IT to HR and reconcile each result to source rows.",
      "Try * in real Excel; it is a wildcard, not a literal department. The browser demo explicitly rejects unsupported criteria."
    ],
    "module": "02",
    "minutes": 75
  },
  {
    "id": "s09",
    "title": "Excel Report Inputs and Labels",
    "html": "<p>A reusable report has one place to change its criteria. If the selected department changes but the title does not, a reader may misunderstand a correct number.</p><h2>Join text with a value</h2><p><code>=\"Department: \"&A2</code> joins a literal label with the value in A2. Quotation marks surround text; the cell address is outside them.</p><p>Combine this label with a SUMIFS report using the same input. A heading is documentation, not the underlying calculation.</p><h2>Currency formats</h2><p>Currency and Accounting change display and alignment. Use RM when the example is in Malaysian ringgit; do not multiply values merely to change a currency symbol. Converting currencies requires an explicit exchange rate and date.</p>",
    "example": {
      "cells": {
        "A1": "Selected department",
        "A2": "IT"
      },
      "formula": "=\"Department: \"&A2",
      "output": "B2",
      "rule": {
        "op": "label",
        "ref": "A2",
        "prefix": "Department: "
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": "Finance"
      },
      "explain": "The heading follows the selected department. Only the literal words and space are inside the quotation marks."
    },
    "quiz": {
      "question": "Which part is a cell reference in =\"Department: \"&A2?",
      "options": [
        "\"Department: \"",
        "A2",
        "The quotation marks"
      ],
      "correct": 1,
      "explain": "A2 reads a cell. The quoted part is literal text."
    },
    "practice": [
      "Build a department heading linked to the same input as a SUMIFS report.",
      "Change IT to Finance once and inspect heading and total.",
      "Apply RM number formatting without changing stored values."
    ],
    "module": "02",
    "minutes": 60
  },
  {
    "id": "s10",
    "title": "Excel Multiple Conditions",
    "html": "<p><strong>AND</strong> means every condition must hold for the same record. A report for IT purchases of at least RM500 must check both department and amount.</p><h2>Write criteria clearly</h2><p><code>=COUNTIFS(A2:A5,D2,B2:B5,\">=\"&E2)</code> counts rows matching D2 whose amount is at least E2. The comparison operator is text joined to the threshold cell.</p><p>At least means <code>&gt;=</code>. More than means <code>&gt;</code>. Test the boundary itself: a purchase of exactly 500 belongs in an “at least 500” report.</p><p>Keep thresholds and selected departments in labelled input cells. You can then test a new rule without rewriting all the formulas.</p>",
    "example": {
      "cells": {
        "A1": "Department",
        "B1": "Amount",
        "A2": "IT",
        "B2": 1200,
        "A3": "HR",
        "B3": 800,
        "A4": "IT",
        "B4": 450,
        "A5": "IT",
        "B5": 500,
        "D1": "Department",
        "D2": "IT",
        "E1": "Minimum",
        "E2": 500
      },
      "formula": "=COUNTIFS(A2:A5,D2,B2:B5,\">=\"&E2)",
      "output": "F2",
      "rule": {
        "op": "countifs",
        "criteria": [
          {
            "range": "A2:A5",
            "cell": "D2"
          },
          {
            "range": "B2:B5",
            "cell": "E2",
            "compare": ">="
          }
        ]
      },
      "editable": [
        "E2"
      ],
      "change": {
        "cell": "E2",
        "value": 501
      },
      "explain": "Two IT records are at least 500. Raising the threshold to 501 excludes the record exactly at 500."
    },
    "quiz": {
      "question": "Does a value of exactly 500 pass a >=500 condition?",
      "options": [
        "Yes",
        "No",
        "Only if it is formatted as currency"
      ],
      "correct": 0,
      "explain": ">= includes equality. Formatting does not change that rule."
    },
    "practice": [
      "Test minimum thresholds 499, 500 and 501 on the example.",
      "List each qualifying record manually and compare COUNTIFS.",
      "Write what at least means and show one excluded boundary record."
    ],
    "module": "02",
    "minutes": 75
  },
  {
    "id": "s11",
    "title": "Excel OR Without Double Counting",
    "html": "<p><strong>OR</strong> includes a record if either condition is true. A record meeting both conditions still belongs only once.</p><h2>Why adding counts can fail</h2><p>Adding “IT records” to “amounts at least 1000” counts an IT record of 1200 twice. In the example, three different records meet either condition.</p><p>The array formula tests both conditions per row, adds their TRUE/FALSE results, then checks whether that sum is greater than zero. The double minus converts the resulting TRUE/FALSE values to 1/0 for SUM.</p><p class=\"note\">This is a more advanced formula. First explain the four rows manually. Do not memorise the punctuation without understanding the record-level rule.</p>",
    "example": {
      "cells": {
        "A1": "Department",
        "B1": "Amount",
        "A2": "IT",
        "B2": 1200,
        "A3": "HR",
        "B3": 800,
        "A4": "IT",
        "B4": 450,
        "A5": "Finance",
        "B5": 1100
      },
      "formula": "=SUM(--(((A2:A5=\"IT\")+(B2:B5>=1000))>0))",
      "output": "D2",
      "rule": {
        "op": "countOr",
        "first": "A2:A5",
        "equals": "IT",
        "second": "B2:B5",
        "minimum": 1000
      },
      "editable": [
        "B3"
      ],
      "change": {
        "cell": "B3",
        "value": 1800
      },
      "explain": "The first, third and fourth rows qualify: 3 records. Raising HR to 1800 makes all four qualify, not five."
    },
    "quiz": {
      "question": "One IT purchase also exceeds 1000. How many times should it count in the OR report?",
      "options": [
        "Zero",
        "Once",
        "Twice"
      ],
      "correct": 1,
      "explain": "The unit being counted is a record. Meeting two conditions does not create a second record."
    },
    "practice": [
      "List records meeting IT OR amount at least 1000.",
      "Compare correct OR count with adding two separate counts.",
      "Identify the overlap record; change one input and check that it still counts once."
    ],
    "module": "02",
    "minutes": 90
  },
  {
    "id": "s12",
    "title": "Excel Conditional Report Project",
    "html": "<p>Define the population, measure and filters before writing a report. Each SUMIFS range must cover corresponding records. The small example gives IT/KL=1300 and IT/PJ=700.</p><h2>First safe CSV import</h2><p>This import creates a repeatable connection called a query. For now, concentrate on preserving the supplied fields. The later import chapters explain refresh in depth.</p><ol><li>Create a blank Excel workbook. Choose Data → Get Data → From File → From Text/CSV and select Data/Office-requests.csv.</li><li>Confirm UTF-8, comma delimiter and six columns, then choose Transform Data. The Power Query Editor opens.</li><li>Inspect Applied Steps on the right. If an automatic Changed Type step is present, select its X to remove it before setting types. This returns to raw imported text; simply changing a number back to Text later cannot restore lost leading zeros.</li><li>If needed choose Home → Use First Row as Headers. Set RequestID, Department, Status and Description to Text, Amount to Decimal Number, and RequestedDate to Date using its YYYY-MM-DD source. Use the type icon on each column header.</li><li>Check the preview: eight records, R006/FINANCE/100, and the expected headers. Choose Home → Close &amp; Load → Close &amp; Load To → Table → New worksheet. The loaded result is already an Excel Table.</li><li>Save My-request-report.xlsx in My-work. Keep the source CSV unchanged. Click the Table to see Table Design; the earlier Ctrl+T method is for a plain worksheet range, not this already-loaded Table.</li></ol><h2>Independent request report</h2><p>The loaded columns are A:F. Put selected Department in H2 and Status in I2. Use =SUMIFS(E2:E9,B2:B9,H2,C2:C9,I2) for Amount and =COUNTIFS(B2:B9,H2,C2:C9,I2) for request count. Start HR/Closed, then test an absent department. A valid zero total is not a failed import.</p>",
    "example": {
      "cells": {
        "A1": "Department",
        "B1": "Location",
        "C1": "Amount",
        "A2": "IT",
        "B2": "KL",
        "C2": 1000,
        "A3": "HR",
        "B3": "KL",
        "C3": 500,
        "A4": "IT",
        "B4": "PJ",
        "C4": 700,
        "A5": "IT",
        "B5": "KL",
        "C5": 300,
        "E1": "Department",
        "E2": "IT",
        "F1": "Location",
        "F2": "KL"
      },
      "formula": "=SUMIFS(C2:C5,A2:A5,E2,B2:B5,F2)",
      "output": "G2",
      "rule": {
        "op": "sumifs",
        "sum": "C2:C5",
        "criteria": [
          {
            "range": "A2:A5",
            "cell": "E2"
          },
          {
            "range": "B2:B5",
            "cell": "F2"
          }
        ]
      },
      "editable": [
        "F2"
      ],
      "change": {
        "cell": "F2",
        "value": "PJ"
      },
      "explain": "IT in KL totals 1300. IT in PJ totals 700. These are different subsets of the same source."
    },
    "quiz": {
      "question": "Which must stay aligned across SUMIFS ranges?",
      "options": [
        "The font sizes",
        "The corresponding record rows",
        "The column colours"
      ],
      "correct": 1,
      "explain": "Each condition and amount must refer to the same records."
    },
    "practice": [
      "Follow the safe CSV import above. Confirm eight records and the loaded Excel Table before calculating.",
      "Build selected-department count and Amount total, then add Status.",
      "Test HR/Closed and an absent department; reconcile against the raw source."
    ],
    "module": "02",
    "minutes": 120
  },
  {
    "id": "s13",
    "title": "Excel Number Formatting",
    "html": "<p>A percentage is a proportion displayed per hundred. The stored value 0.03 is 3%. Typing 3 and applying Percentage normally displays 300%.</p><h2>Separate units and values</h2><p>Label the input as a rate and enter 3%, or enter 0.03 and apply Percentage. In both cases Excel can store the same proportion.</p><p>General, Number, Currency, Accounting and Date formats serve different reading needs. They do not normally rewrite the underlying number.</p><p>To calculate a percentage amount, multiply the base by the proportion. The small example uses a fictional 3% rate solely to demonstrate arithmetic; it is not a payroll or tax rule.</p>",
    "example": {
      "cells": {
        "A1": "Base (RM)",
        "B1": "Rate",
        "A2": 1000,
        "B2": 0.03
      },
      "formula": "=A2*B2",
      "output": "C2",
      "rule": {
        "op": "product",
        "refs": [
          "A2",
          "B2"
        ]
      },
      "editable": [
        "B2"
      ],
      "change": {
        "cell": "B2",
        "value": 0.05
      },
      "explain": "1000 × 0.03 is 30. A 5% rate is entered as 0.05 in this numeric demo."
    },
    "quiz": {
      "question": "Which stored value represents 3%?",
      "options": [
        "3",
        "0.03",
        "300"
      ],
      "correct": 1,
      "explain": "3% means 3 divided by 100: 0.03."
    },
    "practice": [
      "Enter 0.03 and format Percentage. Compare with entering 3%.",
      "Multiply a base of 1000 by that rate.",
      "Change format to General and explain why the result value is unchanged."
    ],
    "module": "03",
    "minutes": 60
  },
  {
    "id": "s14",
    "title": "Excel ROUND",
    "html": "<p><code>ROUND(number, digits)</code> changes the calculated value to a stated precision. Formatting to two decimal places only changes how the value appears.</p><h2>Round at the required stage</h2><p>If a rule says each invoice line is rounded to two decimals, use ROUND on each line before summing. If the rule instead says round only the final invoice total, follow that rule. The two methods can differ.</p><p>The example multiplies a quantity by a price and rounds the result to two decimal places. Keep the original unit price separate so you can explain where the rounded amount came from.</p><p class=\"note\">These are spreadsheet mechanics, not Malaysian payroll, tax or accounting policy. Obtain the actual business rounding rule before applying a calculation operationally.</p>",
    "example": {
      "cells": {
        "A1": "Quantity",
        "B1": "Unit price",
        "A2": 3,
        "B2": 1.235
      },
      "formula": "=ROUND(A2*B2,2)",
      "output": "C2",
      "rule": {
        "op": "roundProduct",
        "refs": [
          "A2",
          "B2"
        ],
        "digits": 2
      },
      "editable": [
        "B2"
      ],
      "change": {
        "cell": "B2",
        "value": 1.237
      },
      "explain": "3 × 1.235 is 3.705; ROUND(...,2) produces 3.71. The rounded result is used by later calculations."
    },
    "quiz": {
      "question": "Which operation implements an explicit two-decimal calculation rule?",
      "options": [
        "Widen the column",
        "ROUND(value,2)",
        "Change the font"
      ],
      "correct": 1,
      "explain": "ROUND changes the calculated number. Column width and fonts change presentation."
    },
    "practice": [
      "Compare raw and ROUND(...,2) versions of 3 × 1.235.",
      "Change price to 1.237 and inspect both formula bar and displayed result.",
      "Write the point where your fictional invoice rule rounds."
    ],
    "module": "03",
    "minutes": 75
  },
  {
    "id": "s15",
    "title": "Excel Spilled Arrays",
    "html": "<p>A spilled formula returns a rectangle controlled by its top-left cell. Keep that destination clear and outside a Table calculated column.</p><h2>Two directions</h2><p>In the example each sales column uses the rate in the same column. The first corner is 100 × 10% = 10; the last is 400 × 20% = 80. Changing the second rate must not change first-column results.</p><p>Select the top-left result to edit the formula. #SPILL! often means blocked cells or merged cells. Inspect the obstruction before deleting anything; move valuable content instead.</p>",
    "example": {
      "cells": {
        "A1": "Week 1",
        "B1": "Week 2",
        "A2": 100,
        "B2": 200,
        "A3": 300,
        "B3": 400,
        "A5": "Rate 1",
        "B5": "Rate 2",
        "A6": 0.1,
        "B6": 0.2
      },
      "formula": "=ROUND(A2:B3*A6:B6,2)",
      "output": "D2",
      "rule": {
        "op": "matrixRates",
        "matrix": "A2:B3",
        "rates": "A6:B6",
        "digits": 2
      },
      "editable": [
        "B6"
      ],
      "change": {
        "cell": "B6",
        "value": 0.25
      },
      "explain": "The rectangle is 10, 40 on its first row and 30, 80 on its second. Changing week 2 affects only the second output column."
    },
    "quiz": {
      "question": "Where should you edit a spilled formula?",
      "options": [
        "In every output cell",
        "In the top-left formula cell",
        "Only in the worksheet name"
      ],
      "correct": 1,
      "explain": "The top-left formula owns the whole spilled result."
    },
    "practice": [
      "Paste the example data and enter the formula in a clear D2:E3 area.",
      "Inspect all four corners and change the second weekly rate.",
      "Put a temporary value in the destination to observe #SPILL!, then safely remove only your test obstruction."
    ],
    "module": "03",
    "minutes": 30
  },
  {
    "id": "s16",
    "title": "Excel Rounding Project",
    "html": "<p>Rounding is a business rule, not a decoration. State whether to round each line before adding or round only the final total.</p><h2>Compare two valid questions</h2><p>Two amounts of 1.234 round separately to 1.23 + 1.23 = 2.46. Rounding their unrounded sum gives 2.47. The formulas answer different rules.</p><h2>Independent invoice</h2><p>Create three lines: quantity 3 at 12.345, quantity 2 at 8.675, and quantity 5 at 1.235. Calculate each line to two decimals before summing. Compare with rounding only the final sum and write which rule the report uses. This is fictional arithmetic, not a tax policy.</p>",
    "example": {
      "cells": {
        "A1": "Unrounded line",
        "A2": 1.234,
        "A3": 1.234
      },
      "formula": "=SUM(ROUND(A2:A3,2))",
      "output": "B2",
      "rule": {
        "op": "sumRounded",
        "range": "A2:A3",
        "digits": 2
      },
      "editable": [
        "A3"
      ],
      "change": {
        "cell": "A3",
        "value": 1.236
      },
      "explain": "Two rounded lines of 1.23 total 2.46. Rounding their unrounded sum instead would produce 2.47."
    },
    "quiz": {
      "question": "Can rounding each line and rounding only the final sum produce different totals?",
      "options": [
        "Yes",
        "No",
        "Only when cells have borders"
      ],
      "correct": 0,
      "explain": "Rounding discards precision at different stages, so the totals can differ."
    },
    "practice": [
      "Build the independent three-line invoice described above.",
      "Compare sum of rounded lines with rounded final sum.",
      "Write the rule, test a changed quantity and restore it."
    ],
    "module": "03",
    "minutes": 75
  },
  {
    "id": "s17",
    "title": "Excel Dates and Working Days",
    "html": "<p>Excel normally stores dates as serial numbers. A date format turns the number into a readable calendar date. Subtracting two date values gives elapsed days.</p><h2>Avoid ambiguous input</h2><p>Use <code>DATE(year,month,day)</code>, such as DATE(2024,1,1), instead of guessing whether 1/2 means January 2 or February 1. Display dates as dd mmm yyyy when useful.</p><h2>Count working days</h2><p>NETWORKDAYS.INTL counts eligible dates, including the endpoints when they are working days. Weekend code 1 means Saturday/Sunday. A separate holiday range excludes stated holidays. WORKDAY.INTL instead finds a date a number of working days away.</p><p>The example uses a historical, fictional calendar with no holidays. It is not a Malaysian public-holiday list.</p>",
    "example": {
      "cells": {
        "A1": "January end day",
        "A2": 7
      },
      "formula": "=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,A2),1)",
      "output": "B2",
      "rule": {
        "op": "workdays",
        "year": 2024,
        "month": 1,
        "endDay": "A2"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 8
      },
      "explain": "1–7 January 2024 contains five Monday–Friday days. Including Monday 8 January makes six."
    },
    "quiz": {
      "question": "What does NETWORKDAYS.INTL return?",
      "options": [
        "A count of eligible working days",
        "The name of the last weekday",
        "A current Malaysian holiday calendar"
      ],
      "correct": 0,
      "explain": "It counts dates using your chosen weekends and holiday inputs; it does not supply a national calendar."
    },
    "practice": [
      "Create dates with DATE and compare elapsed and working-day counts.",
      "Add one weekday holiday to a holiday list and pass that range to NETWORKDAYS.INTL.",
      "Test a weekend endpoint and state whether endpoints count."
    ],
    "module": "04",
    "minutes": 75
  },
  {
    "id": "s18",
    "title": "Excel Month Deadlines",
    "html": "<p>“One month later” and “the end of next month” are different requirements. Use the function that matches the written rule.</p><h2>EDATE and EOMONTH</h2><p>EDATE moves by a number of months while preserving the day when possible. EOMONTH returns the last day of the target month. From 15 January 2024, one month later is 15 February; the next month-end is 29 February.</p><p>YEAR, MONTH and DAY extract calendar attributes from a date value. Do not extract dates by guessing character positions in a formatted label.</p><h2>Time values</h2><p>Excel times are fractions of a day. Half a day is 0.5; multiplying a duration in days by 24 gives decimal hours. Keep a duration’s calculation separate from its display.</p>",
    "example": {
      "cells": {
        "A1": "Months ahead",
        "A2": 1
      },
      "formula": "=EOMONTH(DATE(2024,1,15),A2)",
      "output": "B2",
      "rule": {
        "op": "monthEnd",
        "year": 2024,
        "month": 1,
        "offset": "A2"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 2
      },
      "explain": "The next month-end is 29 Feb 2024. Moving two months ahead gives 31 Mar 2024.",
      "display": "date"
    },
    "quiz": {
      "question": "A rule says “last day of the following month”. Which function fits?",
      "options": [
        "EOMONTH",
        "COUNTA",
        "TEXTJOIN"
      ],
      "correct": 0,
      "explain": "EOMONTH explicitly returns month-end rather than the same numbered day."
    },
    "practice": [
      "Compare EDATE and EOMONTH from 15 January 2024 with offset 1.",
      "Change offset to 2 and verify each calendar date.",
      "Enter a half-day duration and show both decimal hours and a time format."
    ],
    "module": "04",
    "minutes": 75
  },
  {
    "id": "s19",
    "title": "Excel Overnight Hours",
    "html": "<p>A time after midnight can be numerically smaller than a starting time before midnight. For a shift shorter than 24 hours, MOD can wrap the negative difference into one day.</p><h2>Calculate, then display</h2><p>The example uses typed time text, so TIMEVALUE converts it to Excel time values. MOD(end-start,1) returns a day fraction; multiplying by 24 returns hours.</p><p>22:00 to 06:00 is eight hours. In a normal worksheet, properly entered time cells already hold time values, so you can use <code>=MOD(B2-A2,1)*24</code> directly.</p><p class=\"note\">Time-only inputs cannot prove that a shift lasted more than a whole day. For multi-day work, store full start/end date-times and subtract them. Subtract stated break time before calculating paid hours.</p>",
    "example": {
      "cells": {
        "A1": "Start text",
        "B1": "End text",
        "A2": "22:00",
        "B2": "06:00"
      },
      "formula": "=MOD(TIMEVALUE(B2)-TIMEVALUE(A2),1)*24",
      "output": "C2",
      "rule": {
        "op": "overnight",
        "start": "A2",
        "end": "B2"
      },
      "editable": [
        "B2"
      ],
      "change": {
        "cell": "B2",
        "value": "07:30"
      },
      "explain": "The result is 8 decimal hours, not a clock reading. Ending at 07:30 gives 9.5 hours."
    },
    "quiz": {
      "question": "Can 22:00 and 06:00 alone distinguish an 8-hour shift from a 32-hour shift?",
      "options": [
        "Yes",
        "No",
        "Only if formatted as Currency"
      ],
      "correct": 1,
      "explain": "The dates are missing. Time-only values cannot establish whole elapsed days."
    },
    "practice": [
      "Calculate 09:00–17:00 and 22:00–06:00 using time values.",
      "Create full start/end date-times for a 32-hour period.",
      "Explain why the time-only MOD result cannot distinguish 8 from 32 hours."
    ],
    "module": "04",
    "minutes": 90
  },
  {
    "id": "s20",
    "title": "Excel Date and Time Project",
    "html": "<p>A duration needs a counting rule. From 1 through 10 January inclusive there are ten calendar dates, but the elapsed difference is nine days. Use stated historical dates, not TODAY, for a reproducible comparison.</p><h2>Import the training bookings safely</h2><ol><li>Use the From Text/CSV → Transform Data route already practised in the request report. Select Data/Staff-training.csv.</li><li>Remove any automatic Changed Type step first. From the raw text/header step set EmployeeID and BookingID to Text; do not merely convert an already-stripped 1 back to Text. Before loading, the first EmployeeID must visibly remain 001.</li><li>Set Hours to Whole Number and StartDate/DueDate to Date. Keep Department, Course and Status as Text. Close &amp; Load To → Table → New worksheet.</li><li>Inspect EmployeeID again after load. Save Training-report.xlsx, refresh the query, then save, close and reopen. Confirm the first ID is still text 001, not numeric 1 or text 1. If it is wrong, return to the raw source step and repair type conversion before continuing.</li></ol><h2>Calculate and check</h2><p>StartDate is column G and DueDate is H. Put CalendarDays in I1, =H2-G2 in I2, and fill down. Put InvalidDates in J1 and =H2&lt;G2 in J2, then fill down: TRUE means the due date precedes the start. Format duration as Number, not Date. The first two bookings are seven elapsed days; the remaining four are fourteen. Test a reversed-date record only in a disposable copy and retain the invalid flag instead of using ABS.</p>",
    "example": {
      "cells": {
        "A1": "January end day",
        "A2": 10
      },
      "formula": "=DATE(2024,1,A2)-DATE(2024,1,1)+1",
      "output": "B2",
      "rule": {
        "op": "inclusiveDays",
        "year": 2024,
        "month": 1,
        "endDay": "A2"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 11
      },
      "explain": "The inclusive count through 10 January is 10. Extending to 11 January adds one day."
    },
    "quiz": {
      "question": "What must be decided before adding 1 to a date difference?",
      "options": [
        "Whether the start day belongs in the count",
        "The worksheet tab colour",
        "Whether Excel is maximised"
      ],
      "correct": 0,
      "explain": "The business question determines inclusive versus elapsed counting."
    },
    "practice": [
      "Follow the prior CSV-import steps, deleting automatic numeric conversion before setting EmployeeID to Text. Verify 001 after preview, load, refresh and reopen.",
      "Calculate DueDate minus StartDate for each booking.",
      "Add a disposable reversed-date record and flag it rather than applying ABS."
    ],
    "module": "04",
    "minutes": 90
  },
  {
    "id": "s21",
    "title": "Excel IF",
    "html": "<p>IF chooses between two results. Its pattern is <code>IF(test, result_if_true, result_if_false)</code>.</p><h2>Turn a sentence into a test</h2><p>“Review the amount if it exceeds the limit” becomes A2&gt;B2. A value equal to the limit does not exceed it. Text outputs belong in quotation marks.</p><p>The example returns Review above RM5000 and Within limit otherwise. Test a value below, equal to and above the limit before trusting the rule.</p><p>Keep the amount and threshold separate from the formula. In a real report, also check whether missing or text input should be rejected rather than treated as a valid business value.</p>",
    "example": {
      "cells": {
        "A1": "Amount",
        "B1": "Limit",
        "A2": 6500,
        "B2": 5000
      },
      "formula": "=IF(A2>B2,\"Review\",\"Within limit\")",
      "output": "C2",
      "rule": {
        "op": "ifThreshold",
        "ref": "A2",
        "threshold": "B2",
        "compare": ">",
        "yes": "Review",
        "no": "Within limit"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 5000
      },
      "explain": "6500 exceeds 5000. Exactly 5000 returns Within limit because the test uses >, not >=."
    },
    "quiz": {
      "question": "With a >5000 rule, does exactly 5000 pass the test?",
      "options": [
        "Yes",
        "No",
        "Only if rounded"
      ],
      "correct": 1,
      "explain": "A strict greater-than comparison excludes equality."
    },
    "practice": [
      "Test amounts 4999, 5000 and 5001 against the example limit.",
      "Change the limit cell and inspect both labels.",
      "Explain the TRUE and FALSE branches in one sentence."
    ],
    "module": "05",
    "minutes": 75
  },
  {
    "id": "s22",
    "title": "Excel AND, OR and NOT",
    "html": "<p>AND is TRUE only when every test is true. OR is TRUE when at least one is true. NOT reverses TRUE and FALSE.</p><h2>Combine a rule with IF</h2><p>The fictional eligibility example requires at least 60 credits and a GPA of at least 3.5. IF turns that combined test into a readable label.</p><p>Test both boundaries, a record passing only the credit rule and a record passing only the GPA rule. An OR formula would answer a different question.</p><p class=\"note\">This is synthetic rule practice, not a real admissions decision. Real eligibility rules need authorised criteria and review. For invoice input, test missing or non-numeric values before applying financial calculations.</p>",
    "example": {
      "cells": {
        "A1": "Credits",
        "B1": "GPA",
        "A2": 60,
        "B2": 3.5
      },
      "formula": "=IF(AND(A2>=60,B2>=3.5),\"Eligible\",\"Not eligible\")",
      "output": "C2",
      "rule": {
        "op": "eligibility",
        "credits": "A2",
        "gpa": "B2"
      },
      "editable": [
        "A2",
        "B2"
      ],
      "change": {
        "cell": "A2",
        "value": 59
      },
      "explain": "Both inclusive boundaries pass. Reducing credits to 59 fails even though GPA still passes."
    },
    "quiz": {
      "question": "A record meets only one of two required conditions. Does AND return TRUE?",
      "options": [
        "Yes",
        "No",
        "It counts the record twice"
      ],
      "correct": 1,
      "explain": "AND requires every condition. OR would permit either one."
    },
    "practice": [
      "Test both pass, first-only, second-only and neither for the example AND rule.",
      "Compare OR on the same inputs.",
      "Write a truth table and identify a case where AND and OR differ."
    ],
    "module": "05",
    "minutes": 75
  },
  {
    "id": "s23",
    "title": "Excel Errors and IFNA",
    "html": "<p>Investigate an error before replacing it. IFNA handles #N/A; IFERROR handles many errors and can hide a broken reference or arithmetic.</p><h2>Missing is not zero</h2><p>An absent asset ID should say Missing ID, not invent a zero acquisition cost. Test an existing key and an absent key. Use FORMULATEXT on a formula cell to inspect the expression as text.</p><p>For #SPILL! inspect the output space; for #VALUE! inspect types; for #REF! inspect references. Compare raw and handled errors before accepting a repair.</p>",
    "example": {
      "cells": {
        "A1": "AssetID",
        "B1": "Cost",
        "A2": "A001",
        "B2": 1200,
        "A3": "A002",
        "B3": 2200,
        "A4": "A003",
        "B4": 450,
        "D1": "Find ID",
        "D2": "A099"
      },
      "formula": "=IFNA(XLOOKUP(D2,A2:A4,B2:B4),\"Missing ID\")",
      "output": "E2",
      "rule": {
        "op": "lookup",
        "key": "D2",
        "keys": "A2:A4",
        "values": "B2:B4",
        "missing": "Missing ID"
      },
      "editable": [
        "D2"
      ],
      "change": {
        "cell": "D2",
        "value": "A002"
      },
      "explain": "A099 is absent, so the result says Missing ID. A002 returns its actual cost, 2200."
    },
    "quiz": {
      "question": "Why is replacing every lookup error with 0 risky?",
      "options": [
        "It can make missing data look like a valid zero amount",
        "It always changes the source IDs",
        "It prevents saving"
      ],
      "correct": 0,
      "explain": "Missing and zero mean different things. Hiding that difference can mislead a report reader."
    },
    "practice": [
      "Test a present and absent lookup key before adding IFNA.",
      "Make a separate divide-by-zero example; compare IFNA with IFERROR.",
      "Use FORMULATEXT to inspect a formula and retain the real failure cause in Notes."
    ],
    "module": "05",
    "minutes": 75
  },
  {
    "id": "s24",
    "title": "Excel Rule Testing Project",
    "html": "<p>IFS returns the first matching result. For descending thresholds test the highest band first; 85 satisfies both ≥80 and ≥60, but should receive the higher label.</p><h2>Make a case table</h2><p>Write expected labels for 59, 60, 79, 80 and 85 before running the formula. Include exactly-on-threshold cases. One ordinary successful record is not enough.</p><h2>Independent office rule</h2><p>Create a fictional approval model: amounts above 2000 require Director review, amounts above 500 require Manager review, and the rest require Routine review. Test 500, 501, 2000 and 2001. Keep thresholds labelled. This is training, not a real employer's approval policy.</p>",
    "example": {
      "cells": {
        "A1": "Score",
        "A2": 75
      },
      "formula": "=IFS(A2>=80,\"Strong\",A2>=60,\"Pass\",TRUE,\"Review\")",
      "output": "B2",
      "rule": {
        "op": "bands",
        "ref": "A2",
        "thresholds": [
          80,
          60
        ],
        "labels": [
          "Strong",
          "Pass",
          "Review"
        ]
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 85
      },
      "explain": "75 returns Pass. 85 returns Strong because the first true test wins."
    },
    "quiz": {
      "question": "Which threshold should come first in descending bands of 80 and 60?",
      "options": [
        "60",
        "80",
        "Either always gives the same result"
      ],
      "correct": 1,
      "explain": "Testing 80 first prevents high scores being captured by the lower band."
    },
    "practice": [
      "Test the example score boundaries 59, 60, 79 and 80.",
      "Build the original office-review rule using thresholds 500 and 2000.",
      "Test exactly 500, 501, 2000 and 2001 and explain rule order."
    ],
    "module": "05",
    "minutes": 90
  },
  {
    "id": "s25",
    "title": "Excel XLOOKUP",
    "html": "<p>XLOOKUP retrieves information using a key. A product ID is usually a safer key than a description that may be duplicated or edited.</p><h2>Exact matching</h2><p><code>XLOOKUP(D2,A2:A4,B2:B4,\"Missing ID\")</code> searches the IDs in A and returns the corresponding price from B. Exact matching is the default.</p><p>The lookup and return ranges must have matching dimensions. A key stored as text may not match a number that merely looks the same. Check types and leading zeros before changing match modes.</p><h2>Add a dropdown</h2><p>Select D2. Choose Data → Data Validation. Under Settings, choose Allow: List and set Source to <code>=$A$2:$A$4</code>. Confirm, then use the arrow in D2 to choose a key. This controls input; it does not prove source IDs are unique or the returned record is correct.</p>",
    "example": {
      "cells": {
        "A1": "Product ID",
        "B1": "Price",
        "A2": "AP",
        "B2": 350,
        "A3": "SW",
        "B3": 1200,
        "A4": "LT",
        "B4": 2200,
        "D1": "Selected ID",
        "D2": "AP"
      },
      "formula": "=XLOOKUP(D2,A2:A4,B2:B4,\"Missing ID\")",
      "output": "E2",
      "rule": {
        "op": "lookup",
        "key": "D2",
        "keys": "A2:A4",
        "values": "B2:B4",
        "missing": "Missing ID"
      },
      "editable": [
        "D2"
      ],
      "change": {
        "cell": "D2",
        "value": "LT"
      },
      "explain": "AP returns 350 and LT returns 2200. An absent key returns a visible Missing ID message."
    },
    "quiz": {
      "question": "Which ranges must have corresponding dimensions?",
      "options": [
        "The lookup and return ranges",
        "All coloured cells",
        "The worksheet titles"
      ],
      "correct": 0,
      "explain": "XLOOKUP aligns the found position with the same position in the return range."
    },
    "practice": [
      "Reproduce the three-product exact lookup.",
      "Add Data Validation → List for the selected ID; test a valid choice.",
      "Test a missing ID and inspect whether source IDs are unique."
    ],
    "module": "06",
    "minutes": 75
  },
  {
    "id": "s26",
    "title": "Excel Approximate Lookups",
    "html": "<p>Use approximate matching only when the business rule describes bands or capacity limits. It is not a repair for misspelled IDs.</p><h2>Choose a direction</h2><p>XLOOKUP match mode -1 allows exact or next smaller. Mode 1 allows exact or next larger. A commission threshold usually uses the lower band; a minimum adequate capacity usually needs the next larger value.</p><p>The example uses lower bounds 0, 500 and 1000. An amount of 750 belongs to the band beginning at 500.</p><p>Keep threshold tables ordered and check boundary, between-band and out-of-range values. LOOKUP and binary-search modes have additional sorting requirements. Do not choose a search algorithm merely because it sounds faster.</p>",
    "example": {
      "cells": {
        "A1": "Lower bound",
        "B1": "Band",
        "A2": 0,
        "B2": "Low",
        "A3": 500,
        "B3": "Medium",
        "A4": 1000,
        "B4": "High",
        "D1": "Amount",
        "D2": 750
      },
      "formula": "=XLOOKUP(D2,A2:A4,B2:B4,\"No band\",-1)",
      "output": "E2",
      "rule": {
        "op": "lookup",
        "key": "D2",
        "keys": "A2:A4",
        "values": "B2:B4",
        "mode": "smaller",
        "missing": "No band"
      },
      "editable": [
        "D2"
      ],
      "change": {
        "cell": "D2",
        "value": 1000
      },
      "explain": "750 selects Medium; exactly 1000 selects High. The match direction comes from the lower-bound rule."
    },
    "quiz": {
      "question": "A capacity must be at least as large as the requirement. Which approximate direction fits?",
      "options": [
        "Next smaller",
        "Next larger",
        "Any misspelling match"
      ],
      "correct": 1,
      "explain": "The next larger listed capacity can satisfy the minimum. A smaller capacity cannot."
    },
    "practice": [
      "Test 499, 500, 501 and 1000 against lower-bound bands.",
      "Test a value below the first bound.",
      "Explain why minimum capacity requires next-larger rather than next-smaller."
    ],
    "module": "06",
    "minutes": 75
  },
  {
    "id": "s27",
    "title": "Excel Duplicate and Two-Way Lookups",
    "html": "<p>Repeated keys need a rule. An exact lookup normally returns the first matching position; search mode -1 searches last-to-first.</p><h2>Last is not latest</h2><p>The last AP row in the demo returns Chen, but row position alone does not prove the newest date. To claim latest, inspect dates and ordering. Sort complete records, never one isolated column.</p><h2>Two directions</h2><p>For a two-way rate table, first identify its row and column. Start with a small manually checked case before combining two lookups. Wildcards match patterns; they do not repair spelling or duplicate IDs.</p>",
    "example": {
      "cells": {
        "A1": "Device",
        "B1": "Owner",
        "A2": "AP",
        "B2": "Alex",
        "A3": "Laptop",
        "B3": "Bo",
        "A4": "AP",
        "B4": "Chen",
        "D1": "Device",
        "D2": "AP"
      },
      "formula": "=XLOOKUP(D2,A2:A4,B2:B4,\"No match\",0,-1)",
      "output": "E2",
      "rule": {
        "op": "lookup",
        "key": "D2",
        "keys": "A2:A4",
        "values": "B2:B4",
        "mode": "last",
        "missing": "No match"
      },
      "editable": [
        "D2"
      ],
      "change": {
        "cell": "D2",
        "value": "Laptop"
      },
      "explain": "AP returns Chen, the last matching row. Laptop has one match and returns Bo."
    },
    "quiz": {
      "question": "Does last-to-first lookup alone prove that a record is the newest?",
      "options": [
        "Yes",
        "No",
        "Only when the text is bold"
      ],
      "correct": 1,
      "explain": "It proves source position, not chronological recency. Inspect dates and ordering."
    },
    "practice": [
      "Compare first and last search on repeated device labels.",
      "Sort whole records differently and observe the result.",
      "Create a small two-way department/month table and trace one selected intersection manually."
    ],
    "module": "06",
    "minutes": 75
  },
  {
    "id": "s28",
    "title": "Excel FILTER and List Comparison",
    "html": "<p>XLOOKUP normally returns one matching record. FILTER returns all records whose include test is TRUE.</p><h2>Extract complete rows</h2><p>The example filters A2:C4 using the department column. The returned name, department and amount stay together. Keep the output range clear.</p><p>To join a filtered name list into one cell, use TEXTJOIN with a delimiter. For this example, <code>=TEXTJOIN(\", \",TRUE,F2:F3)</code> joins the two filtered names after the IT result is present.</p><h2>Compare lists</h2><p>XMATCH returns a match position or #N/A. ISNUMBER(XMATCH(...)) creates matched flags; ISNA(XMATCH(...)) identifies missing items. Check one shared and one absent record manually.</p>",
    "example": {
      "cells": {
        "A1": "Name",
        "B1": "Department",
        "C1": "Amount",
        "A2": "Alex",
        "B2": "IT",
        "C2": 350,
        "A3": "Bo",
        "B3": "HR",
        "C3": 800,
        "A4": "Chen",
        "B4": "IT",
        "C4": 450,
        "E1": "Department",
        "E2": "IT"
      },
      "formula": "=FILTER(A2:C4,B2:B4=E2,\"No matches\")",
      "output": "F2",
      "rule": {
        "op": "filter",
        "data": "A2:C4",
        "criteria": "B2:B4",
        "key": "E2",
        "missing": "No matches"
      },
      "editable": [
        "E2"
      ],
      "change": {
        "cell": "E2",
        "value": "HR"
      },
      "explain": "IT returns two complete records. Selecting HR returns Bo’s record only; it does not leave stale IT rows."
    },
    "quiz": {
      "question": "Which function is suitable when every matching row is needed?",
      "options": [
        "FILTER",
        "A first-match XLOOKUP alone",
        "ROUND"
      ],
      "correct": 0,
      "explain": "FILTER can return multiple records. A single matching position is insufficient."
    },
    "practice": [
      "Filter complete example records for IT, HR and an absent department.",
      "Clear enough space for the spill and inspect all fields stay aligned.",
      "Use TEXTJOIN on the returned names and explain what a no-match result means."
    ],
    "module": "06",
    "minutes": 90
  },
  {
    "id": "s29",
    "title": "Excel Lookup Invoice Project",
    "html": "<p>Trace a calculated quantity through its units. Units × kilograms per unit produces kilograms. Dividing produces a different measure.</p><h2>Invoice check</h2><p>Two units weighing 1.2 kg each weigh 2.4 kg. Increase the order to three: weight becomes 3.6 kg. A correct product lookup should not alter the quantity.</p><h2>Independent catalogue</h2><p>Create three products with unique text IDs, unit prices and unit weights. Build an order sheet with a validated product input and quantity. Return price and weight by exact lookup, calculate both totals, and visibly flag an absent ID. A missing price is not free stock.</p>",
    "example": {
      "cells": {
        "A1": "Units",
        "B1": "Product",
        "A2": 2,
        "B2": "AP",
        "D1": "Product",
        "E1": "Unit kg",
        "D2": "AP",
        "E2": 1.2,
        "D3": "Switch",
        "E3": 2.5,
        "D4": "Laptop",
        "E4": 1.8
      },
      "formula": "=A2*XLOOKUP(B2,D2:D4,E2:E4)",
      "output": "C2",
      "rule": {
        "op": "lookupProduct",
        "quantity": "A2",
        "key": "B2",
        "keys": "D2:D4",
        "values": "E2:E4"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 3
      },
      "explain": "2 × 1.2 kg = 2.4 kg. Three units weigh 3.6 kg. The product lookup supplies the unit weight."
    },
    "quiz": {
      "question": "How do you calculate total weight from units and unit weight?",
      "options": [
        "Units ÷ unit weight",
        "Units × unit weight",
        "Add the product descriptions"
      ],
      "correct": 1,
      "explain": "The units cancel correctly when quantity is multiplied by weight per unit."
    },
    "practice": [
      "Create your own three-product price/weight catalogue and order sheet.",
      "Use exact lookups, quantity multiplication and a visible missing-ID message.",
      "Change quantity, then product; predict and verify both monetary and weight totals."
    ],
    "module": "06",
    "minutes": 120
  },
  {
    "id": "s30",
    "title": "Excel Mixed References",
    "html": "<p>A mixed reference fixes only one part of an address. <code>$A2</code> fixes column A; <code>B$1</code> fixes row 1.</p><h2>Build a two-way model</h2><p>In <code>=$A2*B$1</code>, copying right keeps the amount in column A but changes the rate column. Copying down moves to the next amount while keeping the rate row.</p><p>For A2=100, B1=0.1 and C1=0.2, B2 is 10 and a copy in C2 becomes 20. With A3=200, a copy in B3 becomes 20.</p><h2>Highlight complete rows</h2><p>A conditional-formatting formula such as <code>=$C2=$F$2</code> fixes the tested column and the selected input while permitting the row to move. Apply it to the complete intended row range.</p>",
    "example": {
      "cells": {
        "A1": "Amount",
        "A2": 100,
        "A3": 200,
        "B1": 0.1,
        "C1": 0.2
      },
      "formula": "=$A2*B$1",
      "output": "B2",
      "rule": {
        "op": "product",
        "refs": [
          "A2",
          "B1"
        ]
      },
      "editable": [
        "B1"
      ],
      "change": {
        "cell": "B1",
        "value": 0.15
      },
      "explain": "The first amount times the first rate gives 10. Change the rate to 0.15 and that result becomes 15."
    },
    "quiz": {
      "question": "Which reference fixes column A but allows the row to move?",
      "options": [
        "A$2",
        "$A2",
        "$A$2"
      ],
      "correct": 1,
      "explain": "The dollar sign before A fixes the column, not the row number."
    },
    "practice": [
      "Fill =$A2*B$1 across and down on the example.",
      "Inspect all four corner formulas.",
      "Apply a row-highlighting formula to a full test range with the right mixed references."
    ],
    "module": "07",
    "minutes": 75
  },
  {
    "id": "s31",
    "title": "Excel Order of Operations",
    "html": "<p>Excel evaluates multiplication and division before addition and subtraction. Parentheses change the grouping.</p><h2>Explain before editing</h2><p>With 2, 3 and 4, <code>=A2+B2*C2</code> returns 14. <code>=(A2+B2)*C2</code> returns 20. These formulas implement different rules.</p><p>Use Formulas → Evaluate Formula to step through a confusing expression. Check inputs, types and references before wrapping it in IFERROR.</p><h2>Text-number mismatches</h2><p>An extracted ID may be text while the lookup table uses numbers. Convert deliberately with VALUE or a suitable numeric coercion when the identifier really is numeric. Do not strip meaningful zeros from an ID that should remain text.</p>",
    "example": {
      "cells": {
        "A1": "First",
        "B1": "Second",
        "C1": "Multiplier",
        "A2": 2,
        "B2": 3,
        "C2": 4
      },
      "formula": "=(A2+B2)*C2",
      "output": "D2",
      "rule": {
        "op": "groupProduct",
        "first": "A2",
        "second": "B2",
        "multiplier": "C2"
      },
      "editable": [
        "C2"
      ],
      "change": {
        "cell": "C2",
        "value": 5
      },
      "explain": "Parentheses make 2+3 happen first, then multiply by 4: 20. Without the grouping, the result would be 14."
    },
    "quiz": {
      "question": "What does =(2+3)*4 return?",
      "options": [
        "14",
        "20",
        "24"
      ],
      "correct": 1,
      "explain": "The parentheses make the addition happen before multiplication."
    },
    "practice": [
      "Compare =A2+B2*C2 and =(A2+B2)*C2.",
      "Use Formulas → Evaluate Formula to inspect the grouping.",
      "Test a text versus numeric lookup key and document the correct ID type."
    ],
    "module": "07",
    "minutes": 75
  },
  {
    "id": "s32",
    "title": "Excel Tables",
    "html": "<p>An Excel Table is a named, structured range. It can expand with new records and maintain calculated columns.</p><h2>Create a Table</h2><p>Select the complete source range, press Ctrl+T and confirm My table has headers. Use Table Design to give it a meaningful name. Do not include unrelated notes inside it.</p><p>In a calculated TotalValue column, <code>=[@Quantity]*[@UnitCost]</code> means “this row’s quantity times this row’s unit cost”. The @ refers to the current row.</p><p>Add a new record below the Table and check that the Table and formula expand. A fixed range may miss the new row. Dynamic spills belong outside Table calculated columns; use a normal row formula inside the Table.</p>",
    "example": {
      "cells": {
        "A1": "Quantity",
        "B1": "UnitCost",
        "C1": "TotalValue",
        "A2": 2,
        "B2": 350,
        "A3": 3,
        "B3": 450
      },
      "formula": "=[@Quantity]*[@UnitCost]",
      "output": "C2",
      "rule": {
        "op": "product",
        "refs": [
          "A2",
          "B2"
        ]
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 3
      },
      "explain": "The first row calculates 2 × 350 = 700. The guided demo shows this one row; use Ctrl+T in Excel to test actual Table expansion.",
      "table": "A1:C3"
    },
    "quiz": {
      "question": "What does @Quantity mean in a Table formula?",
      "options": [
        "The whole workbook’s quantity",
        "The current row’s Quantity value",
        "A worksheet password"
      ],
      "correct": 1,
      "explain": "@ selects the current row within a structured Table reference."
    },
    "practice": [
      "Create the real Table and calculated TotalValue column.",
      "Append another complete record and inspect formula expansion.",
      "Name the Table and make sure a report source uses it, not a fixed old range."
    ],
    "module": "07",
    "minutes": 90
  },
  {
    "id": "x04",
    "title": "Excel Filtered Totals",
    "minutes": 45,
    "module": "07",
    "html": "<p>A filtered report can show a grand total or a visible-row total. Label which one you mean.</p><h2>Use a Table total</h2><p>Turn on Table Design → Total Row and choose Sum in the Amount column. Excel normally uses SUBTOTAL so filtered-out records do not contribute. SUBTOTAL(109,range) also ignores manually hidden rows; SUM normally does not.</p><p>The browser example has all rows visible and no filtering engine. Test actual hidden/filtered behaviour in Excel. Clear filters before comparing a company-wide total.</p>",
    "example": {
      "cells": {
        "A1": "Department",
        "B1": "Amount",
        "A2": "IT",
        "B2": 100,
        "A3": "HR",
        "B3": 50,
        "A4": "IT",
        "B4": 150
      },
      "formula": "=SUBTOTAL(109,B2:B4)",
      "output": "D2",
      "rule": {
        "op": "sum",
        "range": "B2:B4"
      },
      "editable": [
        "B3"
      ],
      "change": {
        "cell": "B3",
        "value": 80
      },
      "explain": "All-visible sum is 300. The real Excel result changes when HR is filtered out."
    },
    "quiz": {
      "question": "With HR filtered out, which visible-row total should remain?",
      "options": [
        "300",
        "250",
        "50"
      ],
      "correct": 1,
      "explain": "Only the IT amounts of 100 and 150 remain visible."
    },
    "practice": [
      "Create a real Table and compare SUM with SUBTOTAL(109,...).",
      "Filter Department to IT and verify visible subtotal 250 while ordinary SUM remains 300.",
      "Clear the filter, manually hide HR and compare functions 9 and 109. Unhide afterwards."
    ]
  },
  {
    "id": "s33",
    "title": "Excel Text Extraction and Debugging",
    "html": "<p>A description may contain a usable key plus extra text. Extract the key first, verify it, then perform the lookup.</p><h2>Split at a delimiter</h2><p>TEXTBEFORE(A2,\":\") returns the text before the first colon. In this example the ID is numeric, so the double minus converts that extracted text to a number before XLOOKUP.</p><p>If the colon is missing, investigate the input rather than assuming the whole description is a valid key. If an ID has meaningful leading zeros, keep the key as text instead of using numeric coercion.</p><h2>Choose the right repair</h2><p>#NAME? can mean an unsupported function or a misspelling. #N/A often means no match. #VALUE! points to an incompatible value or operation. Inspect the smallest failing part before changing the complete formula.</p>",
    "example": {
      "cells": {
        "A1": "Description",
        "A2": "7204: Office-laptop",
        "D1": "ID",
        "E1": "Price",
        "D2": 7102,
        "E2": 31.5,
        "D3": 7204,
        "E3": 47.25,
        "D4": 7306,
        "E4": 29.75
      },
      "formula": "=XLOOKUP(--TEXTBEFORE(A2,\":\"),D2:D4,E2:E4,\"Unknown\")",
      "output": "B2",
      "rule": {
        "op": "extractLookup",
        "ref": "A2",
        "delimiter": ":",
        "keys": "D2:D4",
        "values": "E2:E4",
        "missing": "Unknown"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": "7306: Office-monitor"
      },
      "explain": "7204 is extracted and matched to 47.25. Changing the ID to 7306 returns 29.75."
    },
    "quiz": {
      "question": "What should you verify before debugging the returned price?",
      "options": [
        "The extracted key and its type",
        "The worksheet colour",
        "The number of PDF pages"
      ],
      "correct": 0,
      "explain": "A correct lookup cannot repair an incorrectly extracted or mistyped key."
    },
    "practice": [
      "Run TEXTBEFORE alone before adding the numeric lookup.",
      "Test a valid description, missing delimiter and absent key.",
      "Explain when numeric conversion would wrongly remove meaningful zeros."
    ],
    "module": "07",
    "minutes": 75
  },
  {
    "id": "x05",
    "title": "Excel Text Cleaning and Delimiters",
    "minutes": 60,
    "module": "07",
    "html": "<p>Clean only according to a known field rule. Keep the raw input beside the cleaned result.</p><h2>Extract a document code</h2><p>LEFT(A2,3) takes three leading characters; RIGHT takes the end; MID takes a start position and length. LEN counts characters. For variable-length fields use a known delimiter instead of guessed positions.</p><p>TEXTBEFORE and TEXTAFTER can split an agreed separator. Test missing and repeated separators. TRIM handles ordinary spaces, but non-breaking spaces need deliberate replacement, for example SUBSTITUTE(A2,CHAR(160),\" \") before TRIM. Do not remove meaningful internal spaces from names or IDs.</p>",
    "example": {
      "cells": {
        "A1": "Document",
        "A2": "FIN-2026-014"
      },
      "formula": "=LEFT(A2,3)",
      "output": "B2",
      "rule": {
        "op": "textPart",
        "ref": "A2",
        "count": 3
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": "HRM-2026-015"
      },
      "explain": "The department prefix is FIN. The changed document returns HRM. This fixed-width rule assumes exactly three prefix characters."
    },
    "quiz": {
      "question": "What should happen before overwriting raw imported text?",
      "options": [
        "Keep a raw copy and define the cleaning rule",
        "Remove every space everywhere",
        "Convert every ID to a number"
      ],
      "correct": 0,
      "explain": "Preserve evidence and the field’s meaning before transforming it."
    },
    "practice": [
      "Extract prefix, year and sequence from FIN-2026-014 using LEFT/MID/RIGHT; check LEN.",
      "Repeat using TEXTBEFORE/TEXTAFTER and test an input with no hyphen.",
      "Compare ordinary versus non-breaking spaces in Excel; document your chosen cleanup."
    ]
  },
  {
    "id": "s34",
    "title": "Excel Reconciliation Exercise",
    "html": "<p>Reconciliation asks which keys agree and which explain a difference. Equal row counts do not prove equal records.</p><h2>Return missing keys</h2><p>XMATCH returns a position or #N/A. ISNA flags absent keys; FILTER returns the flagged items. Check one shared and one absent key manually.</p><h2>Independent lists</h2><p>Create an Expected list of A001, A002, A003 and A004 and a Counted list of A001, A003, A003 and A005. Identify missing IDs, unexpected IDs and duplicates separately. A set comparison does not measure duplicate quantities.</p>",
    "example": {
      "cells": {
        "A1": "Expected",
        "C1": "Counted",
        "A2": "AP",
        "A3": "Switch",
        "A4": "Laptop",
        "C2": "AP",
        "C3": "Laptop"
      },
      "formula": "=FILTER(A2:A4,ISNA(XMATCH(A2:A4,C2:C3)),\"None\")",
      "output": "E2",
      "rule": {
        "op": "missing",
        "expected": "A2:A4",
        "actual": "C2:C3",
        "none": "None"
      },
      "editable": [
        "C3"
      ],
      "change": {
        "cell": "C3",
        "value": "Switch"
      },
      "explain": "Switch is absent from the counted list. Replacing Laptop with Switch makes Laptop the missing item instead."
    },
    "quiz": {
      "question": "Do equal row counts prove that two lists contain the same items?",
      "options": [
        "Yes",
        "No",
        "Only if the cells are green"
      ],
      "correct": 1,
      "explain": "Different records or duplicates can produce equal counts. Compare the keys."
    },
    "practice": [
      "Build the Expected and Counted ID lists specified above.",
      "Return missing IDs; separately list unexpected and duplicate IDs.",
      "Change Counted A005 to A004 and verify the exact changed exceptions."
    ],
    "module": "07",
    "minutes": 90
  },
  {
    "id": "s35",
    "title": "Excel Text Numbers and Results",
    "html": "<p>A displayed 10 can be text or a number. VALUE deliberately converts suitable numeric text; Number formatting alone may not convert existing content.</p><h2>Score comparison</h2><p>The example converts both text scores before comparing, so 10 is greater than 7. Keep identifiers with meaningful leading zeros as text instead of converting them.</p><h2>Inspect before converting</h2><p>Try ISNUMBER and ISTEXT on sample cells. Preserve the raw input column. For a suitable single numeric column, Data → Text to Columns → Finish can convert values, but inspect several results and do not run it blindly on IDs.</p>",
    "example": {
      "cells": {
        "A1": "Score text",
        "B1": "Other score text",
        "A2": "10",
        "B2": "7"
      },
      "formula": "=IF(VALUE(A2)>VALUE(B2),\"Win\",\"Not win\")",
      "output": "C2",
      "rule": {
        "op": "compareNumericText",
        "first": "A2",
        "second": "B2"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": "5"
      },
      "explain": "Numeric conversion makes 10 > 7 true. A score of 5 is not greater than 7."
    },
    "quiz": {
      "question": "Does applying Number format always convert existing text numbers?",
      "options": [
        "Yes",
        "No",
        "Only after changing the font"
      ],
      "correct": 1,
      "explain": "Formatting and conversion are different operations. Verify the stored type after a deliberate conversion."
    },
    "practice": [
      "Store 10 and 7 as text and use ISNUMBER to inspect them.",
      "Compare after deliberate VALUE conversion, then try invalid text.",
      "Keep original raw input and document why IDs must not be converted blindly."
    ],
    "module": "07",
    "minutes": 75
  },
  {
    "id": "s36",
    "title": "Excel UNIQUE and Dropdown Lists",
    "html": "<p>UNIQUE builds a distinct list and SORT orders it. Put the spill outside a Table in clear space. Duplicate source rows remain in the source unless you explicitly remove them.</p><h2>Connect a dropdown</h2><p>Enter =SORT(UNIQUE(A2:A5)) in C2. Select E2 → Data → Data Validation → List and set Source to =$C$2#. The # references the whole spill. Test a new source value and inspect the list.</p><p>For this example, clearing A5 leaves Ali, Bo and a final numeric 0 in the spill, not an empty dropdown option. Inspect the source before using the list.</p><p>Validation controls entry. Conditional Formatting changes appearance. Neither proves that your imported source records are complete or correct.</p>",
    "example": {
      "cells": {
        "A1": "Name",
        "A2": "Ali",
        "A3": "Bo",
        "A4": "Ali",
        "A5": "Chen"
      },
      "formula": "=SORT(UNIQUE(A2:A5))",
      "output": "C2",
      "rule": {
        "op": "uniqueSort",
        "range": "A2:A5"
      },
      "editable": [
        "A5"
      ],
      "change": {
        "cell": "A5",
        "value": "Bo"
      },
      "explain": "Ali occurs twice but appears once in the sorted result. Replacing Chen with Bo leaves only Ali and Bo."
    },
    "quiz": {
      "question": "Which operation removes duplicates for the dropdown source?",
      "options": [
        "Data Validation by itself",
        "UNIQUE in a helper formula",
        "Conditional Formatting"
      ],
      "correct": 1,
      "explain": "UNIQUE creates the distinct list. Validation then uses that list; formatting controls highlighting."
    },
    "practice": [
      "Build a UNIQUE/SORT spill and connect a validation dropdown to it.",
      "Add a new source name within the formula range and inspect its list entry.",
      "Test duplicate source names and explain why validation does not clean the source."
    ],
    "module": "07",
    "minutes": 75
  },
  {
    "id": "s37",
    "title": "Excel Formula Structures",
    "html": "<p>Choose a formula structure that makes the rule easy to inspect and maintain. A copied formula, dynamic spill and Table calculated column solve different layout problems.</p><h2>Check an intermediate result</h2><p>The example extracts numeric product ID 7204 from a description. Confirm that value before combining it with a price lookup.</p><p>A copied formula needs correctly moving references. A spill needs a clear destination and one controlling formula. A Table calculated column uses a current-row reference and can extend with new records.</p><p>When comparing structures, keep the inputs and rounding rule identical. Agreement between three incorrectly specified formulas is still not proof that they answer the business question.</p>",
    "example": {
      "cells": {
        "A1": "Description",
        "A2": "7204: Office-laptop"
      },
      "formula": "=--TEXTBEFORE(A2,\":\")",
      "output": "B2",
      "rule": {
        "op": "extractNumber",
        "ref": "A2",
        "delimiter": ":"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": "7306: Office-monitor"
      },
      "explain": "The returned key is the number 7204, not the complete description. A changed prefix changes the key."
    },
    "quiz": {
      "question": "Where should you place a dynamic spilling formula?",
      "options": [
        "Inside a Table calculated column",
        "Outside the Table with clear output space",
        "In every cell of the spill"
      ],
      "correct": 1,
      "explain": "A spill needs its own clear destination. Table calculated columns use row-by-row formulas instead."
    },
    "practice": [
      "Extract a description key and verify its type before adding a lookup.",
      "Compare one copied formula, one spill and one Table calculated column on the same inputs.",
      "Change a source value and explain which formula controls each output."
    ],
    "module": "07",
    "minutes": 75
  },
  {
    "id": "s38",
    "title": "Excel Payroll Model Exercise",
    "html": "<p>A model needs distinct units: Hours, HourlyRate, Sales and CommissionRate. Keep changing assumptions in labelled cells.</p><h2>Fictional gross amount</h2><p>Eight hours at 25 plus 3% of 1000 is 230. Two more hours should add 50 without changing the commission. Round only at the stage required by the stated rule.</p><p>This example teaches arithmetic. It does not implement Malaysian payroll, statutory deductions, tax, leave or overtime law. Do not use training rates to make real pay decisions.</p>",
    "example": {
      "cells": {
        "A1": "Hours",
        "B1": "Hourly rate",
        "C1": "Sales",
        "D1": "Commission rate",
        "A2": 8,
        "B2": 25,
        "C2": 1000,
        "D2": 0.03
      },
      "formula": "=ROUND(A2*B2+C2*D2,2)",
      "output": "E2",
      "rule": {
        "op": "gross",
        "hours": "A2",
        "rate": "B2",
        "sales": "C2",
        "commission": "D2"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 10
      },
      "explain": "Hourly earnings of 200 plus commission of 30 give 230. Two more hours make the total 280."
    },
    "quiz": {
      "question": "Which input is a proportion rather than an amount?",
      "options": [
        "Hours",
        "Commission rate",
        "Gross earnings"
      ],
      "correct": 1,
      "explain": "A rate such as 3% is a proportion used to calculate an amount."
    },
    "practice": [
      "Build the fictional gross model and test eight versus ten hours.",
      "Change commission rate from 3% to 5% without rewriting formulas.",
      "Label it synthetic; explain why this is not a compliant payroll calculation."
    ],
    "module": "07",
    "minutes": 90
  },
  {
    "id": "s39",
    "title": "Excel PivotTables",
    "html": "<p>A PivotTable groups records and calculates a chosen measure. Count, Sum and Average are not interchangeable.</p><h2>Create the real object</h2><ol><li>Copy the example data to Excel. Select its whole range, press Ctrl+T and confirm headers.</li><li>Choose Insert → PivotTable → From Table/Range → New Worksheet.</li><li>Drag Department to Rows and Amount to Values.</li><li>Open Value Field Settings and select Sum. Set Number Format there.</li></ol><p>The browser example shows a SUMIFS benchmark, not a browser PivotTable. In Excel the IT subtotal must equal 1450. Text amounts can cause Count to appear by default; fix types before accepting a monetary report.</p>",
    "example": {
      "cells": {
        "A1": "Department",
        "B1": "Amount",
        "A2": "IT",
        "B2": 1000,
        "A3": "HR",
        "B3": 800,
        "A4": "IT",
        "B4": 450,
        "D1": "Department",
        "D2": "IT"
      },
      "formula": "=SUMIFS(B2:B4,A2:A4,D2)",
      "output": "E2",
      "rule": {
        "op": "sumifs",
        "sum": "B2:B4",
        "criteria": [
          {
            "range": "A2:A4",
            "cell": "D2"
          }
        ]
      },
      "editable": [
        "D2"
      ],
      "change": {
        "cell": "D2",
        "value": "HR"
      },
      "explain": "The IT subtotal is 1450. A real PivotTable using Sum should reconcile to this same subset."
    },
    "quiz": {
      "question": "Your source Amount column contains text numbers and Values defaults to Count. What should you do?",
      "options": [
        "Accept Count as money",
        "Inspect types and choose the intended measure",
        "Rename Count to Sales"
      ],
      "correct": 1,
      "explain": "A label does not fix a wrong measure. Inspect source types and value-field settings."
    },
    "practice": [
      "Create the actual Table-backed PivotTable from the small example.",
      "Choose Sum explicitly and reconcile IT=1450 with source records.",
      "Save and reopen; inspect the PivotTable fields, not just a screenshot."
    ],
    "module": "08",
    "minutes": 75
  },
  {
    "id": "s40",
    "title": "Excel Refresh, Charts and Slicers",
    "html": "<p>A source Table expands, but a PivotTable cache still needs refresh. Add a source row, then right-click the PivotTable → Refresh.</p><h2>Make a readable chart</h2><p>Click the PivotTable → PivotTable Analyze → PivotChart. Use a column chart for category comparisons and a line for ordered time. Label measure and units; do not call a count revenue.</p><h2>Filter deliberately</h2><p>Use PivotTable Analyze → Insert Slicer, choose a field, then select an item. Clear its filter before reconciling grand totals. The example's closed-share denominator is all four populated responses, not only visible closed rows.</p>",
    "example": {
      "cells": {
        "A1": "Status",
        "A2": "Closed",
        "A3": "Open",
        "A4": "Closed",
        "A5": "Closed",
        "C1": "Status",
        "C2": "Closed"
      },
      "formula": "=COUNTIF(A2:A5,C2)/COUNTA(A2:A5)",
      "output": "D2",
      "rule": {
        "op": "share",
        "range": "A2:A5",
        "key": "C2"
      },
      "editable": [
        "A3"
      ],
      "change": {
        "cell": "A3",
        "value": "Closed"
      },
      "explain": "Three of four tickets are Closed: 75%. Closing the fourth makes it 100%.",
      "display": "percent"
    },
    "quiz": {
      "question": "What must you do to a PivotTable after adding a source record?",
      "options": [
        "Only change its title",
        "Refresh it",
        "Delete all formulas"
      ],
      "correct": 1,
      "explain": "The Pivot cache needs a refresh to include new source data."
    },
    "practice": [
      "Add one complete source row and refresh the PivotTable.",
      "Create a measure-labelled PivotChart and a department slicer.",
      "Filter, then clear; verify the unfiltered totals and chart units."
    ],
    "module": "08",
    "minutes": 90
  },
  {
    "id": "s41",
    "title": "Excel Reporting Page Project",
    "html": "<p>A one-page report needs a question, traceable source, explicit measures and readable output. A matching grand total alone does not prove every subtotal is right.</p><h2>Average the records</h2><p>AVERAGEIFS averages matching records. Do not average department averages to make a company average unless the weights are correct. A small department otherwise has the same influence as a large one.</p><h2>Independent office report</h2><p>Use Data/Office-requests.csv. Make department Sum and Average of Amount, plus Count of RequestID. Add a column chart of total Amount labelled RM. Clear filters, check totals against source rows and use File → Print Preview. Export PDF only after confirming it is readable.</p>",
    "example": {
      "cells": {
        "A1": "Month",
        "B1": "Revenue",
        "A2": "Jan",
        "B2": 100,
        "A3": "Jan",
        "B3": 300,
        "A4": "Feb",
        "B4": 800,
        "A5": "Feb",
        "B5": 1200,
        "D1": "Month",
        "D2": "Jan"
      },
      "formula": "=AVERAGEIFS(B2:B5,A2:A5,D2)",
      "output": "E2",
      "rule": {
        "op": "averageifs",
        "sum": "B2:B5",
        "criteria": [
          {
            "range": "A2:A5",
            "cell": "D2"
          }
        ]
      },
      "editable": [
        "D2"
      ],
      "change": {
        "cell": "D2",
        "value": "Feb"
      },
      "explain": "January’s record average is (100+300)/2 = 200. February’s is 1000."
    },
    "quiz": {
      "question": "How should the company-wide average Amount be calculated?",
      "options": [
        "Average department averages without weights",
        "Average all underlying request amounts",
        "Add the department names"
      ],
      "correct": 1,
      "explain": "Use the underlying records, or apply correct record-count weights. Unweighted group averages can mislead."
    },
    "practice": [
      "Build department Sum, Average and Count from Office-requests.csv.",
      "Add a readable RM total chart; reconcile each department manually.",
      "Set print area, orientation and scaling; inspect PDF at normal reading size."
    ],
    "module": "08",
    "minutes": 90
  },
  {
    "id": "x06",
    "title": "Excel Weighted Averages",
    "minutes": 50,
    "module": "08",
    "html": "<p>Different-sized groups need weights. Averaging two group averages gives each group equal influence regardless of its size.</p><h2>Use the denominator</h2><p>A group of two bookings averaging 10 hours and a group of eight averaging 20 hours have combined average (10×2 + 20×8)/(2+8) = 18 hours, not 15.</p><p>Use SUMPRODUCT(means,counts)/SUM(counts) only when each mean and count describes the same measure/population. Prefer averaging underlying records when available. A zero total count is not a zero average; investigate it.</p>",
    "example": {
      "cells": {
        "A1": "Group mean",
        "B1": "Records",
        "A2": 10,
        "B2": 2,
        "A3": 20,
        "B3": 8
      },
      "formula": "=SUMPRODUCT(A2:A3,B2:B3)/SUM(B2:B3)",
      "output": "D2",
      "rule": {
        "op": "weighted",
        "values": "A2:A3",
        "weights": "B2:B3"
      },
      "editable": [
        "B2"
      ],
      "change": {
        "cell": "B2",
        "value": 8
      },
      "explain": "The combined mean is 18. Equal group sizes of eight make it 15."
    },
    "quiz": {
      "question": "What is wrong with averaging department averages without weights?",
      "options": [
        "It may give small and large groups equal influence",
        "Excel cannot add averages",
        "It always returns zero"
      ],
      "correct": 0,
      "explain": "The group record counts determine their contribution to a combined average."
    },
    "practice": [
      "Reproduce weighted and unweighted averages and explain the difference.",
      "Create the ten underlying records: two 10s and eight 20s. Compare AVERAGE with the weighted calculation.",
      "Test a zero denominator and keep that failure visible."
    ]
  },
  {
    "id": "s42",
    "title": "Excel Sort, Filter and Flash Fill",
    "html": "<p>Sort and filter complete records. Sorting a single field without the related columns can attach the wrong amounts or owners to an ID.</p><h2>Use the whole source</h2><p>Click inside an Excel Table before sorting. Outside a Table, select the whole data range and confirm the header row. Filtering hides records; it does not delete them.</p><h2>Flash Fill is a one-time result</h2><p>Type a representative pattern beside a source text column, then choose Data → Flash Fill or Ctrl+E. Inspect every generated value. Unlike a formula, Flash Fill does not automatically update when the source text changes.</p><p>The count below is a small frequency report. Counts should reconcile to the source responses after filters are cleared.</p>",
    "example": {
      "cells": {
        "A1": "Phone",
        "A2": "Apple",
        "A3": "Android",
        "A4": "Apple",
        "A5": "Android",
        "C1": "Phone",
        "C2": "Apple"
      },
      "formula": "=COUNTIF(A2:A5,C2)",
      "output": "D2",
      "rule": {
        "op": "countifs",
        "criteria": [
          {
            "range": "A2:A5",
            "cell": "C2"
          }
        ]
      },
      "editable": [
        "A3"
      ],
      "change": {
        "cell": "A3",
        "value": "Apple"
      },
      "explain": "Apple occurs twice. Changing one Android response to Apple increases the count to three."
    },
    "quiz": {
      "question": "Does Flash Fill automatically recalculate after its source text changes?",
      "options": [
        "Yes",
        "No",
        "Only in a blue Table"
      ],
      "correct": 1,
      "explain": "Flash Fill produces static values. A formula is needed for a live transformation."
    },
    "practice": [
      "Sort Office-requests.csv as complete records and verify an ID/Amount pair.",
      "Use Flash Fill on a disposable text-pattern column and inspect every result.",
      "Change source text to prove Flash Fill is static, unlike a formula."
    ],
    "module": "09",
    "minutes": 75
  },
  {
    "id": "s43",
    "title": "Excel Import CSV with Power Query",
    "html": "<p>Power Query records import steps for later refresh. Start with the supplied original Data/Office-requests.csv, not employer data.</p><h2>Import safely</h2><ol><li>Create a blank workbook. Data → Get Data → From File → From Text/CSV. Choose Office-requests.csv.</li><li>Confirm UTF-8, comma delimiter and six expected headers in the preview. Choose Transform Data.</li><li>Use First Row as Headers only if needed. Set RequestID, Department, Status and Description to Text, Amount to Decimal Number, RequestedDate to Date using its explicit ISO date convention.</li><li>Close &amp; Load To → Table → New Worksheet. Save Requests-report.xlsx in My-work.</li></ol><p>Inspect the eight source records. The browser count example is only an ID-type illustration. Importing and refreshing real data must be done in Excel.</p>",
    "example": {
      "cells": {
        "A1": "ID (text)",
        "A2": "001",
        "A3": "002",
        "A4": "003"
      },
      "formula": "=COUNTA(A2:A4)",
      "output": "B2",
      "rule": {
        "op": "counta",
        "range": "A2:A4"
      },
      "editable": [
        "A4"
      ],
      "change": {
        "cell": "A4",
        "value": ""
      },
      "explain": "Three non-empty IDs give three records. Keeping the IDs as Text preserves their leading zeros."
    },
    "quiz": {
      "question": "Which type preserves meaningful leading zeros in an ID?",
      "options": [
        "Text",
        "Decimal number",
        "Currency"
      ],
      "correct": 0,
      "explain": "IDs are keys, not amounts. Text retains zeros that numeric conversion could remove."
    },
    "practice": [
      "Import Office-requests.csv through From Text/CSV, not just File → Open.",
      "Check six headers, eight records, ID text, dates and numeric Amount.",
      "Save Requests-report.xlsx with its query for the next lesson."
    ],
    "module": "09",
    "minutes": 75
  },
  {
    "id": "s44",
    "title": "Excel Clean and Refresh a Query",
    "html": "<p>Recurring cleanup belongs in query steps, not manual edits to its loaded Table. A refresh can overwrite the latter.</p><h2>Clean without guessing</h2><ol><li>Data → Queries &amp; Connections → right-click the request query → Edit.</li><li>Select Department → Transform → Format → Trim. Then Uppercase only because these supplied department codes have an agreed uppercase convention.</li><li>Keep the original Description; do not split arbitrary human names or force their capitalisation as an identity repair.</li><li>Close &amp; Load. Check the record count and Amount total stayed unchanged.</li></ol><p>Excel TRIM collapses ordinary internal spaces as well as ends. Power Query Text.Trim normally trims the ends, not all internal runs. Non-breaking spaces need deliberate handling. Do not assume the two operations are identical.</p><h2>Test refresh in a disposable copy</h2><p>Keep the supplied CSV and your original report unchanged. Copying a workbook does not change the source path stored in its query.</p><ol><li>Save and close My-work/Requests-report.xlsx.</li><li>Create My-work/Requests-copy. Copy Requests-report.xlsx and the supplied Data/Office-requests.csv into that folder.</li><li>Open only My-work/Requests-copy/Requests-report.xlsx. Choose Data → Get Data → Data Source Settings → Data sources in current workbook. Select the original Office-requests.csv source → Change Source → Browse to My-work/Requests-copy/Office-requests.csv. Confirm the copied path. Do not clear permissions or change global privacy settings.</li><li>Refresh the copied query and wait for its Table to finish loading. Confirm eight records, Amount total RM1170 and R006 Department FINANCE.</li><li>Open only the copied Office-requests.csv in a plain-text editor. In record R006 change the Department field from FINANCE to HR. Preserve the other fields, six-column comma-delimited structure and UTF-8 encoding, then save the CSV.</li><li>Refresh the copied query again and wait for loaded rows. Refresh any PivotTable that reads this Table afterwards.</li><li>Inspect the loaded R006 row: Department must be HR. Reconcile eight records and RM1170 total, with HR RM500 and FINANCE RM150. An unchanged grand total alone does not prove that the new source was read.</li><li>Save and close the copied workbook. Check that the supplied original CSV still has R006 FINANCE. Reopen the original Requests-report.xlsx: its own source binding and records must remain unchanged, with HR RM400 and FINANCE RM250 after normalisation.</li></ol>",
    "example": {
      "cells": {
        "A1": "Raw name",
        "A2": "  Alex   Tan  "
      },
      "formula": "=TRIM(A2)",
      "output": "B2",
      "rule": {
        "op": "trim",
        "ref": "A2"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": "  Bo  Lee "
      },
      "explain": "The ordinary spaces become Alex Tan. Cleanup should preserve the meaningful words, not guess a person’s identity."
    },
    "quiz": {
      "question": "Where should recurring cleanup normally be recorded?",
      "options": [
        "Only as manual edits in the loaded Table",
        "In query transformation steps",
        "Only in a chart title"
      ],
      "correct": 1,
      "explain": "Recorded transformations can be reapplied to a later export. Manual loaded-result edits are not a reliable refresh process."
    },
    "practice": [
      "Edit the existing request query: trim Department, then use its agreed uppercase codes. Save the original report.",
      "Follow the disposable-copy sequence above. Copy both workbook and CSV, then explicitly bind only the copied workbook to the copied CSV before editing it.",
      "Change R006 to HR only in the copied CSV, refresh and inspect the loaded row and department totals. Verify eight records and RM1170, then check the original CSV and report remain unchanged."
    ],
    "module": "09",
    "minutes": 90
  },
  {
    "id": "s45",
    "title": "Excel Combine Files from a Folder",
    "html": "<p>Combine only files with the same schema and one record meaning. Keep unrelated CSVs and Add-later outside the source folder.</p><h2>Monthly sales workflow</h2><ol><li>Copy Data/Monthly-sales to My-work/Sales-source. Create My-sales.xlsx. Data → Get Data → From File → From Folder → select Sales-source → Combine &amp; Transform.</li><li>Confirm comma delimiter, headers and the matching Product, Department, Units, UnitPrice and Date columns. Set text, whole-number, decimal and date types explicitly.</li><li>Add Column → Custom Column: Revenue = [Units]*[UnitPrice]. Close &amp; Load To → Table.</li><li>Create a PivotTable from that Table with Department in Rows and Sum of Revenue in Values.</li><li>Copy the supplied March CSV from Data/Add-later into Sales-source once. Refresh the query and wait until its loaded row count updates; only then refresh the PivotTable.</li></ol><p>Record initial and refreshed row counts, units and revenue. Loading is not proof that every file belongs in the report.</p>",
    "example": {
      "cells": {
        "A1": "Monthly sales",
        "A2": 100,
        "A3": 200,
        "A4": 300
      },
      "formula": "=SUM(A2:A4)",
      "output": "B2",
      "rule": {
        "op": "sum",
        "range": "A2:A4"
      },
      "editable": [
        "A4"
      ],
      "change": {
        "cell": "A4",
        "value": 350
      },
      "explain": "The toy total is 600. Updating the third amount to 350 makes 650; the real folder report needs its own query and Pivot refresh."
    },
    "quiz": {
      "question": "Which folder should be the initial source for the prepared monthly task?",
      "options": [
        "The entire Data folder",
        "Data/Monthly-sales only",
        "Data/Add-later only"
      ],
      "correct": 1,
      "explain": "Only the dedicated monthly source contains the intended compatible initial files."
    },
    "practice": [
      "Build the real monthly folder query, loaded Table and PivotTable.",
      "Check initial record count, total units and revenue before adding March once.",
      "Refresh query first, then Pivot; reconcile all three new totals."
    ],
    "module": "09",
    "minutes": 90
  },
  {
    "id": "s46",
    "title": "Excel Join Descriptions to Records",
    "html": "<p>A lookup adds an attribute to a transaction; it must not create extra transactions. Check key uniqueness and types before enriching a report.</p><h2>Keep the grain</h2><p>The example maps LT to Laptop. An unmatched ID remains visibly unmatched, not assigned a convenient description. If a lookup table repeats a key, decide the business rule rather than hiding ambiguity with a first match.</p><p>In a two-table query merge, duplicate right-side keys can multiply records when expanded. Compare row count and total units before and after. Stop and investigate a changed measure before building charts.</p>",
    "example": {
      "cells": {
        "A1": "Product ID",
        "B1": "Description",
        "A2": "AP",
        "B2": "Access point",
        "A3": "SW",
        "B3": "Switch",
        "A4": "LT",
        "B4": "Laptop",
        "D1": "Transaction ID",
        "D2": "LT"
      },
      "formula": "=XLOOKUP(D2,A2:A4,B2:B4,\"Unmatched\")",
      "output": "E2",
      "rule": {
        "op": "lookup",
        "key": "D2",
        "keys": "A2:A4",
        "values": "B2:B4",
        "missing": "Unmatched"
      },
      "editable": [
        "D2"
      ],
      "change": {
        "cell": "D2",
        "value": "AP"
      },
      "explain": "LT is labelled Laptop; AP becomes Access point. Descriptions come from the matching lookup record."
    },
    "quiz": {
      "question": "What should happen to total units when you add correct descriptions by lookup?",
      "options": [
        "They should double",
        "They should stay unchanged",
        "They should become text"
      ],
      "correct": 1,
      "explain": "Adding an attribute should not create new transactions or alter their measure."
    },
    "practice": [
      "Create a unique product-description lookup and add descriptions to monthly transactions.",
      "Test an absent key visibly.",
      "Reconcile source rows and units before and after; deliberately test a duplicate lookup key in a disposable copy."
    ],
    "module": "09",
    "minutes": 75
  },
  {
    "id": "s47",
    "title": "Excel Frequency and Trend Project",
    "html": "<p>A frequency report counts categories and divides by an explicit population. Three iOS responses out of four populated responses give 75%.</p><h2>Independent request frequency</h2><p>Use the eight Office-requests records. Count Status values and show each as a percentage of all requests. Clear filters and reconcile the counts to eight. Rounding can make displayed percentages differ slightly from 100%.</p><h2>Trend axes</h2><p>For a monthly sales trend, use the actual months on the horizontal axis. Point numbers 1, 2 and 3 are not enough context. Keep the count/percentage measure distinct from sales revenue.</p>",
    "example": {
      "cells": {
        "A1": "Response",
        "A2": "iOS",
        "A3": "Android",
        "A4": "iOS",
        "A5": "iOS",
        "C1": "Category",
        "C2": "iOS"
      },
      "formula": "=COUNTIF(A2:A5,C2)/COUNTA(A2:A5)",
      "output": "D2",
      "rule": {
        "op": "share",
        "range": "A2:A5",
        "key": "C2"
      },
      "editable": [
        "A3"
      ],
      "change": {
        "cell": "A3",
        "value": "iOS"
      },
      "explain": "The category has three of four responses: 75%. Moving the remaining response into that category makes 100%.",
      "display": "percent"
    },
    "quiz": {
      "question": "What should a yearly trend’s horizontal axis show?",
      "options": [
        "Unrelated worksheet row numbers",
        "The actual years",
        "Only the chart title"
      ],
      "correct": 1,
      "explain": "The axis must identify the time periods represented by the points."
    },
    "practice": [
      "Count Office-requests status categories and calculate shares of all requests.",
      "Clear filters and reconcile to eight records and 100% before rounding.",
      "Make a monthly trend with actual month labels rather than point numbers."
    ],
    "module": "09",
    "minutes": 75
  },
  {
    "id": "s48",
    "title": "Excel Local Data Analysis Project",
    "html": "<p>Apply the reporting method to unfamiliar records. Data/Staff-training.csv is synthetic: one row is an employee/course booking, not a unique employee.</p><h2>HR/admin report</h2><p>Import IDs as Text and dates as Date. Create department counts of bookings, a total Hours measure and completed bookings using Status. A person on two courses appears twice, so booking count is not headcount.</p><h2>Explain the grain</h2><p>Use UNIQUE on EmployeeID to count distinct employees separately. Do not present training hours as paid work or a legal entitlement. Record what a missing date means instead of inventing it.</p>",
    "example": {
      "cells": {
        "A1": "Department",
        "B1": "Hourly amount",
        "A2": "IT",
        "B2": 20,
        "A3": "IT",
        "B3": 30,
        "A4": "HR",
        "B4": 20,
        "A5": "HR",
        "B5": 40,
        "D1": "Department",
        "D2": "IT"
      },
      "formula": "=AVERAGEIFS(B2:B5,A2:A5,D2)",
      "output": "E2",
      "rule": {
        "op": "averageifs",
        "sum": "B2:B5",
        "criteria": [
          {
            "range": "A2:A5",
            "cell": "D2"
          }
        ]
      },
      "editable": [
        "D2"
      ],
      "change": {
        "cell": "D2",
        "value": "HR"
      },
      "explain": "IT averages 25 and HR averages 30 in this synthetic sample. Counts matter when combining department results."
    },
    "quiz": {
      "question": "An employee appears on two course bookings. What does counting those rows measure?",
      "options": [
        "Distinct employees",
        "Bookings",
        "Paid overtime"
      ],
      "correct": 1,
      "explain": "The grain is one employee/course booking. Count distinct EmployeeID separately for people."
    },
    "practice": [
      "Import Staff-training.csv and define one row as one booking.",
      "Report bookings, Hours, completed bookings and distinct EmployeeID separately.",
      "Explain why six bookings are not six distinct employees and document a missing-date rule."
    ],
    "module": "09",
    "minutes": 120
  },
  {
    "id": "s49",
    "title": "Excel Report Documentation",
    "html": "<p>A handover names source files, one row's meaning, field types, transformations, assumptions, refresh order and acceptance checks. A copied workbook can silently keep reading its original absolute source path.</p><h2>Controlled recovery lab</h2><ol><li>Save and close the working My-sales.xlsx. Leave its Sales-source folder unchanged. Copy the workbook to My-work/Recovery-lab/My-sales-copy.xlsx and copy all current sales CSVs into Recovery-lab/Source.</li><li>Open only the copied workbook. Rebind its sales query to Recovery-lab/Source: Data → Get Data → Data Source Settings → Data sources in current workbook → select the old sales folder → Change Source → browse the copied Source. Do not clear permissions or change global privacy settings. If editing Source steps instead, check the main query and any generated Sample File helper that still names the old folder.</li><li>Prove the copy is independent. In Notepad create probe.csv in Recovery-lab/Source with the header and single record shown below. Use Save as type All Files and UTF-8 so it is not probe.csv.txt. Refresh the copied query, wait for loaded rows, then refresh its Pivot. Expect nine records, fourteen units and RM12623; the original source/report must remain at eight/thirteen/RM12500.</li><li>Save and close the copied workbook. Move only its disposable Recovery-lab/Source folder to Recovery-lab/Moved-source. Do not move the original Sales-source.</li><li>Reopen the copied workbook and refresh. Observe the missing-source failure. Do not accept cached old numbers as a successful refresh.</li><li>Change the copied workbook's source from Recovery-lab/Source to Recovery-lab/Moved-source, including any generated helper source still using the old lab path. Refresh query first, then Pivot; verify nine/fourteen/RM12623 again.</li><li>Save, close and reopen both reports separately. Confirm the original still reads its original source and its data is unchanged. Record the recovery steps in your handover.</li></ol><pre>Date,Product,Department,Units,UnitPrice\n2026-04-01,AP,IT,1,123</pre><p>If your copied report never shows the probe row, it is probably still reading the original source. Stop and fix the binding before trying the move. Never fabricate a completed refresh or lower security settings to hide a failure.</p>",
    "example": {
      "cells": {
        "A1": "Units sold",
        "A2": 2,
        "A3": 3,
        "A4": 4
      },
      "formula": "=SUM(A2:A4)",
      "output": "B2",
      "rule": {
        "op": "sum",
        "range": "A2:A4"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 1
      },
      "explain": "Both views of these same transactions should reconcile to nine units. Reducing the first row by one makes eight."
    },
    "quiz": {
      "question": "What should a report handover include?",
      "options": [
        "Only a screenshot",
        "Inputs, steps, refresh order, checks and limitations",
        "Only the author’s favourite formula"
      ],
      "correct": 1,
      "explain": "A handover must let another person trace and repeat the workflow."
    },
    "practice": [
      "Write a one-page handover for the monthly report with source, grain, types, steps and refresh order.",
      "Rebind a disposable workbook copy to its own copied source, prove the probe row changes only that report, then move and repair that disposable source using the steps above.",
      "Refresh again and verify totals; retain your original working report unchanged."
    ],
    "module": "09",
    "minutes": 75
  },
  {
    "id": "s50",
    "title": "Excel IT Asset Report",
    "html": "<p>Build a bridge from IT to office reporting using original synthetic asset records. One row can represent several devices, so records and quantities answer different questions.</p><h2>Build your own workbook</h2><ol><li>Open Data/Starter-assets.csv and immediately save My-asset-report.xlsx in My-work.</li><li>Create a Table named Assets. Add TotalValue = [@Quantity]*[@UnitCost]. This is acquisition value, not current resale value.</li><li>Create a department PivotTable with Count of AssetID, Sum of Quantity and Sum of TotalValue. Make a clearly labelled acquisition-value column chart.</li><li>Filter for missing Owner and record the affected IDs. Keep IDs as Text and dates unambiguous.</li></ol><p>Check every department against its underlying records. The browser example is smaller than the supplied dataset and cannot grade your workbook.</p>",
    "example": {
      "cells": {
        "A1": "Quantity",
        "B1": "UnitCost",
        "A2": 2,
        "B2": 350,
        "A3": 1,
        "B3": 1200,
        "A4": 3,
        "B4": 100
      },
      "formula": "=SUMPRODUCT(A2:A4,B2:B4)",
      "output": "C2",
      "rule": {
        "op": "sumproduct",
        "left": "A2:A4",
        "right": "B2:B4"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 3
      },
      "explain": "This toy dataset has six devices and value 2200. It is not the full 12-record starter dataset."
    },
    "quiz": {
      "question": "Why can 12 asset records represent 15 devices?",
      "options": [
        "Some records have Quantity greater than one",
        "Excel has miscounted every row",
        "A record always means exactly one device"
      ],
      "correct": 0,
      "explain": "The dataset’s grain is an asset group. Sum Quantity to count devices."
    },
    "practice": [
      "Build the full original asset dataset report, not the three-row demo.",
      "Check 12 records, 15 devices and RM15000 acquisition value before changes.",
      "Explain one department subtotal and a record-count versus quantity difference."
    ],
    "module": "10",
    "minutes": 90
  },
  {
    "id": "s51",
    "title": "Excel Refreshable Asset Report",
    "html": "<p>Make the asset report repeatable. Preserve the source, query, loaded Table and PivotTable as separate stages.</p><h2>Capstone refresh</h2><ol><li>Create My-work/Asset-source with a copy of Starter-assets.csv only.</li><li>From a blank workbook choose Data → Get Data → From File → From Folder. Combine compatible CSVs. Set AssetID to Text, Quantity to Whole Number, UnitCost to Decimal Number, and date fields to Date.</li><li>Add TotalValue = [Quantity]*[UnitCost]. Add OwnerCheck = if [Owner] = null or Text.Trim([Owner]) = \"\" then \"Missing owner\" else \"OK\".</li><li>Load to an Excel Table. Create the department PivotTable and chart. Check initial totals before proceeding.</li><li>Copy Data/New-assets.csv into Asset-source once. Refresh the query; wait for loaded rows; then refresh the PivotTable. Never accept a stale Pivot just because the query finished.</li><li>Save, close and reopen. Export a readable PDF and write a one-page handover with source-path recovery.</li></ol><p>Finish by explaining a source change, the resulting measure change and the missing-owner limitation without reading a script.</p>",
    "example": {
      "cells": {
        "A1": "AssetID",
        "B1": "Owner",
        "A2": "A013",
        "B2": "Alex",
        "A3": "A014",
        "B3": "Bo",
        "A4": "A015",
        "B4": ""
      },
      "formula": "=COUNTIF(B2:B4,\"\")",
      "output": "C2",
      "rule": {
        "op": "blankCount",
        "range": "B2:B4"
      },
      "editable": [
        "B4"
      ],
      "change": {
        "cell": "B4",
        "value": "Chen"
      },
      "explain": "One owner is blank. Entering a verified owner removes that missing-value flag; do not fabricate a name merely to make the check pass."
    },
    "quiz": {
      "question": "Which refresh order avoids accepting an old Pivot result?",
      "options": [
        "Pivot first, then stop",
        "Query first; wait; then PivotTable",
        "Only change the chart title"
      ],
      "correct": 1,
      "explain": "The PivotTable must refresh after the query has loaded the new source records."
    },
    "practice": [
      "Demonstrate the full folder-based asset report before and after adding New-assets.csv once.",
      "Check 15 records, 19 devices, RM17700 and missing Owner on A015.",
      "Save/reopen, export a readable PDF, write the handover and explain a change in five minutes."
    ],
    "module": "10",
    "minutes": 90
  },
  {
    "id": "x07",
    "title": "Later: LET and Formula Auditing",
    "minutes": 60,
    "module": "11",
    "html": "<p>Finish the core report first. LET names an intermediate calculation inside one formula. It does not create a workbook-scoped name.</p><h2>Explain the expression</h2><p>LET(total,SUM(A2:A4),total*B2) calculates total once and uses it with a rate. Use meaningful local names, not an abstraction that makes a short formula harder to read.</p><p>Formulas → Trace Precedents shows inputs; Trace Dependents shows users. Evaluate Formula steps through an expression, and Watch Window monitors selected cells. Remove tracing arrows afterwards. Check Automatic calculation if results seem stale; do not change workbook settings without understanding the effect.</p>",
    "example": {
      "cells": {
        "A1": "Sales",
        "B1": "Rate",
        "A2": 100,
        "A3": 200,
        "A4": 300,
        "B2": 0.1
      },
      "formula": "=LET(total,SUM(A2:A4),total*B2)",
      "output": "C2",
      "rule": {
        "op": "sumTimes",
        "range": "A2:A4",
        "rate": "B2"
      },
      "editable": [
        "B2"
      ],
      "change": {
        "cell": "B2",
        "value": 0.2
      },
      "explain": "Total sales of 600 times 10% gives 60. A 20% rate gives 120."
    },
    "quiz": {
      "question": "Where does LET’s name total exist?",
      "options": [
        "Inside that formula",
        "In every workbook globally",
        "Only as a worksheet tab"
      ],
      "correct": 0,
      "explain": "LET local names belong to the expression that declares them."
    },
    "practice": [
      "Compare LET with a helper-cell formula using identical inputs.",
      "Use Evaluate Formula and Trace Precedents; inspect a changed input.",
      "Add a Watch Window entry, verify it changes, then remove it."
    ]
  },
  {
    "id": "x08",
    "title": "Later: PMT and Finance Assumptions",
    "minutes": 60,
    "module": "11",
    "html": "<p>PMT models constant periodic payments at a constant periodic rate. It is not a loan offer or a complete cost disclosure.</p><h2>Match time units and signs</h2><p>For a fictional RM12000 principal, nominal annual rate 6% divided into monthly periods, and 12 payments, use =-PMT(B2/12,C2,A2). The negative sign presents outgoing payment as a positive amount. Defaults assume zero final balance and payment at period-end.</p><p>Annual percentage divided by 12 is an explicit nominal-rate assumption, not a rule for every quoted effective rate. Fees, changing rates, insurance and local lending terms are excluded. With zero interest, payment is principal divided by periods.</p>",
    "example": {
      "cells": {
        "A1": "Principal",
        "B1": "Nominal annual rate",
        "C1": "Monthly periods",
        "A2": 12000,
        "B2": 0.06,
        "C2": 12
      },
      "formula": "=-PMT(B2/12,C2,A2)",
      "output": "D2",
      "rule": {
        "op": "pmt",
        "principal": "A2",
        "annual": "B2",
        "periods": "C2"
      },
      "editable": [
        "B2"
      ],
      "change": {
        "cell": "B2",
        "value": 0
      },
      "explain": "At 6% nominal annual interest the monthly payment is about RM1032.80. With zero interest it is RM1000."
    },
    "quiz": {
      "question": "Which rate matches 12 monthly payments under this nominal-rate assumption?",
      "options": [
        "6% per month",
        "6%/12 per month",
        "6% × 12 per month"
      ],
      "correct": 1,
      "explain": "The periodic rate must match the payment period; keep the rate convention explicit."
    },
    "practice": [
      "Build the fictional loan model and compare zero versus positive interest.",
      "Calculate total payments and interest, clearly excluding fees.",
      "State rate convention, timing, term and missing costs in a Notes sheet."
    ]
  },
  {
    "id": "x09",
    "title": "Later: Goal Seek and the MOS Route",
    "minutes": 60,
    "module": "11",
    "html": "<p>Goal Seek finds one changeable input for a target formula result. It is approximate and may fail or return an impractical solution. Inspect the final input and residual; cancel if it violates constraints.</p><h2>Original break-even exercise</h2><p>Put units 20 in A2, margin per unit 18 in B2 and fixed cost 900 in C2. In D2 use =A2*B2-C2. Data → What-If Analysis → Goal Seek: Set cell D2, To value 0, By changing cell A2. A result of 50 units is the break-even point. The demo only recalculates an entered input; it does not implement Goal Seek.</p><h2>Optional certification</h2><p>MO-210 is an Associate performance exam. MO-211 adds workbook settings, macros, advanced charts and other objectives beyond this job-oriented course. Power Query is a valuable job skill, not a listed MO-211 objective. Use the linked official outlines for a separate gap checklist; finishing this course is not proof of complete exam readiness.</p><p>Microsoft currently lists 60-month renewal. Verify the current policy before paying. No exam questions or pass guarantees are included.</p>",
    "example": {
      "cells": {
        "A1": "Units",
        "B1": "Margin per unit",
        "C1": "Fixed cost",
        "A2": 20,
        "B2": 18,
        "C2": 900
      },
      "formula": "=A2*B2-C2",
      "output": "D2",
      "rule": {
        "op": "breakEven",
        "units": "A2",
        "margin": "B2",
        "fixed": "C2"
      },
      "editable": [
        "A2"
      ],
      "change": {
        "cell": "A2",
        "value": 50
      },
      "explain": "20 units give a loss of 540. At 50 units the result is zero."
    },
    "quiz": {
      "question": "What does Goal Seek change in this task?",
      "options": [
        "One input, A2, to target formula D2",
        "Every source cell randomly",
        "The workbook licence"
      ],
      "correct": 0,
      "explain": "Goal Seek changes one selected input for one formula target; inspect whether the solution is valid."
    },
    "practice": [
      "Run the actual Goal Seek task in Excel and check 50 units and zero residual.",
      "Change fixed cost to 990, rerun and check 55 units. Then restore the original.",
      "Read the official MOS outlines in SOURCES.md and list unpractised objectives before considering an exam."
    ]
  }
];
