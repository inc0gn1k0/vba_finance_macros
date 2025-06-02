### How to Install the Modeling Shortcuts
This .xlam add-in will:
	•	Load your macros globally
	•	Register all keyboard shortcuts (Ctrl+Shift+1, 2, 5, etc.)
	•	Let you model like a wizard on any machine

⸻

## STEP 1: Open Excel VBA Editor
	1.	Open Excel
	2.	Press Alt + F11 to launch the Visual Basic for Applications editor

⸻

## STEP 2: Import Your Macro Module
	1.	In the VBA editor: File > Import File…
	2.	Select FinancialModelingShortcuts.bas from your GitHub repo
	3.	You’ll see a new module appear in the Project pane (likely named Module1 or FinancialModelingShortcuts)

⸻

## STEP 3: Register Keyboard Shortcuts
	1.	In the Project Pane, double-click ThisWorkbook
	2.	Paste the following into the Workbook_Open() event:
	
```
Private Sub Workbook_Open()
    Application.OnKey "^+1", "CtrlShift1_NumberCycle"
    Application.OnKey "^+2", "CtrlShift2_DateCycle"
    Application.OnKey "^+5", "CtrlShift5_PercentCycle"
    Application.OnKey "^+8", "CtrlShift8_MultipleCycle"
    Application.OnKey "^%+A", "CtrlAltA_Autocolour"
    Application.OnKey "^%+{UP}", "CtrlAltShift_BorderCycle"
    Application.OnKey "^%+{DOWN}", "CtrlAltShift_BorderCycle"
    Application.OnKey "^%+{LEFT}", "CtrlAltShift_BorderCycle"
    Application.OnKey "^%+{RIGHT}", "CtrlAltShift_BorderCycle"
    Application.OnKey "^+N", "CtrlShiftN_SwitchToNegative"
    Application.OnKey "^.", "CtrlPeriod_IncreaseDecimal"
    Application.OnKey "^,", "CtrlComma_DecreaseDecimal"
End Sub

```

## STEP 4: Save as an Add-In (.xlam)

	1.	In Excel (not VBA editor): File > Save As
	2.	Choose a location (e.g., your GitHub folder)
	3.	Set file type to Excel Add-In (*.xlam)
	4.	Name it: ModelingShortcuts.xlam
	5.	Click Save

⸻

## STEP 5: Install the Add-In

	1.	In Excel: File > Options > Add-ins
	2.	At the bottom: Manage: Excel Add-ins > Go
	3.	Click Browse… and select your ModelingShortcuts.xlam
	4.	Check the box ✅ to enable it
	5.	Click OK


## Directory Structure

```
/excel-shortcuts/
│
├── ModelingShortcuts.xlam
├── README.md
└── src/
    └── FinancialModelingShortcuts.bas
    
```