# VBA-Automate-Excel-Formula

# Excel Formula Automation Using VBA

This repository contains a simple VBA script to automate the process of applying the `SUM` function across multiple worksheets in an Excel workbook.

## 📁 Files Included

- `RawData.xlsx`: Sample data file.
- `AutomateSumFunction.xlsm`: Macro-enabled workbook containing the VBA script.

## ⚙️ Description

The VBA macro loops through all worksheets in the workbook and adds a `SUM` formula in the cell below the last used cell in column **F** on each sheet.

### 🧾 VBA Code Snippet

```vba
Public Sub AutomateSum()
    Dim lastcell As String
    Dim i As Integer

    i = 1
    Do While i <= Worksheets.Count
        Worksheets(i).Select

        ' Select the F2 cell of the active sheet
        Range("F2").Select
        ' Select the last cell in the column
        Selection.End(xlDown).Select
        lastcell = ActiveCell.Address(False, False)

        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "=SUM(F2:" & lastcell & ")"

        i = i + 1
    Loop
End Sub
````

## ▶️ How to Run the Macro

1. Open `AutomateSumFunction.xlsm` in Excel.
2. Navigate to **Developer** → **Visual Basic** → **Module1** to view the code.

   * Or press `F8` to step through the macro line by line.
3. To run the macro:

   * Go to **Developer** → **Macros**
   * Select `AutomateSum` and click **Run**

## 📧 Contact

* **LinkedIn**: [Max Nguyen Hoang Minh](https://www.linkedin.com/in/max-nguyen-hoang-minh)
* **Email**: [maxnguyenhoangminh@gmail.com](mailto:maxnguyenhoangminh@gmail.com)

```





