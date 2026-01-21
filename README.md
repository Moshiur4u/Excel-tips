## Excel Tips for Splite a file Name_Wize

---

Option Explicit
Sub Split_Master()
Dim iRow As Integer
Dim loopcounter As Integer
Dim Rng As Range
Dim wkb As Workbook

    iRow = Sheet1.Cells(Rows.Count, "F").End(xlUp).Row
    Set Rng = Sheet1.Range("F1:F" & iRow)
    Rng.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheet1.Range("Z1"), Unique:=True

    For loopcounter = 2 To Sheet1.Cells(Rows.Count, "Z").End(xlUp).Row

    Sheet1.Range("A1:X1").AutoFilter Field:=6, Criteria1:=Sheet1.Range("Z" & loopcounter).Value
    Set wkb = Workbooks.Add
    Sheet1.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy

    wkb.Sheets(1).Range("A1").PasteSpecial xlPasteAll
    wkb.Sheets(1).Name = Sheet1.Range("Z" & loopcounter).Value

    wkb.SaveAs "E:\Loan Adjustment_Micronsurance" & Sheet1.Range("Z" & loopcounter).Value & ".xlsx"
    wkb.Close
    Application.CutCopyMode = False
    Next loopcounter
    Sheet1.AutoFilterMode = False

## End Sub

---

Note:-ViewCode ->Insert->Model->pest Code
Excel*Sheet* SplitUsing_VBCode
Insert Code- Make Module এখনে "F" হল যে নামে ফাইল করতে চাই সেই কলাম ও AutoFilter Field:=6 হবে এর ভালু হবে "F=6" এর মান
