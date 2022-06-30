Attribute VB_Name = "Module1"
Sub append_files()

    'For adding a new sheet

    Worksheets.Add
    ActiveSheet.Name = "Decoy"

    'For deleting sheets already existing in the file

    Workbooks("Append_data_CW.xlsm").Activate
    Application.DisplayAlerts = False

    Dim sht As Worksheet

    For Each sht In Worksheets
        If sht.Name <> "Decoy" Then
            sht.Delete
        End If
    Next sht

    'Physically opening the file

    Sel_Imp = Application.GetOpenFilename()
    
    If Sel_Imp = False Then
        MsgBox ("No File Selcted")
    End If

    Workbooks.Open Sel_Imp

    'Copying the file

    Dim i As Integer
    Dim wkbk As Workbook, wkbks As Workbook

    Set wkbk = ActiveWorkbook
    Set wkbks = Workbooks("Append_data_CW.xlsm")

    For i = 1 To Worksheets.Count
        Worksheets(i).Copy after:=wkbks.Worksheets(wkbks.Worksheets.Count)
        wkbk.Activate
    Next i

    Application.DisplayAlerts = False
    ActiveWorkbook.Close

    'Creating the master file

    Worksheets.Add after:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "Master"

    Dim row_cnt As Long
    row_cnt = 1

    For i = 1 To 3
        Worksheets(i + 1).Range("A1").CurrentRegion.Copy Worksheets("Master").Range("A" & row_cnt)
       
        If i >= 2 Then
            Range("A" & row_cnt).EntireRow.Delete
        End If
        row_cnt = Range("A1", Range("A1").End(xlDown)).Count + 1
    Next i

    Worksheets("Master").Cells.Font.Size = 12
    Worksheets("Decoy").Delete

End Sub
