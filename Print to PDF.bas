Attribute VB_Name = "Module1"
Sub PrinttoPDF()
Attribute PrinttoPDF.VB_Description = "Prints to PDF"
Attribute PrinttoPDF.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' PrinttoPDF Macro
' Prints to PDF
'
' Keyboard Shortcut: Ctrl+Shift+P
'

    Dim filename As String
    filename = "D:\Desktop\" & ThisWorkbook.Name
    ChDir "D:\Desktop"
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
        "D:\Desktop\test.pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
End Sub
