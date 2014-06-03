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
    'Creates the file name from the ActiveWorkbook's name
    Dim filename As String
    filename = "D:\Desktop\" & ActiveWorkbook.Name
    filename = Replace(filename, ".xlsx", "") 'Removes file extension
    filename = Replace(filename, ".xls", "")
    
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
        filename, Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
End Sub
