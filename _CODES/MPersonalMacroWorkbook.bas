Attribute VB_Name = "MPersonalMacroWorkbook"
'-------------------------------------------------------------------------------
' Description : It has some macros that may be used as shortcuts for the Excel's
'               Personal Macro Workbook (PMW).
' Author      : Cesar Raul Urteaga-Reyesvera.
'-------------------------------------------------------------------------------

Option Explicit

'--------------------------------------------------------------------------- PMW
' Description       : It either shows or hide the Personal Macro Workbook.
' Suggested shortcut: Ctrl + Shft + q
Public Sub ShowHidePMW()
Attribute ShowHidePMW.VB_ProcData.VB_Invoke_Func = "Q\n14"
  Windows("PERSONAL.XLSB").Visible = Not Windows("PERSONAL.XLSB").Visible
End Sub
' Description       : It exhibits a userform with information of the active
'                     workbook's sheets.
'                     Please go over the following link in order to get the
'                     FShowHideSheet form:
' https://github.com/Cesar-Urteaga/1907_EXCEL/blob/master/_WORKBOOKS/FShowHideSheets.xlsm
' Suggested shortcut: Ctrl + Shft + s
Public Sub ShowHIdeSheets()
Attribute ShowHIdeSheets.VB_ProcData.VB_Invoke_Func = "S\n14"
  FShowHideSheets.Show
End Sub

'------------------------------------------------------------------------ FORMAT
' Description       : It creates a margin in the active sheet.
' Suggested shortcut: Ctrl + Shft + w
Public Sub CreateMargin()
Attribute CreateMargin.VB_ProcData.VB_Invoke_Func = "W\n14"
  With ActiveSheet
    .Columns("A:B").ColumnWidth = 0.5
    .Rows("1:2").RowHeight = 5
    ' The following code can be omitted.
    With .[C3]
      .HorizontalAlignment = xlLeft
      .NumberFormat = "@* "":"""
      .Value = "Description"
      .Columns.AutoFit
    End With
    .Rows("4").RowHeight = 5
    .[D3].Font.Bold = True
  End With
End Sub

'--------------------------------------------------------------- GET INFORMATION
' Description       : It creates a link to the current workbook.
' Suggested shortcut: Ctrl + Shft + d
Public Sub CreateHyperlinkCurrentFile()
Attribute CreateHyperlinkCurrentFile.VB_ProcData.VB_Invoke_Func = "Z\n14"
  ActiveCell.Select
  With ActiveCell
    .Hyperlinks.Add Anchor:=Selection, _
      Address:=ActiveWorkbook.Path & "\" & ActiveWorkbook.Name, _
      TextToDisplay:=ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
  End With
End Sub

' Description       : It gets the contents of a specified folder hierarchically.
' Suggested shortcut: Ctrl + Shft + f
'                     Please go over the following link in order to get the
'                     DisplayHierarchicalContent code:
' https://github.com/Cesar-Urteaga/1907_EXCEL/blob/master/_CODES/MGetHierarchicalTreeContents.bas
Public Sub GetHierarchicalContents()
Attribute GetHierarchicalContents.VB_ProcData.VB_Invoke_Func = "D\n14"
  DisplayHierarchicalContent
End Sub
