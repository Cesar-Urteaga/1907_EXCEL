VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FShowHideSheets 
   Caption         =   "Show/Hide Worksheets"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8205
   OleObjectBlob   =   "FShowHideSheets.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FShowHideSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Description : Userform that depicts information regarding the sheets of the
'               active workbook.
' Author      : Cesar Raul Urteaga-Reyesvera.
'-------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------------ MODULE DECLARATIONS
Enum fshsListSelection
  fshsListSelectionTrue = 1
  fshsListSelectionFalse
  fshsListSelectionInverse
End Enum

Private Sub lblNumberOfVeryHiddenSheets_Click()

End Sub

'------------------------------------------------------------------------ EVENTS
Private Sub lstSheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim li As Long
  '
  For li = 0 To lstSheets.ListCount - 1
    If lstSheets.Selected(li) Then
      ActiveWorkbook.Sheets(lstSheets.List(li)).Visible = -1
      ActiveWorkbook.Sheets(lstSheets.List(li)).Activate
      Unload Me
      Exit For
    End If
  Next li
End Sub

Private Sub lstSheets_MouseDown(ByVal Button As Integer, _
                                ByVal Shift As Integer, _
                                ByVal X As Single, ByVal Y As Single)
  If Button = vbKeyRButton Then
    ChangeListSelection fshsListSelectionInverse
  End If
End Sub

Private Sub lstSheets_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  Dim li As Long
  On Error Resume Next
  '
  Select Case KeyAscii
    Case 13:
      For li = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(li) Then
          ActiveWorkbook.Sheets(lstSheets.List(li)).Visible = -1
          ActiveWorkbook.Sheets(lstSheets.List(li)).Activate
          Unload Me
          Exit For
        End If
      Next li
    Case 27: Unload Me
    Case Asc("A"), Asc("a"):
      ChangeListSelection fshsListSelectionTrue
    Case Asc("I"), Asc("i"):
      ChangeListSelection fshsListSelectionInverse
    Case Asc(" "):
      ChangeListSelection fshsListSelectionFalse
    Case Asc("S"), Asc("s"):
      Call ShowHIdeSheets(xlSheetVisible)
      Call GetVisibleHiddenVeryHiddenSheets
      Call GetWorksheetNames
    Case Asc("H"), Asc("h"):
      Call ShowHIdeSheets(xlSheetHidden)
      Call GetVisibleHiddenVeryHiddenSheets
      Call GetWorksheetNames
    Case Asc("V"), Asc("v"):
      Call ShowHIdeSheets(xlSheetVeryHidden)
      Call GetVisibleHiddenVeryHiddenSheets
      Call GetWorksheetNames
  End Select
End Sub

Private Sub UserForm_Initialize()
  Call GetVisibleHiddenVeryHiddenSheets
  '
  Call CreateListBoxHeader(lstSheets, lstHeaders, _
                           Array("Sheet Name", "Visibility", "Non-empty Cells"))
  Call GetWorksheetNames
End Sub

'-------------------------------------------------------- PROCEDURES & FUNCTIONS
Private Sub GetVisibleHiddenVeryHiddenSheets()
  Dim lv As Long, lh As Long, lvh As Long
  Dim wks As Worksheet
  '
  For Each wks In ActiveWorkbook.Sheets
    Select Case wks.Visible
      Case xlSheetVisible: lv = lv + 1
      Case xlSheetHidden: lh = lh + 1
      Case xlSheetVeryHidden: lvh = lvh + 1
    End Select
  Next wks
  '
  lblNumberOfSheets.Caption = "Total number of sheets (Visible): " + _
                              CStr(ActiveWorkbook.Sheets.Count) + _
                              " (" + Format(lv, "#,#") + ")"
  '
  lblNumberOfHiddenSheets.Visible = CBool(lh)
  If lh Then lblNumberOfHiddenSheets.Caption = "# of hidden sheets: " + _
                                               Format(lh, "#,#")
  '
  lblNumberOfVeryHiddenSheets.Visible = CBool(lvh)
  If lvh Then lblNumberOfVeryHiddenSheets.Caption = "# of very hidden sheets: " _
                                                    + Format(lvh, "#,#")
End Sub

' Based on the solution of "Johas_Hess".
' Please refer to https://stackoverflow.com/a/43381634
Private Sub CreateListBoxHeader(ByRef lstBody As MSForms.ListBox, _
                                ByRef lstHeader As MSForms.ListBox, _
                                ByRef asHeader As Variant)
  Dim li As Long
  '
  With lstHeader
    ' Dimension both listboxes equally.
    .ColumnCount = lstBody.ColumnCount
    .ColumnWidths = lstBody.ColumnWidths
    ' Adds headers.
    .Clear
    .AddItem
    For li = 0 To UBound(asHeader)
      .List(0, li) = asHeader(li)
    Next li
    ' Sets the header up.
    lstBody.ZOrder (1)
    lstHeader.ZOrder (0)
    lstHeader.SpecialEffect = fmSpecialEffectFlat
    lstHeader.BackColor = RGB(240, 240, 240)
    lstHeader.Height = 10
    ' Makes the alignment between both listboxes.
    lstHeader.Width = lstBody.Width
    lstHeader.Left = lstBody.Left
    lstHeader.Top = lstBody.Top - (lstHeader.Height - 1)
  End With
End Sub

Private Sub GetWorksheetNames()
  Dim wks As Worksheet
  Dim li As Long
  '
  lstSheets.Clear
  li = 0
  For Each wks In ActiveWorkbook.Sheets
    With lstSheets
      .AddItem wks.Name
      .List(li, 0) = wks.Name
      .List(li, 1) = Choose(wks.Visible + 2, " ", "h", "", "vh")
      .List(li, 2) = Format(WorksheetFunction.CountA(wks.UsedRange), "#,#")
    End With
    li = li + 1
  Next wks
End Sub

Private Sub ChangeListSelection(ByVal uSelection As fshsListSelection)
  Dim li As Long
  '
  Select Case uSelection
    Case fshsListSelectionTrue, fshsListSelectionFalse:
      For li = 0 To lstSheets.ListCount - 1
        lstSheets.Selected(li) = uSelection - 2
      Next li
    Case Else:
      For li = 0 To lstSheets.ListCount - 1
        lstSheets.Selected(li) = Not lstSheets.Selected(li)
      Next li
  End Select
End Sub

Private Sub ShowHIdeSheets(ByVal iVisibility As Integer)
  Dim li As Long
  '
  For li = 0 To lstSheets.ListCount - 1
    If lstSheets.Selected(li) Then
      ActiveWorkbook.Sheets(lstSheets.List(li)).Visible = iVisibility
    End If
  Next li
End Sub


