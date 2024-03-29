Attribute VB_Name = "MGetHierarchicalTreeContents"
'-------------------------------------------------------------------------------
' Description : It creates a macro that displays all the contents of a specified
'               folder recursively.
' Author      : Cesar Raul Urteaga-Reyesvera.
'-------------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------- PROCEDURES
' Description       : It creates a hierarchical tree of the contents of a
'                     specified folder.
' Parameters:
'   * rngPivotCell    : Cell that indicates the upper-left corner of the table
'                       that will have the contents.  If it is not specified,
'                       it will be the active cell.
'   * sHostFolderPath : Path of the host folder for which the information will
'                       be extracted.  If it is missing, it will prompt the user
'                       to pick out a host folder.
'   * sStartingFolder : String that states the starting path in which the pop-up
'                       window that requests the host folder will start off.
'
' N.B.: This code was based on the following information:
'
'   - https://stackoverflow.com/a/26392703
'   - https://stackoverflow.com/a/22645439
'   - https://trumpexcel.com/vba-filesystemobject/#Example-3-Get-a-List-of-All-Files-in-a-Folder
'   - https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/folder-object
'   - https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-object
'   - https://stackoverflow.com/a/23357807
'   - https://docs.microsoft.com/en-us/office/vba/api/excel.hyperlinks.add
'   - https://trumpexcel.com/sort-data-vba/
'   - https://stackoverflow.com/q/7751464
Public Sub DisplayHierarchicalContent(Optional ByRef rngPivotCell As Range _
                                                                = Nothing, _
                                      Optional ByRef sHostFolderPath As String _
                                                               = vbNullString, _
                                      Optional ByRef sStartingFolder As String _
                                                               = vbNullString)
  Dim fd As FileDialog
  Dim objFileSystem As Object, objHostFolder As Object
  Dim rng As Range
  '------------------ SETUP
  If sHostFolderPath = vbNullString Then
    ' Gets the complete path of the selected folder
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
      .Title = "Please select a folder"
      .AllowMultiSelect = False
      ' Sets the default folder to the "Desktop" folder (You can change this).
      ' e.g., Environ$("USERPROFILE") & "\Desktop\"
      .InitialFileName = IIf(sStartingFolder = vbNullString, _
                             Environ$("USERPROFILE") & "\Desktop\", _
                             sStartingFolder)
      ' Binary property that indicates whether the user selected or not a folder
      ' (if selected it takes -1; otherwise, 0).
      If .Show <> 0 Then sHostFolderPath = .SelectedItems(1) Else Exit Sub
    End With
  End If
  '
  If rngPivotCell Is Nothing Then Set rngPivotCell = ActiveCell
  '
  ' Creates the titles where we will put the information in the active cell.
  ' Type        : Type of the object (i.e., folder [D] or file [F]).
  ' Name        : Name of the object.
  ' Folder Path : Parent folder where the object is stored (displayed as a
  '               hyperlink).
  ' File Path   : If it is a file, it is its path; otherwise, it holds the
  '               same value as "Folder Path". Also, it is displayed as a
  '               hyperlink.
  ' #           : Number level of hierarchy.
  ' Hierarchy   : Graphically displays the hierarchy ("|" for a folder and
  '               "*" for a file).
  rngPivotCell.Resize(, 8) = Array("Date Created", "Date Last Modified", _
                                   "Type", "Name", "Folder Path", _
                                   "File Path", "#", "Hierarchy")
  '------------------ MAIN PROCESS
  ' Defines a file system object for the host folder.
  Set objFileSystem = CreateObject("Scripting.FileSystemObject")
  Set objHostFolder = objFileSystem.GetFolder(sHostFolderPath)
  ' Gets the host folder information.
  Set rng = rngPivotCell.Offset(1)
  WriteFileSystemObjectInfo objHostFolder, objHostFolder, rng, "D"
  ' Recursively gets the information of the contents of the host folder.
  IterateFolder objHostFolder, objHostFolder, rng
  '------------------ FORMAT
  ' Adds borders to titles.
  With rngPivotCell.Resize(, 8).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.249946592608417
    .Weight = xlThin
  End With
  ' Sorts the contents.
  With Range(rngPivotCell, rng.Offset(-1, 7))
    .Sort Key1:=ActiveCell.Offset(, 5), Order1:=xlAscending, header:=xlYes
  End With
  ' Autofits the columns and includes a filter.
  With rngPivotCell
    .Resize(, 8).AutoFilter
    .Offset(1).Resize(, 2).Columns.AutoFit
    .Offset(, 2).Resize(, 5).Columns.AutoFit
  End With
End Sub
' Description       : It gets the information of the contents of a folder
'                     recursively.
Private Sub IterateFolder(ByRef objHostFolder As Object, _
                          ByRef objFolder As Object, _
                          ByRef rng As Range)
  Dim objSubFolder As Object, objFile As Object
  '
  On Error GoTo NoSubFolders
  For Each objSubFolder In objFolder.SubFolders
    WriteFileSystemObjectInfo objHostFolder, objSubFolder, rng, "D"
    IterateFolder objHostFolder, objSubFolder, rng
  Next objSubFolder
NoSubFolders:
  '
  On Error GoTo NoFiles
  For Each objFile In objFolder.Files
    WriteFileSystemObjectInfo objHostFolder, objFile, rng
  Next objFile
NoFiles:
End Sub

'--------------------------------------------------------------------- FUNCTIONS
' Description       : It puts the information of a filesystem object in a
'                     specific range (starting in "rng").
Private Function WriteFileSystemObjectInfo(ByRef objFSFolder As Object, _
                                           ByRef objFS As Object, _
                                           ByRef rng As Range, _
                                           Optional ByVal sType As String = "F")
  Dim iHierarchyNumber As Integer
  Dim sRelevantPath As String
  '
  sRelevantPath = Replace(IIf(sType = "F", objFS.ParentFolder, objFS.Path), _
                          objFSFolder.Path, "")
  iHierarchyNumber = Len(sRelevantPath) - Len(Replace(sRelevantPath, "\", ""))
  '
  With rng
    .Value2 = objFS.DateCreated
    .Resize(, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    .Resize(, 8).HorizontalAlignment = xlLeft
    .Offset(, 0).Value2 = objFS.DateCreated
    .Offset(, 1).Value2 = objFS.DateLastModified
    .Offset(, 2).Value2 = sType
    .Offset(, 3).Value2 = objFS.Name
    .Offset(, 4).Value2 = IIf(sType = "F", objFS.ParentFolder, objFS.Path)
    .Offset(, 4).Hyperlinks.Add Anchor:=.Offset(, 4), _
      Address:=.Offset(, 4).Value2, _
      TextToDisplay:=.Offset(, 4).Value2
    .Offset(, 5).Value2 = objFS.Path
    .Offset(, 5).Hyperlinks.Add Anchor:=.Offset(, 5), _
      Address:=.Offset(, 5).Value2, _
      TextToDisplay:=.Offset(, 5).Value2
    .Offset(, 6).Value2 = iHierarchyNumber
    .Offset(, 7).Value2 = VBA.String(iHierarchyNumber - (sType = "F"), "-") + _
                          IIf(sType = "F", "*", "|")
  End With
  Set rng = rng.Offset(1)
End Function
