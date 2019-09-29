Attribute VB_Name = "MAnimation"
'-------------------------------------------------------------------------------
' Description : Macro that shows how to animate a cell.
' Author      : Cesar Raul Urteaga-Reyesvera.
'-------------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------------- APIS
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'-------------------------------------------------------------------- PROCEDURES
' Description       : It carries out an emoticon animation within a specified
'                     cell.
Public Sub Animate()
  Dim sEmoticons As String, sCell As String
  Dim avEmoticons As Variant
  Dim lCycles As Long, lLaps As Long, lSpeed As Long
  Dim li As Long, lj As Long, lk As Long, lLeadingSpaces As Long
  '------------------ PARAMETERS
  sCell = "C7"  ' Cell where the animation will take place.
  lSpeed = 150  ' In milliseconds (i.e., 1000 is one second).
  lCycles = 2   ' Number of laps.
  lLaps = 3     ' Number of whole spins.
  '------------------ MAIN CODE
  sEmoticons = " \:,  |:,  /:,  .-.,   :\,   :|,   :/,   ._."
  avEmoticons = VBA.Split(sEmoticons, ",")
  For li = 1 To lCycles
    ' Forward
    wksAnimation.Range(sCell) = "._."
    For lj = 1 To lLaps
      For lk = LBound(avEmoticons) To UBound(avEmoticons)
        Sleep lSpeed
        lLeadingSpaces = lLeadingSpaces + 1
        wksAnimation.Range(sCell) = VBA.String(lLeadingSpaces, " ") + avEmoticons(lk)
        DoEvents
      Next lk
      lLeadingSpaces = lLeadingSpaces + 3
    Next lj
    lLeadingSpaces = lLeadingSpaces - 3
    ' Backward
    For lj = 1 To 3
      For lk = UBound(avEmoticons) To LBound(avEmoticons) Step -1
        Sleep lSpeed
        wksAnimation.Range(sCell) = VBA.String(lLeadingSpaces, " ") + avEmoticons(lk)
        lLeadingSpaces = lLeadingSpaces - 1
        DoEvents
      Next lk
      lLeadingSpaces = lLeadingSpaces - 3
    Next lj
    Sleep lSpeed
    wksAnimation.Range(sCell) = "._."
    ' Cycle end.
    lLeadingSpaces = 0
    DoEvents
  Next li
End Sub
