Attribute VB_Name = "mod_Misc"
Declare Function GetTickCount Lib "kernel32" () As Long

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
     ByVal hwnd As Long, _
     ByVal lpOperation As String, _
     ByVal lpFile As String, _
     ByVal lpParameters As String, _
     ByVal lpDirectory As String, _
     ByVal nShowCmd As Long) As Long
     
' ## used for FPS calculation only
Function TimeElapsed(Current As Long, Delay As Long, Optional CurrentTickCount As Long) As Boolean
    
    If CurrentTickCount = 0 Then
       CurrentTickCount = GetTickCount()
    End If
    
    If CurrentTickCount < 0 Then
       TimeElapsed = IIf(CurrentTickCount + Current <= Delay, True, False)
    Else
       TimeElapsed = IIf(CurrentTickCount - Current >= Delay, True, False)
    End If

End Function

