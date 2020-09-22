<div align="center">

## An App\.Previnstance example to call the running app to the front\.


</div>

### Description

This code was useful to me when trying to find a way to call the exsisting application in memory to the front when another instance of the same application would be attempted by a user.
 
### More Info
 
That you know about window handles.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Heath](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-heath.md)
**Level**          |Intermediate
**User Rating**    |4.9 (88 globes from 18 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-heath-an-app-previnstance-example-to-call-the-running-app-to-the-front__1-21131/archive/master.zip)





### Source Code

```
'This was found at the Microsoft Knowledgebase, Article ID: Q185730
'Paste the following code into the code Module for Form1:
Option Explicit
Private Sub Form_Load()
  If App.PrevInstance Then
   ActivatePrevInstance
  End If
End Sub
'2) Add a Standard Module to the Project.
'3) Paste the following code into the module:
Option Explicit
Public Const GW_HWNDPREV = 3
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Sub ActivatePrevInstance()
  Dim OldTitle As String
  Dim PrevHndl As Long
  Dim result As Long
  'Save the title of the application.
  OldTitle = App.Title
  'Rename the title of this application so FindWindow
  'will not find this application instance.
  App.Title = "unwanted instance"
  'Attempt to get window handle using VB4 class name.
  PrevHndl = FindWindow("ThunderRTMain", OldTitle)
  'Check for no success.
  If PrevHndl = 0 Then
   'Attempt to get window handle using VB5 class name.
   PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
  End If
  'Check if found
  If PrevHndl = 0 Then
    'Attempt to get window handle using VB6 class name
    PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
  End If
  'Check if found
  If PrevHndl = 0 Then
   'No previous instance found.
   Exit Sub
  End If
  'Get handle to previous window.
  PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
  'Restore the program.
  result = OpenIcon(PrevHndl)
  'Activate the application.
  result = SetForegroundWindow(PrevHndl)
  'End the application.
  End
End Sub
BHeath
Deffacto Web Designs Team
http://www.deffacto.com
```

