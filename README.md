<div align="center">

## Locking multiple files


</div>

### Description

Program opens text file for input, reads name of files in list, then locks those

files. Uses form and module, also shows system tray icon.
 
### More Info
 
Code is NOT FINAL, must be customized to needs.

Have picture box as picture1 hidden and the picture be your system tray icon.

Have timer as timer1 and set interval for time to be shown.

IMPORTANT SYSTEM TRAY!!! Have a picture box as picHook and visible = false.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tim Fischer](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tim-fischer.md)
**Level**          |Unknown
**User Rating**    |4.2 (165 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tim-fischer-locking-multiple-files__1-4239/archive/master.zip)

### API Declarations

```
Option Explicit
'System tray stuff
Type NOTIFYICONDATA
 cbSize       As Long
 hwnd        As Long
 uID         As Long
 uFlags       As Long
 uCallbackMessage  As Long
 hIcon        As Long
 szTip        As String * 64
End Type
Public Const NIM_ADD = 0
Public Const NIM_MODIFY = 1
Public Const NIM_DELETE = 2
Public Const NIF_MESSAGE = 1
Public Const NIF_ICON = 2
Public Const NIF_TIP = 4
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const WM_NULL = &H0
Declare Function Shell_NotifyIconA Lib "shell32" _
(ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer
Declare Function PostMessage Lib "user32" _
 Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
 ByVal wParam As Long, ByVal lParam As Long) As Long
```


### Source Code

```
Private Sub Form_Load()
Dim fileLock As String
Open "C:\Text.txt" For Input As #1 ' This is the file that it will read from.
Do While Not EOF(1) ' Loop until end of file.
  Line Input #1, fileLock 'Each line of the file is the path name
  FileNumber = FreeFile() 'Findout next available file number
  Open fileLock For Binary Shared As #FileNumber
  Lock #FileNumber 'Lock file
  Loop
Close #1
'System tray stuff
Dim nd As NOTIFYICONDATA
 Dim lRet As Long
 With nd
  .cbSize = Len(nd)
  .hwnd = picHook.hwnd
  .uID = 1&
  .szTip = "Lock on" & Chr(0)
  .uCallbackMessage = WM_MOUSEMOVE
  .hIcon = Picture1.Picture 'Icon for system tray
  .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
 End With
 lRet = Shell_NotifyIconA(NIM_ADD, nd)
 'Error check here
 'lRet = PostMessage(mnuPophwnd, WM_NULL, 0&, 0&) 'hrmf
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Dim nd As NOTIFYICONDATA
 Dim iRet As Integer
 With nd
  .cbSize = Len(nd)
  .hwnd = picHook.hwnd
  .uID = 1&
 End With
  iRet = Shell_NotifyIconA(NIM_DELETE, nd)
  If FreeFile() <> 1 Then 'Remove files from memory
  For X = 1 To FreeFile() - 1
  Close #X
  Next
  End If
End Sub
Private Sub Timer1_Timer() 'Puts form in background
frmSplash.Hide
Timer1.Enabled = False
End Sub
Private Sub picHook_MouseMove(Button As Integer, Shift As Integer, _
  X As Single, Y As Single)
 Static bRunning As Boolean
 Dim lMsg As Long
 lMsg = X / Screen.TwipsPerPixelX
 If Not (bRunning) Then 'avoid cascades
  bRunning = True
  Select Case lMsg
   Case WM_LBUTTONDBLCLK:
   If InputBox("Please enter Password:", "Lock") = "password" Then Unload Me 'Password check
   Case WM_LBUTTONDOWN:
   Case WM_LBUTTONUP:
   Case WM_RBUTTONDBLCLK:
   Case WM_RBUTTONDOWN:
   End Select
    bRunning = False
  End If
End Sub
```

