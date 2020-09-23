Attribute VB_Name = "mod1"
Option Explicit
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
   ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_ACTIVATEAPP = &H1C
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const MAX_TOOLTIP As Integer = 64
Public Const GWL_WNDPROC = (-4)

Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * MAX_TOOLTIP
End Type
Public nfIconData As NOTIFYICONDATA
Private FHandle As Long
Private WndProc As Long
Private Hooking As Boolean

Public Sub AddIconToTray(MeHwnd As Long, MeIcon As Long, MeIconHandle As Long, Tip As String)
With nfIconData
   .hwnd = MeHwnd
   .uID = MeIcon
   .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
   .uCallbackMessage = WM_RBUTTONUP
   .hIcon = MeIconHandle
   .szTip = Tip & Chr$(0)
   .cbSize = Len(nfIconData)
End With
Shell_NotifyIcon NIM_ADD, nfIconData
End Sub

Public Sub RemoveIconFromTray()
Shell_NotifyIcon NIM_DELETE, nfIconData
End Sub

Public Sub Hook(Lwnd As Long)
If Hooking = False Then
   FHandle = Lwnd
   WndProc = SetWindowLong(Lwnd, GWL_WNDPROC, AddressOf WindowProc)
   Hooking = True
End If
End Sub

Public Sub Unhook()
If Hooking = True Then
   SetWindowLong FHandle, GWL_WNDPROC, WndProc
   Hooking = False
End If
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If Hooking = True Then
   If uMsg = WM_RBUTTONUP And lParam = WM_RBUTTONDOWN Then
      form1.SysTrayMouseEventHandler
      WindowProc = True
      Exit Function
   End If
   WindowProc = CallWindowProc(WndProc, hw, uMsg, wParam, lParam)
End If
End Function
