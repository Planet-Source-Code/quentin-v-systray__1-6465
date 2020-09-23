VERSION 5.00
Begin VB.Form form1 
   Caption         =   "Form1"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "put in Systray"
      Height          =   540
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1545
   End
   Begin VB.Menu RCPopup 
      Caption         =   "RCPopup"
      Visible         =   0   'False
      Begin VB.Menu Rest1 
         Caption         =   "Return"
      End
      Begin VB.Menu msg1 
         Caption         =   "send a message"
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Hook Me.hwnd
AddIconToTray Me.hwnd, Me.Icon, Me.Icon.Handle, "This is a test tip"
Me.Hide
End Sub


Public Sub SysTrayMouseEventHandler()
SetForegroundWindow Me.hwnd
PopupMenu RCPopup, vbPopupMenuRightButton
End Sub

Private Sub msg1_Click()
MsgBox "This is a message", vbOKOnly, "Systray"
End Sub

Private Sub Rest1_Click()
Unhook
Me.Show
RemoveIconFromTray
End Sub
