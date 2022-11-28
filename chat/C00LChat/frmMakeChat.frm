VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMakeChat 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Make a Chat Room"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   ScaleHeight     =   2685
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock IPTell 
      Left            =   120
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ok"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tell My Ip"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtHandle 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Text            =   " "
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Handle"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmMakeChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
' Tell's your current IP Address
MsgBox ("Your Ip is: " & IPTell.LocalIP & "")
End Sub

Private Sub cmdOk_Click()
If txtHandle.Text = " " Then
MsgBox ("You can't keep Handle field empty")
Else
strHandle = txtHandle.Text
strChatMode = "MAKE"
frmChat.Show
Unload Me
End If
End Sub

