VERSION 5.00
Begin VB.Form frmJoin 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConnect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Connect"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtIP 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtHandle 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   " "
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Chat Room IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Handle"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
If txtHandle.Text = " " Then
MsgBox ("You can't keep Handle field empty")
Else
strHandle = txtHandle.Text
strJoinIP = txtIP.Text
strChatMode = "JOIN"
frmChat.Show
Unload Me
End If
End Sub

