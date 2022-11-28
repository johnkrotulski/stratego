VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "C00L Chat v1.0 by Amey Chaugule"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "About"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Make Chat Room"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Join Chat Room"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C00L Ch@t"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' C00L Ch@t by Amey Chaugule
' 21 st April 2003
' Hope this is useful to you.
' email: ameychaugule@eth.net

Private Sub Command1_Click()
frmJoin.Show
Unload Me
End Sub
Private Sub Command2_Click()
frmMakeChat.Show
Unload Me
End Sub
Private Sub Command3_Click()
frmAbout.Show
End Sub
Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()

End Sub
