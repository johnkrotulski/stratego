VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2610
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   2610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox iptxt 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tell IP"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()

Do While IP = IP
        
            Input #1, IP
            
            
Loop

End Sub

Private Sub Form_Load()

Shell "getIP.bat"
Open "ipconfig.txt" For Input As #1

End Sub
