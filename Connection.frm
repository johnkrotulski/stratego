VERSION 5.00
Begin VB.Form Connection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connection"
   ClientHeight    =   1065
   ClientLeft      =   5565
   ClientTop       =   5610
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ipaddesstx 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Connection 
      Caption         =   "Enter IP Address to Connect to:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
