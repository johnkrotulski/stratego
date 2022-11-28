VERSION 5.00
Begin VB.Form Stratego 
   BackColor       =   &H00004000&
   Caption         =   "Stratego"
   ClientHeight    =   8415
   ClientLeft      =   600
   ClientTop       =   1935
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   13710
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   11
      Left            =   600
      TabIndex        =   127
      Top             =   7320
      Width           =   1815
   End
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   10
      Left            =   600
      TabIndex        =   126
      Top             =   7080
      Width           =   1815
   End
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   9
      Left            =   600
      TabIndex        =   125
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   8
      Left            =   600
      TabIndex        =   124
      Top             =   6600
      Width           =   1815
   End
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   7
      Left            =   600
      TabIndex        =   123
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   6
      Left            =   600
      TabIndex        =   122
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   121
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   120
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   119
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   118
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   117
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   11
      Left            =   600
      TabIndex        =   116
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   10
      Left            =   600
      TabIndex        =   115
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   9
      Left            =   600
      TabIndex        =   114
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   8
      Left            =   600
      TabIndex        =   113
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   7
      Left            =   600
      TabIndex        =   112
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   6
      Left            =   600
      TabIndex        =   111
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   110
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   109
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   108
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   107
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   106
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox EnterText 
      BackColor       =   &H00C0C0FF&
      Height          =   1335
      Left            =   10080
      TabIndex        =   105
      Top             =   6840
      Width           =   3255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5295
      Left            =   13320
      TabIndex        =   104
      Top             =   1560
      Width           =   255
   End
   Begin VB.ListBox ChatBox 
      BackColor       =   &H00C0C0FF&
      Height          =   5325
      Left            =   10080
      TabIndex        =   103
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox Player1armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   102
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Player2armies 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   101
      Top             =   1560
      Width           =   1815
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   99
      Left            =   9240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   99
      Top             =   6960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   98
      Left            =   8520
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   98
      Top             =   6960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   97
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   97
      Top             =   6960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   96
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   96
      Top             =   6960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   95
      Left            =   6360
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   95
      Top             =   6960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   94
      Left            =   5640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   94
      Top             =   6960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   93
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   93
      Top             =   6960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   92
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   92
      Top             =   6960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   91
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   91
      Top             =   6960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   90
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   90
      Top             =   6960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   89
      Left            =   9240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   89
      Top             =   6360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   88
      Left            =   8520
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   88
      Top             =   6360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   87
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   87
      Top             =   6360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   86
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   86
      Top             =   6360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   85
      Left            =   6360
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   85
      Top             =   6360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   84
      Left            =   5640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   84
      Top             =   6360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   83
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   83
      Top             =   6360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   82
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   82
      Top             =   6360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   81
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   81
      Top             =   6360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   80
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   80
      Top             =   6360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   79
      Left            =   9240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   79
      Top             =   5760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   78
      Left            =   8520
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   78
      Top             =   5760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   77
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   77
      Top             =   5760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   76
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   76
      Top             =   5760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   75
      Left            =   6360
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   75
      Top             =   5760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   74
      Left            =   5640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   74
      Top             =   5760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   73
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   73
      Top             =   5760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   72
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   72
      Top             =   5760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   71
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   71
      Top             =   5760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   70
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   70
      Top             =   5760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   69
      Left            =   9240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   69
      Top             =   5160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   68
      Left            =   8520
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   68
      Top             =   5160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   67
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   67
      Top             =   5160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   66
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   66
      Top             =   5160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   65
      Left            =   6360
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   65
      Top             =   5160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   64
      Left            =   5640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   64
      Top             =   5160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   63
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   63
      Top             =   5160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   62
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   62
      Top             =   5160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   61
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   61
      Top             =   5160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   60
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   60
      Top             =   5160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   59
      Left            =   9240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   59
      Top             =   4560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   58
      Left            =   8520
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   58
      Top             =   4560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00800000&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   57
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   57
      Top             =   4560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00800000&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   56
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   56
      Top             =   4560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   55
      Left            =   6360
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   55
      Top             =   4560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   54
      Left            =   5640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   54
      Top             =   4560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00800000&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   53
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   53
      Top             =   4560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00800000&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   52
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   52
      Top             =   4560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   51
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   51
      Top             =   4560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   50
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   50
      Top             =   4560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   49
      Left            =   9240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   49
      Top             =   3960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   48
      Left            =   8520
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   48
      Top             =   3960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00800000&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   47
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   47
      Top             =   3960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00800000&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   46
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   46
      Top             =   3960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   45
      Left            =   6360
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   45
      Top             =   3960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   44
      Left            =   5640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   44
      Top             =   3960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00800000&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   43
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   43
      Top             =   3960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00800000&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   42
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   42
      Top             =   3960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   41
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   41
      Top             =   3960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   40
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   40
      Top             =   3960
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   39
      Left            =   9240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   39
      Top             =   3360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   38
      Left            =   8520
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   38
      Top             =   3360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   37
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   37
      Top             =   3360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   36
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   36
      Top             =   3360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   35
      Left            =   6360
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   35
      Top             =   3360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   34
      Left            =   5640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   34
      Top             =   3360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   33
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   33
      Top             =   3360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   32
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   32
      Top             =   3360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   31
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   31
      Top             =   3360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   30
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   30
      Top             =   3360
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   29
      Left            =   9240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   29
      Top             =   2760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   28
      Left            =   8520
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   28
      Top             =   2760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   27
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   27
      Top             =   2760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   26
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   26
      Top             =   2760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   25
      Left            =   6360
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   25
      Top             =   2760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   24
      Left            =   5640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   24
      Top             =   2760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   23
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   23
      Top             =   2760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   22
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   22
      Top             =   2760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   21
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   21
      Top             =   2760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   20
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   20
      Top             =   2760
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   19
      Left            =   9240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   19
      Top             =   2160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   18
      Left            =   8520
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   18
      Top             =   2160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   17
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   17
      Top             =   2160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   16
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   16
      Top             =   2160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   15
      Left            =   6360
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   15
      Top             =   2160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   14
      Left            =   5640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   14
      Top             =   2160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   13
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   13
      Top             =   2160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   12
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   12
      Top             =   2160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   11
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   11
      Top             =   2160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   10
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   10
      Top             =   2160
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   9
      Left            =   9240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   9
      Top             =   1560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   8
      Left            =   8520
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   1560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   7
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   1560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   6
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   1560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   5
      Left            =   6360
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   1560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   4
      Left            =   5640
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   1560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   3
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   1560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   2
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   1560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   0
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   1560
      Width           =   732
   End
   Begin VB.PictureBox box 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C0C0FF&
      Height          =   612
      Index           =   1
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   1560
      Width           =   732
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   5880
      Picture         =   "stratego.frx":0000
      Top             =   7680
      Width           =   1110
   End
   Begin VB.Image Chat 
      Height          =   540
      Left            =   11160
      Picture         =   "stratego.frx":118E
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Turn 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      Caption         =   "Player One's Turn"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5760
      TabIndex        =   100
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Image Armies 
      Height          =   525
      Left            =   840
      Picture         =   "stratego.frx":2144
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   4320
      Picture         =   "stratego.frx":33E0
      Top             =   120
      Width           =   4725
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSurrender 
         Caption         =   "&Surrender"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuOfferDraw 
         Caption         =   "&Offer Draw"
      End
   End
   Begin VB.Menu mnuConnection 
      Caption         =   "&Connection"
      Begin VB.Menu mnuPing 
         Caption         =   "&Ping"
      End
      Begin VB.Menu mnuNewPlayer 
         Caption         =   "&New Player"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutStratego 
         Caption         =   "&About Stratego"
      End
      Begin VB.Menu mnuAboutAuthor 
         Caption         =   "&About The Author"
      End
   End
End
Attribute VB_Name = "stratego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim Position(0 To 99) As Integer
Dim FirstPlace As Integer, SecondPlace As Integer
Dim Mode As Integer




Private Sub box_Click(Index As Integer)
    
    If Mode = 1 Then
    
        FirstPlace = Index
        Mode = 2
        
    End If
    
    If Mode = 2 Then
        SecondPlace = Index
        Mode = 1
        Position(SecondPlace) = Position(FirstPlace)
        
        box(SecondPlace).FontSize = 24
        box(SecondPlace).Print Str(Position(SecondPlace))
        box(FirstPlace).Cls
        Position(FirstPlace) = 100
        
        
    End If
    

End Sub






Private Sub Form_Load()

    stratego.Show
    Connection.Show
    frmSplash.Show
    
    
End Sub





Private Sub mnuAboutStratego_Click()
    frmAbout.Show
End Sub

Private Sub mnuNewPlayer_Click()
    Connection.Show
End Sub
