VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmChat 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cnt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Connect"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Send 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Send"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Reciever 
      Index           =   0
      Left            =   600
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Sender 
      Left            =   120
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   7695
   End
   Begin VB.ListBox lstChat 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5685
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Connected to:"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label HisIP 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label tIT 
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
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumSockets As Integer
Dim SysMsg As String
Private Sub Cnt_Click()
' Tells Sender to connet to the IP which is connected to Reciever
Sender.RemoteHost = "" & HisIP.Caption & ""
Sender.RemotePort = 8888
Sender.Connect ' Connect's Sender ofcourse
SysMsg = "Connected ..."
lstChat.AddItem (SysMsg)
' These labels tell user which IP he is connected to
Label2.Visible = True
HisIP.Visible = True
Cnt.Visible = False
End Sub
Private Sub Command1_Click()
' Disconnect
SysMsg = " Disconnected ... "
lstChat.AddItem (SysMsg)
Sender.Close
End Sub
Private Sub Form_Load()
Select Case strChatMode
Case "MAKE"
Make
Case "JOIN"
Join
End Select
frmChat.Caption = "C00L Ch@t v1.0 by Amey Chaugule User Is " & strHandle & " and Date is " & Month(Now) & "/" & Day(Now) & "/" & Year(Now) & ""
End Sub
Private Sub Make()
' Event's to follow when user chooses to make a Chatroom
SysMsg = "Made a ChatRoom; IP:" & Sender.LocalIP & ""
lstChat.AddItem (SysMsg)
Reciever(0).LocalPort = 9999 'Set's the Port to which to listen to
Reciever(0).Listen
End Sub
Private Sub Join()
' Event's to follow when user chooses to Join a Chatroom
Sender.RemoteHost = "" & strJoinIP & ""
Sender.RemotePort = 9999 ' Sets the port on which Sender will connect to Reciever of other client CAUTION: KEEP THE PORT OF maker's Reciever and joiners sender same , vice versa
Sender.Connect
Reciever(0).LocalPort = 8888
SysMsg = "Joined ChatRoom with IP " & strJoinIP & ""
lstChat.AddItem (SysMsg)
Reciever(0).Listen
End Sub
Private Sub Reciever_Close(Index As Integer)
SysMsg = "Disconnected"
lstChat.AddItem (SysMsg)
' Action to be taken in event of disconnection
Reciever(Index).Close
Unload Reciever(Index)
NumSockets = NumSockets - 1
End Sub
Private Sub Reciever_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Index = 0 Then
NumSockets = NumSockets + 1
Load Reciever(NumSockets)
Reciever(NumSockets).Accept requestID ' Accept the connection request
HisIP.Caption = Reciever(NumSockets).RemoteHostIP ' Set the value to HisIP label to the IP which is connected to Reciever
SysMsg = "Reciever has accepted connection request..."
lstChat.AddItem (SysMsg)
If strChatMode = "MAKE" Then
' If user has made chatroom then make Connect command visible
Cnt.Visible = True
SysMsg = "Please click on connect to connect to the joiner ..."
lstChat.AddItem (SysMsg)
End If
End If
End Sub
Private Sub Reciever_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Reciever(Index).GetData strData, vbString ' Gets data( chat message ) from the other client
lstChat.AddItem (strData)
End Sub
Private Sub Send_Click()
Dim strSend As String
' Joins Handle and Message
strSend = "" & strHandle & ":" & txtSend.Text & ""
'  Sends Chat Message
Sender.SendData strSend
' Add the message to our lstChat
lstChat.AddItem (strSend)
txtSend.Text = ""
End Sub
