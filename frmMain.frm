VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galaxy Telnet Chat Server"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer OffOut 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7440
      Top             =   720
   End
   Begin VB.Timer OffIn 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6960
      Top             =   720
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   8175
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "out"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "in"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "listen"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   185
         TabIndex        =   5
         Top             =   120
         Width           =   375
      End
      Begin VB.Shape lOut 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   960
         Shape           =   2  'Oval
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape lIn 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   600
         Shape           =   2  'Oval
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape lListen 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   240
         Shape           =   2  'Oval
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "Loading, Please Wait..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   6735
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   3889
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Menu"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5160
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
      Begin VB.CommandButton Command2 
         Caption         =   "&copy ip address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1560
         TabIndex        =   13
         Top             =   1440
         Width           =   1350
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&local login"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1350
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Roomname:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSWinsockLib.Winsock Listener 
      Left            =   6480
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   0
      Left            =   6000
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Servers As Integer
Dim LastMessage2 As String
Dim LastMessage1 As String
Dim Lastmessage0 As String
Dim PeopleInChat As Integer


Private Sub Command1_Click()
Shell "telnet localhost", vbNormalFocus
End Sub

Private Sub Command2_Click()
Clipboard.SetText Listener.LocalIP
End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then
MsgBox "Galaxy Telnet Chat Server v2.1 is already loaded.", vbExclamation, "Error": End
End If


txtName.Text = GetSetting("Galaxy Telnet Chat", "Settings", "RoomName", "Galaxy")
txtPassword.Text = GetSetting("Galaxy Telnet Chat", "Settings", "RoomPassword")


Status "Preparing network connections.."
Listener.LocalPort = 23
Listener.Listen
lListen.FillColor = vbGreen
Status "Listening for connections on port 23"
Status "Server ready."
lblStatus = "Waiting for incomming connections (" & Listener.LocalIP & ":23)"

End Sub

Private Sub Status(Text As String)
Text1 = Text1 & Time & ":  " & Text & vbCrLf

End Sub

Private Sub Label3_Click()
End Sub

Private Sub Label9_Click()
Shell "telnet localhost", vbNormalFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("This is close all connections to server, are you sure you want to exit?", vbExclamation + vbYesNo, "Exit Warning") = vbYes Then
End
Else
Cancel = 1
End If

End Sub

Private Sub Listener_Close()
lListen.FillColor = vbRed
End Sub

Private Sub Listener_ConnectionRequest(ByVal requestID As Long)
Status "Incomming connection request from " & Listener.RemoteHostIP
InTraffic

Dim TheX As Integer

If Not ListView1.ListItems.Count = 0 Then
Do
TheX = TheX + 1
If ListView1.ListItems(TheX).ForeColor = vbRed Then
ListView1.ListItems(TheX).ForeColor = vbBlack
ListView1.ListItems(TheX).SubItems(1) = ""
ListView1.ListItems(TheX).SubItems(2) = ""
ListView1.ListItems(TheX).SubItems(3) = "Waiting for escapes"
ListView1.ListItems(TheX).SubItems(4) = "0"
ListView1.ListItems(TheX).SubItems(5) = "[1;37;40m"
Load Server(TheX)
Server(TheX).Accept requestID
SendData "GALAXY TELNET SERVER V2.1" & vbCrLf & "This server requires you to have local echo enabled" & vbCrLf & vbCrLf & "Press ESCAPE twice to continue", TheX
SendData Chr$(27) & "[4i", TheX
Exit Sub
End If

Loop Until ListView1.ListItems.Count = TheX
End If

Servers = Servers + 1
Load Server(Servers)
Server(Servers).Accept requestID

ListView1.ListItems.Add , , Servers
ListView1.ListItems(Servers).SubItems(3) = "Waiting for escapes"
ListView1.ListItems(Servers).SubItems(4) = "0"
ListView1.ListItems(Servers).SubItems(5) = "[1;37;40m"
SendData "GALAXY TELNET SERVER V2.1" & vbCrLf & "This server requires you to have local echo enabled" & vbCrLf & vbCrLf & "Press ESCAPE twice to continue", Servers
SendData Chr$(27) & "[4i", Servers
End Sub

Private Sub SendMenu(MenuPath As String, ServerSocket As Integer)

Dim ANSI As String
Open MenuPath For Input As #1
Input #1, ANSI
Close #1
ppl$ = PeopleInChat

On Error Resume Next
ANSI = RemoveString(ANSI, "%NAME%", ListView1.ListItems(ServerSocket).SubItems(1))
ANSI = RemoveString(ANSI, "%ROOM%", txtName)
ANSI = RemoveString(ANSI, "%PEOPLE%", ppl$)

SendData ANSI, ServerSocket

'GET USERLIST
If ListView1.ListItems(ServerSocket).SubItems(3) = "Chatroom" Then SendUserList
'END GET USERLIST

End Sub
Private Sub SendUserList(Optional Goodbye As Boolean)

Dim TheX As Integer
Dim FoundX As Integer
Dim CurrentName As String * 18
Dim ThisPerson As String
Do
TheX = TheX + 1
If ListView1.ListItems(TheX).SubItems(3) = "Chatroom" Then
FoundX = FoundX + 1
CurrentName = ListView1.ListItems(TheX).SubItems(1)
ThisPerson = ThisPerson & Chr$(27) & "[" & 8 + FoundX & ";59f" & CurrentName
End If
Loop Until TheX = ListView1.ListItems.Count

If Goodbye = True Then
ThisPerson = ThisPerson & Chr$(27) & "[" & FoundX + 9 & ";59f" & "                  "
End If

TheX = 0

Do
TheX = TheX + 1
If ListView1.ListItems(TheX).SubItems(3) = "Chatroom" Then
SendData ThisPerson & Chr$(27) & "[4;25f", TheX
End If

Loop Until TheX = ListView1.ListItems.Count

End Sub

Private Sub Listener_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
lListen.FillColor = vbRed
End Sub

Private Sub OffOut_Timer()
lOut.FillColor = &HE0E0E0
End Sub

Private Sub Server_Close(Index As Integer)
Status "Conncetion with " & ListView1.ListItems(Index).SubItems(1) & " closed"
If ListView1.ListItems(Index).SubItems(3) = "Chatroom" Then
PeopleInChat = PeopleInChat - 1
SendStatus ListView1.ListItems(Index).SubItems(1) & " has left the chatroom", , Index
End If

ListView1.ListItems(Index).SubItems(3) = "Connection closed"
ListView1.ListItems(Index).ForeColor = vbRed
SendUserList True
Unload Server(Index)

End Sub

Private Sub Server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
' 21,6 = STATUS (MAX LEN 50)
' 9,59 = USERLIST (MAX LEN 18)
' 4,25 = TYPE TEXT (MAX LEN 50
' 9,5 = TOP CHAT
' 19,5 = BOTTOM CHAT


Dim incData As String
Server(Index).GetData incData
InTraffic

'WAITING FOR ESCAPES
If ListView1.ListItems(Index).SubItems(3) = "Waiting for escapes" Then
If incData = Chr$(27) Then
If ListView1.ListItems(Index).Tag = "" Then ListView1.ListItems(Index).Tag = "1": Exit Sub
If ListView1.ListItems(Index).Tag = "1" Then
ListView1.ListItems(Index).Tag = ""

ListView1.ListItems(Index).SubItems(3) = "Waiting for name"
SendMenu App.Path & "\login1.ans", Index
Exit Sub
End If
End If
End If


'WAITING FOR USER TO TYPE THEIR NAME
If ListView1.ListItems(Index).SubItems(3) = "Waiting for name" Then

If incData = vbCrLf Or incData = Chr$(13) Then
ListView1.ListItems(Index).ForeColor = vbGreen
If Not txtPassword = "" Then
ListView1.ListItems(Index).SubItems(3) = "Password"
SendMenu App.Path & "\login2.ans", Index
Exit Sub
End If

ListView1.ListItems(Index).SubItems(3) = "Welcome"
SendMenu App.Path & "\welcome.ans", Index
Exit Sub
End If



If incData = vbBack Then
If ListView1.ListItems(Index).SubItems(1) = "" Then Exit Sub
ListView1.ListItems(Index).SubItems(1) = Mid(ListView1.ListItems(Index).SubItems(1), 1, Len(ListView1.ListItems(Index).SubItems(1)) - 1)
'SendData vbBack & " " & vbBack, Index
Exit Sub
End If

If Len(ListView1.ListItems(Index).SubItems(1)) >= 18 Then Exit Sub
ListView1.ListItems(Index).SubItems(1) = ListView1.ListItems(Index).SubItems(1) & incData


End If
'END OF WAITING FOR USER TO TYPE THEIR NAME

If ListView1.ListItems(Index).SubItems(3) = "Password" Then

If incData = vbCrLf Or incData = Chr$(13) Then
If ListView1.ListItems(Index).Tag = txtPassword.Text Then
ListView1.ListItems(Index).SubItems(3) = "Welcome"
ListView1.ListItems.Item(Index).Tag = ""
Status "User " & ListView1.ListItems(Index).SubItems(1) & " (" & Server(Index).RemoteHostIP & ") logged in"
SendMenu App.Path & "\welcome.ans", Index
Exit Sub
Else
SendMenu App.Path & "\login2.ans", Index
ListView1.ListItems.Item(Index).Tag = ""
Exit Sub
End If
End If

ListView1.ListItems.Item(Index).Tag = ListView1.ListItems.Item(Index).Tag & incData
Exit Sub
End If


If ListView1.ListItems(Index).SubItems(3) = "Welcome" Then
SendStatus ListView1.ListItems(Index).SubItems(1) & " has entered the room"
ListView1.ListItems(Index).SubItems(3) = "Chatroom"
PeopleInChat = PeopleInChat + 1
Status "User " & ListView1.ListItems(Index).SubItems(1) & " (" & Server(Index).RemoteHostIP & ") logged in"
SendMenu App.Path & "\chatroom.ans", Index
Exit Sub
End If

'IF IN CHATROOM
If ListView1.ListItems(Index).SubItems(3) = "Chatroom" Then

If incData = vbCrLf Or incData = Chr$(13) Then

If ListView1.ListItems(Index).Tag = "" Then SendData Chr$(27) & "[4;25f" & BlankString & Chr$(27) & "[4;25f", Index: Exit Sub

If ListView1.ListItems(Index).Tag = "help" Then
SendStatus "Commands: colour, redraw, exit", Index
ListView1.ListItems(Index).Tag = ""
Exit Sub
End If

If Mid(ListView1.ListItems(Index).Tag, 1, 6) = "colour" Then
TEMP05$ = ListView1.ListItems(Index).Tag
If Len(ListView1.ListItems(Index).Tag) <= 7 Or IsNumeric(Right(TEMP05$, 1)) = False Then SendStatus "Error, use 'colour n' (n = number from 0 to 7)", Index: ListView1.ListItems(Index).Tag = "": Exit Sub


ListView1.ListItems(Index).Tag = ""
If Right(TEMP05$, 1) = "0" Then ListView1.ListItems(Index).SubItems(5) = "[2;39;40m": SendStatus "Colour change to default", Index
If Right(TEMP05$, 1) = "1" Then ListView1.ListItems(Index).SubItems(5) = "[1;37;40m": SendStatus "Colour change to white", Index
If Right(TEMP05$, 1) = "2" Then ListView1.ListItems(Index).SubItems(5) = "[1;31;40m": SendStatus "Colour change to red", Index
If Right(TEMP05$, 1) = "3" Then ListView1.ListItems(Index).SubItems(5) = "[1;32;40m": SendStatus "Colour change to green", Index
If Right(TEMP05$, 1) = "4" Then ListView1.ListItems(Index).SubItems(5) = "[1;33;40m": SendStatus "Colour change to yellow", Index
If Right(TEMP05$, 1) = "5" Then ListView1.ListItems(Index).SubItems(5) = "[1;34;40m": SendStatus "Colour change to blue", Index
If Right(TEMP05$, 1) = "6" Then ListView1.ListItems(Index).SubItems(5) = "[1;35;40m": SendStatus "Colour change to pink", Index
If Right(TEMP05$, 1) = "7" Then ListView1.ListItems(Index).SubItems(5) = "[1;36;40m": SendStatus "Colour change to cyan", Index

Exit Sub
End If

If ListView1.ListItems(Index).Tag = "exit" Then
Status ListView1.ListItems(Index).SubItems(1) & " exited that chatroom"
PeopleInChat = PeopleInChat - 1
If ListView1.ListItems(Index).SubItems(3) = "Chatroom" Then
SendStatus ListView1.ListItems(Index).SubItems(1) & " exited that chatroom", , Index
End If


ListView1.ListItems(Index).SubItems(3) = "User exited chat"
ListView1.ListItems(Index).ForeColor = vbRed
SendUserList True
Unload Server(Index)
ListView1.ListItems(Index).Tag = ""
Exit Sub
End If

If ListView1.ListItems(Index).Tag = "redraw" Then SendMenu App.Path & "\chatroom.ans", Index: ListView1.ListItems(Index).SubItems(4) = 0: ListView1.ListItems(Index).Tag = "": Exit Sub

SendChat ListView1.ListItems(Index).Tag, Index
ListView1.ListItems(Index).Tag = ""
Exit Sub
End If

ListView1.ListItems(Index).Tag = ListView1.ListItems(Index).Tag & incData

End If

'SendData incData, Index

End Sub
Private Sub SendStatus(Status1 As String, Optional ServerSocket As Integer, Optional NotMe As Integer)
Dim TheX As Integer
Dim Status2 As String * 50
Dim BlankString As String * 50
BlankString = " "

Status2 = Status1

If ServerSocket = 0 Then
Do
TheX = TheX + 1
If ListView1.ListItems(TheX).SubItems(3) = "Chatroom" Then
If Not TheX = NotMe Then SendData Chr$(27) & "[21;6f" & Status2, TheX
End If
Loop Until TheX = ListView1.ListItems.Count

Else
Do
If Len(Status1) < 50 Then
Status1 = Status1 & " "
End If
Loop Until Len(Status1) = 50

SendData Chr$(27) & "[21;6f" & Chr$(27) & "[0m" & Status1 & Chr$(27) & "[4;25f" & BlankString & Chr$(27) & "[4;25f", ServerSocket
End If

End Sub

Private Sub SendChat(Text As String, ServerSocket As Integer)
Dim BlankString As String * 50
Dim ThisSize As Integer
Dim ChatText As String * 61
Dim NameHEader As String
Dim NameHEader2 As String

'NameHEader = Chr$(27) & "[1;37;40m"
NameHEader = Chr$(27) & "[0m"
BlankString = " "
SendData Chr$(27) & "[4;25f" & BlankString & Chr$(27) & "[4;25f" & Chr$(27) & "[0m", ServerSocket
Dim TheX As Integer
Dim THeX2 As Integer

ChatText = ListView1.ListItems(ServerSocket).SubItems(1) & ": " & Chr$(27) & ListView1.ListItems(ServerSocket).SubItems(5) & Text
LastMessage2 = LastMessage1
LastMessage1 = Lastmessage0
Lastmessage0 = NameHEader & ChatText

Do
TheX = TheX + 1
If ListView1.ListItems(TheX).SubItems(3) = "Chatroom" Then

If ListView1.ListItems(TheX).SubItems(4) >= 10 Then
ListView1.ListItems(TheX).SubItems(4) = "2"
SendData Chr$(27) & "[9;5f" & LastMessage2 & Chr$(27) & "[10;5f" & LastMessage1 & Chr$(27) & "[11;5f" & BlankString & Chr$(27) & "[12;5f" & BlankString & Chr$(27) & "[13;5f" & BlankString & Chr$(27) & "[14;5f" & BlankString & Chr$(27) & "[15;5f" & BlankString & Chr$(27) & "[16;5f" & BlankString & Chr$(27) & "[17;5f" & BlankString & Chr$(27) & "[18;5f" & BlankString & Chr$(27) & "[19;5f" & BlankString, TheX
End If

ListView1.ListItems(TheX).SubItems(4) = ListView1.ListItems(TheX).SubItems(4) + 1
SendData Chr$(27) & "[" & 8 + ListView1.ListItems(TheX).SubItems(4) & ";5f" & NameHEader & ChatText & Chr$(27) & "[4;25f" & Chr$(27) & "[0m", TheX
End If

Loop Until TheX = ListView1.ListItems.Count

End Sub
Private Sub Server_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Status "Conncetion with " & ListView1.ListItems(Index).SubItems(1) & " lost due to error"

If ListView1.ListItems(Index).SubItems(3) = "Chatroom" Then
PeopleInChat = PeopleInChat - 1
SendStatus "An error occured with " & ListView1.ListItems(Index).SubItems(1) & "'s connection", , Index
End If


ListView1.ListItems(Index).SubItems(3) = "Connection error"
ListView1.ListItems(Index).ForeColor = vbRed
SendUserList True
Unload Server(Index)

End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub OffIn_Timer()
lIn.FillColor = &HE0E0E0
End Sub

Private Sub InTraffic()
lIn.FillColor = vbYellow
OffIn.Enabled = False
OffIn.Enabled = True
End Sub

Private Sub OutTraffic()
lOut.FillColor = vbYellow
OffOut.Enabled = False
OffOut.Enabled = True
End Sub

Private Sub SendData(Data As String, ServerSocket As Integer)
On Error Resume Next
OutTraffic
Server(ServerSocket).SendData Data
End Sub

Function RemoveString(Entire As String, Word As String, Replace As String) As String
    Dim I As Integer
    I = 1
    Dim LeftPart
    Do While True
        I = InStr(1, Entire, Word)
        If I = 0 Then
            Exit Do
        Else
            LeftPart = Left(Entire, I - 1)
            Entire = LeftPart & Replace & Right(Entire, Len(Entire) - Len(Word) - Len(LeftPart))
        End If
    Loop
    
   RemoveString = Entire
      End Function

Private Sub txtName_Change()
SaveSetting "Galaxy Telnet Chat", "Settings", "RoomName", txtName.Text
End Sub

Private Sub txtPassword_Change()
SaveSetting "Galaxy Telnet Chat", "Settings", "RoomPassword", txtPassword.Text
End Sub
