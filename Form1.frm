VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyIRC Server"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckServer2 
      Left            =   2640
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1770
   End
   Begin VB.Timer tmBeforeDisc 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5640
      Top             =   3120
   End
   Begin MSWinsockLib.Winsock sckUsers 
      Index           =   0
      Left            =   2160
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Left            =   1680
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1768
   End
   Begin MSComDlg.CommonDialog opensave 
      Left            =   3120
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Open list"
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save list"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Unban account"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ban account"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete account"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add account"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
   Begin MSComctlLib.ListView lstUsers 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nickname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Banned"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Contacts"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private user(100) As String

Private Sub Command1_Click()
Dim info As String
'Loop over each column in the listview
For i = 1 To lstUsers.ColumnHeaders.Count
'Ask what to fill in, for example username
info = InputBox("What to fill in at " & lstUsers.ColumnHeaders(i).Text & "?")
If i = 1 Then
'If it's the first item, create it
lstUsers.ListItems.Add , , info
Else
'Otherwise, add subitems to it
lstUsers.ListItems.Item(lstUsers.ListItems.Count).ListSubItems.Add , , info
End If
Next
End Sub

Private Sub Command2_Click()
'Delete the selected item, if there is nothing selected, don't shut down the program
On Error Resume Next
lstUsers.ListItems.Remove lstUsers.SelectedItem.Index
End Sub

Private Sub Command3_Click()
'Change the fifth subitem of the selected item
On Error Resume Next
lstUsers.SelectedItem.ListSubItems.Item(5).Text = "True"
End Sub

Private Sub Command4_Click()
'Change the fifth subitem of the selected item again
On Error Resume Next
lstUsers.SelectedItem.ListSubItems.Item(5).Text = "False"
End Sub

Private Sub Command5_Click()
'Remove the stored filename
opensave.FileName = ""
'Set the filter, so that it filters .mil files
opensave.Filter = "MyIRC Userlists (*.mil)|*.mil"
'Show the save window
opensave.ShowSave
'Check if a file was selected
If opensave.FileName <> "" Then
'Collect the items in the ListView
Dim items, infos As String
For i = 1 To lstUsers.ListItems.Count
infos = ""
For ii = 1 To lstUsers.ListItems.Item(i).ListSubItems.Count
infos = infos & lstUsers.ListItems.Item(i).ListSubItems.Item(ii).Text & "|"
Next
infos = lstUsers.ListItems.Item(i).Text & "|" & infos
items = items & infos & "*"
Next
'Save the file
Dim hFile As Long
hFile = FreeFile
Open opensave.FileName For Output As #hFile
Print #hFile, items
Close #hFile
End If
End Sub

Private Sub Command6_Click()
'Dialog thingy again
opensave.FileName = ""
opensave.Filter = "MyIRC Userlists (*.mil)|*.mil"
opensave.ShowOpen

If opensave.FileName <> "" Then
'Clear the listview
lstUsers.ListItems.Clear

'Open the file
Dim content As String
fnum = FreeFile
Open opensave.FileName For Input As fnum
content = Input$(LOF(fnum), #fnum)
Close fnum
'Now the file is stored into "content"
'Now process it:
Dim items() As String
Dim headers() As String
'Split the items into an array
items = Split(content, "*")
'Now loop over all items
For i = 0 To UBound(items) - 1
'Get the headers (username, password, etc)
headers = Split(items(i), "|")
'Loop over the headers
For ii = 0 To UBound(headers) - 1
If ii = 0 Then
'Create the item
lstUsers.ListItems.Add , , headers(ii)
Else
'Or at the subitems
lstUsers.ListItems.Item(lstUsers.ListItems.Count).ListSubItems.Add , , headers(ii)
End If
Next
Next
End If
End Sub

Private Sub Form_Load()
'Let the server socket wait for connections
sckServer.Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
sckServer.Close
End Sub

Private Sub sckServer_ConnectionRequest(ByVal requestID As Long)
'Close the socket and accept the connection
sckServer.Close
sckServer.Accept requestID
End Sub

Private Sub sckServer_DataArrival(ByVal bytesTotal As Long)
'Put the received data in a string
Dim data As String
sckServer.GetData data
'Process it
Dim info() As String
info = Split(data, "|")
'The client has to send data in this format:
'username|password|version
'So the server can check the authentication, send the contactlist and check for updates
If info(2) < 1 Then
'Not the latest version, send info. With ERR we let the client know, something went wrong.
sckServer.SendData "ERR You don't have the latest version!"
Exit Sub
Else
'Version is good, let's look for the login data
Dim indx As Long
For i = 1 To lstUsers.ListItems.Count
If lstUsers.ListItems.Item(i).ListSubItems.Item(1) = info(0) Then
indx = i
Exit For
End If
Next

If indx = 0 Then
'Account not found, let the user know...
sckServer.SendData "ERR The account doesn't exist!"
Exit Sub
Else
'Found the user, check the password
If lstUsers.ListItems.Item(indx).ListSubItems.Item(2) = info(1) Then
'Password is correct
'The user will be forwarded to an user socket
sckServer.SendData "OK " & Len(info(0)) * Len(info(1)) + Len(info(1))
Exit Sub
Else
'Password is wrong
sckServer.SendData "ERR Wrong password!"
Exit Sub
End If
End If

End If
End Sub

Private Sub sckServer_SendComplete()
sckServer.Close
sckServer.Listen
End Sub

Private Sub sckServer2_ConnectionRequest(ByVal requestID As Long)
'Look for an empty socket
Dim indx As Long
For i = 1 To 100
'Found something!
If sckUsers(i).State = 0 Then
indx = i
Exit For
End If
Next

'Now forward the client to the user socket
sckUsers(i).Accept requestID
sckServer2.Close
sckServer2.Listen 'Again

'After the clients forwarding, it has to send the login information, we'll see that
'In the DataArrival sub!
End Sub

Private Sub sckUsers_Close(Index As Integer)
'Make sure it's closed
sckUsers(Index).Close
'Free up memory
Unload sckUsers(Index)
'Now let other people who have this person in their list know he went offline
For i = 1 To lstUsers.ListItems.Count
'Check if the person is not offline
If lstUsers.ListItems.Item(i).ListSubItems.Item(3) <> "0" Then
'Now check is this people know each other
If InStr(lstUsers.ListItems.Item(i).ListSubItems.Item(6), user(Index)) Then
'Search the socket, the user is connected too
For ii = 0 To UBound(user)
If user(ii) = lstUsers.ListItems.Item(i).ListSubItems.Item(1) Then
sckUsers(ii).SendData "OFF " & user(Index)
Exit Sub
End If
Next
End If
End If
Next
End Sub

Private Sub tmBeforeDisc_Timer()
'Send a time out message to connected the client
sckServer.SendData "TIME OUT"
'Disconnect and listen again
sckServer.Close
sckServer.Listen
End Sub
