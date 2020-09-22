VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Mail Reader"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDel 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheckMail 
      Caption         =   "&Check mailbox"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   5520
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtBody 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Text            =   "frmMain.frx":0000
      Top             =   2760
      Width           =   7095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Messages"
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   7095
      Begin ComctlLib.ListView lvMessages 
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Subject"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Password:"
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "txtPassword"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "User Name:"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "txtUserName"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remote Host:"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "txtHost"
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum POP3States
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_RETR
    POP3_DELE
    POP3_QUIT
End Enum

Private m_State         As POP3States

Private m_oMessage      As CMessage
Private m_colMessages   As New CMessages
'

Private Sub cmdCheckMail_Click()
    
    'Check the emptiness of all the text fields except for the txtBody
    For Each c In Controls
        If TypeOf c Is TextBox Then
            If Len(c.Text) = 0 Then
                MsgBox c.name & " can't be empty", vbCritical
                Exit Sub
            End If
        End If
    Next
    '
    'Change the value of current session state
    m_State = POP3_Connect
    Winsock1.Close
    Winsock1.LocalPort = 0
    Winsock1.connect txtHost, 110
    '
    'Close the socket in case it was opened while another session
    
    '
    'reset the value of the local port in order to let to the
    'Windows Sockets select the new one itself
    'It's necessary in order to prevent the "Address in use" error,
    'which can appear if the Winsock Control has already used while the 
    'previous session
    
    '
    'POP3 server waits for the connection request at the port 110.
    'According with that we want the Winsock Control to be connected to
    'the port number 110 of the server we have supplied in txtHost field
    

End Sub



Private Sub cmdDel_Click()
    Unload Me
End Sub

Private Sub lvMessages_ItemClick(ByVal Item As ComctlLib.ListItem)

    txtBody = m_colMessages(Item.Key).MessageBody
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim strData As String
    
    Static intMessages          As Integer 'the number of messages to be loaded
    Static intCurrentMessage    As Integer 'the counter of loaded messages
    Static strBuffer            As String  'the buffer of the loading message
    '
    'Save the received data into strData variable
    Winsock1.GetData strData
    Debug.Print strData
    
    If Left$(strData, 1) = "+" Or m_State = POP3_RETR Then
        'If the first character of the server's response is "+" then
        'server accepted the client's command and waits for the next one
        'If this symbol is "-" then here we can do nothing
        'and execution skips to the Else section of the code
        'The first symbol may differ from "+" or "-" if the received
        'data are the part of the message's body, i.e. when
        'm_State = POP3_RETR (the loading of the message state)
        Select Case m_State
            Case POP3_Connect
                '
                'Reset the number of messages
                intMessages = 0
                '
                'Change current state of session
                m_State = POP3_USER
                '
                'Send to the server the USER command with the parameter.
                'The parameter is the name of the mail box
                'Don't forget to add vbCrLf at the end of the each command!
                Winsock1.SendData "USER " & txtUserName & vbCrLf
                Debug.Print "USER " & txtUserName
                'Here is the end of Winsock1_DataArrival routine until the
                'next appearing of the DataArrival event. But next time this
                'section will be skipped and execution will start right after
                'the Case POP3_USER section.
            Case POP3_USER
                '
                'This part of the code runs in case of successful response to
                'the USER command.
                'Now we have to send to the server the user's password
                '
                'Change the state of the session
                m_State = POP3_PASS
                Winsock1.SendData "PASS " & txtPassword & vbCrLf
                Debug.Print "PASS " & txtPassword
            Case POP3_PASS
                '
                'The server answered positively to the process of the
                'identification and now we can send the STAT command. As a
                'response the server is going to return the number of
                'messages in the mail box and its size in octets
                '
                ' Change the state of the session
                m_State = POP3_STAT
                '
                'Send STAT command to know how many
                'messages in the mailbox
                Winsock1.SendData "STAT" & vbCrLf
                Debug.Print "STAT"
            Case POP3_STAT
                '
                'The server's response to the STAT command looks like this:
                '"+OK 0 0" (no messages at the mailbox) or "+OK 3 7564"
                '(there are messages). Evidently, the first of all we have to
                'find out the first numeric value that contains in the
                'server's response
                intMessages = CInt(Mid$(strData, 5, _
                              InStr(5, strData, " ") - 5))
                If intMessages > 0 Then
                    '
                    'Oops. There is something in the mailbox!
                    'Change the session state
                    m_State = POP3_RETR
                    '
                    'Increment the number of messages by one
                    intCurrentMessage = intCurrentMessage + 1
                    '
                    'and we're sending to the server the RETR command in
                    'order to retrieve the first message
                    Winsock1.SendData "RETR 1" & vbCrLf
                    Debug.Print "RETR 1"
                Else
                    'The mailbox is empty. Send the QUIT command to the
                    'server in order to close the session
                    m_State = POP3_QUIT
                    Winsock1.SendData "QUIT" & vbCrLf
                    Debug.Print "QUIT"
                    MsgBox "You have not mail.", vbInformation
                End If
            Case POP3_RETR
                'This code executes while the retrieving of the mail body
                'The size of the message could be quite big and the
                'DataArrival event may rise several time. All the received
                'data stores at the strBuffer variable:
                strBuffer = strBuffer & strData
                '
                'If case of presence of the point in the buffer it indicates
                'the end of the message (look at SMTP protocol)
                If InStr(1, strBuffer, vbLf & "." & vbCrLf) Then
                    '
                    'Done! The message has loaded
                    '
                    'Delete the first string-the server's response
                    strBuffer = Mid$(strBuffer, InStr(1, strBuffer, vbCrLf) + 2)
                    '
                    'Delete the last string. It contains only the "." symbol,
                    'which indicates the end of the message
                    strBuffer = Left$(strBuffer, Len(strBuffer) - 3)
                    '
                    'Add new message to m_colMessages collection
                    Set m_oMessage = New CMessage
                    m_oMessage.CreateFromText strBuffer
                    m_colMessages.Add m_oMessage, m_oMessage.MessageID
                    Set m_oMessage = Nothing
                    '
                    'Clear buffer for next message
                    strBuffer = ""
                    'Now we comparing the number of loaded messages with the
                    'one returned as a response to the STAT command
                    If intCurrentMessage = intMessages Then
                        'If these values are equal then all the messages
                        'have loaded. Now we can finish the session. Due to
                        'this reason we send the QUIT command to the server
                        m_State = POP3_QUIT
                        Winsock1.SendData "QUIT" & vbCrLf
                        Debug.Print "QUIT"
                    Else
                        'If these values aren't equal then there are
                        'remain messages. According with that
                        'we increment the messages' counter
                        intCurrentMessage = intCurrentMessage + 1
                        '
                        'Change current state of session
                        m_State = POP3_RETR
                        '
                        'Send RETR command to download next message
                        Winsock1.SendData "RETR " & _
                        CStr(intCurrentMessage) & vbCrLf
                        Debug.Print "RETR " & intCurrentMessage
                    End If
                End If
            Case POP3_QUIT
                'No matter what data we've received it's important
                'to close the connection with the mail server
                Winsock1.Close
                'Now we're calling the ListMessages routine in order to
                'fill out the ListView control with the messages we've          
                'downloaded
                Call ListMessages
        End Select
    Else
        'As you see, there is no sophisticated error
        'handling. We just close the socket and show the server's response
        'That's all. By the way even fully featured mail applications
        'do the same.
            Winsock1.Close
            MsgBox "POP3 Error: " & strData, _
            vbExclamation, "POP3 Error"
    End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    MsgBox "Winsock Error: #" & Number & vbCrLf & _
            Description
            
End Sub

Private Sub ListMessages()

    Dim oMes As CMessage
    Dim lvItem As ListItem
    
    For Each oMes In m_colMessages
        Set lvItem = lvMessages.ListItems.Add
        lvItem.Key = oMes.MessageID
        lvItem.Text = oMes.from
        lvItem.SubItems(1) = oMes.Subject
        lvItem.SubItems(2) = oMes.SendDate
        lvItem.SubItems(3) = oMes.size
    Next
    
End Sub
