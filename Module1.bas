Attribute VB_Name = "Module1"
Public Enum POP3States
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_RETR
    POP3_DELE
    POP3_QUIT
End Enum

Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerProc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Global lngTimerID As Long
Global MsgStatus As String
Global strFilename As String
Global TotalMails As Long
Global m_State As POP3States
Global showForm As Boolean
Global Play As Boolean

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal lngSysTime As Long)
    Call checknewmail
End Sub

Public Sub Main()
   strFilename = App.Path & "\mail.wav"
   showForm = True
   Play = True
   'ChangeIcon Form2.picMail, MsgStatus
   lngTimerID = SetTimer(0, 0, Round(Form2.txtDelay * 60000, 0), AddressOf TimerProc)
End Sub
Public Sub checknewmail()
    DoEvents
    MsgStatus = "Checking New mails from " & Form2.txtHost.Text
    ChangeIcon Form2.picReadingMail(3), MsgStatus
    m_State = POP3_Connect
    Form2.Winsock1.Close
    Form2.Winsock1.LocalPort = 0
    Form2.Winsock1.Connect Form2.txtHost.Text, 110
End Sub
