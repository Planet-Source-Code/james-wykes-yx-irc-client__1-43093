Attribute VB_Name = "Module1"

'here is all the variables, there are probably a load of unnessesary temporary ones for parsing
'the data, but what the hey, it works

Public Host As String
Public Port As Integer

Public Nick As String
Public Username As String
Public NOT_AUT As Long
Public CurrentChannel As String
Public a As Integer  'tepmorary incoming msg buffer
Public b As Integer
Public TempMsgBuff As String 'temporary /msg buffer
Public TempMsgBuff2 As String 'temporary buffer
Public TempMsgBuff3 As String 'temporary buffer
Public temp As String
Public Temp2 As String
Public Temp3 As String
Public MsgVar1 As String
Public MsgVar2 As String
Public MsgVar3 As String
Public InboundMessage As String
Public InboundMessageName As String
Public AbleToReceive As Boolean
Public Connecting As Boolean
Public strNicks As String
Public NickData As String
Public JoiningChannel As Boolean
Public iInt As Integer
Public JoiningNick As String
Public LeavingNick As String

Public GetChannel1 As String
Public GetChannel2 As String
Public GetChannel3 As String

Public frm(1000) As Form2
Public NumWindows As Integer

Public ConnectionTime As Integer

Public sLine As String

Sub NewWindow()
    NumWindows = NumWindows + 1
    Set frm(NumWindows) = New Form2
    frm(NumWindows).Show
    frm(NumWindows).Caption = Mid(CurrentChannel, 2)
    frm(NumWindows).Text1.Text = ">>> Chatting in " & frm(NumWindows).Caption & vbCrLf
End Sub


