VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "yxirc"
   ClientHeight    =   3645
   ClientLeft      =   4965
   ClientTop       =   3855
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   6255
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2520
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   3375
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   6255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'yxirc hope you find the follwing useful
'=========================================================================================

Dim i As Integer
Dim z As Integer
Dim words As Variant
Dim x As Long

Private Sub Form_Load()
'open conf file and check for server and nick lines
    Open "irc.conf" For Input As #1
        Do While Not EOF(1)
            Line Input #1, sLine
            
            If InStr(1, sLine, "nick") Then
                Nick = Mid(sLine, InStr(1, sLine, "=") + 1)
            End If
            
            If InStr(1, sLine, "server") Then
                Host = Mid(sLine, InStr(1, sLine, "=") + 1)
            End If
                
        Loop
    Close #1
    
    Port = 6667
    
    AbleToReceive = False
    Connecting = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
        Cancel = 1
    End If
End Sub

Private Sub Text1_Change()
    'make sure main text box always scrolls to new messages
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Text2.Text = "" Then
            GoTo Error:
        End If
        
        On Error GoTo Error:
            
        If Left(Text2.Text, 5) = "/join" Then
            Winsock1.SendData "JOIN " & Mid(Text2.Text, 6) & vbCrLf
            CurrentChannel = Mid(Text2.Text, 6)
            Text2.Text = ""
            Connecting = False
            AbleToReceive = True
            'new window - see module - calls new form for new channel
            Call NewWindow
            Exit Sub
        End If
        
        'private messaging
        If Left(Text2.Text, 4) = "/msg" Then
            TempMsgBuff = Right(Text2.Text, Len(Text2.Text) - 5)
            MsgVar1 = Mid(TempMsgBuff, InStr(1, TempMsgBuff, " "))
            MsgVar2 = Left(TempMsgBuff, Len(TempMsgBuff) - Len(MsgVar1))
            Winsock1.SendData "PRIVMSG " & MsgVar2 & " :" & MsgVar1 & vbCrLf
            Text2.Text = ""
        End If
        
        If Left(Text2.Text, 1) <> "/" Then
            Winsock1.SendData "PRIVMSG " & CurrentChannel & " :" & Text2.Text & vbCrLf
            Text1.Text = Text1.Text & "<" & Nick & ">" & Text2.Text & vbCrLf
            Text2.Text = ""
        End If
        
    End If
Error:
        Exit Sub
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String
    Dim Buffer As String
    
    Winsock1.GetData Data
    
    On Error GoTo Error:
    
    
    'the commented out part of the if statement below means the status form
    'will show the raw irc data - which is useful if you are adding more commends etc
    
    'If Connecting = True Then
        Text1.Text = Text1.Text & Data & vbCrLf
    'End If

    If Left(Data, 4) = "PING" Then
        
        Buffer = Mid(Data, InStr(1, Data, ":") + 1)
        Winsock1.SendData "PONG " & Buffer
        Winsock1.SendData vbCrLf
    
    End If
    
        If Left(Data, 13) = "NOTICE AUTH :" Then
            NOT_AUT = NOT_AUT + 1
        End If

    If NOT_AUT = 2 Then

        Winsock1.SendData "USER yxirc yxirc yxirc :yxirc"
        Winsock1.SendData vbCrLf
        Winsock1.SendData "NICK " & Nick
        Winsock1.SendData vbCrLf
        NOT_AUT = 0
        
    End If
    
    If InStr(1, Data, "PRIVMSG") Then
        
        If AbleToReceive = True Then
    
        
            'normal private message parsing
            a = InStr(1, Data, "PRIVMSG")
            temp = Mid(Data, a)
            Temp3 = 10 + Len(CurrentChannel)
            InboundMessage = Mid(temp, Temp3)
            
            'find out who posted the message
            Temp2 = InStr(1, Data, "!") - 2
            
            'get person's name in variable
            InboundMessageName = Mid(Data, 2, Temp2)
            
            TempMsgBuff2 = Mid(Data, InStr(1, Data, "#"))
            TempMsgBuff3 = Left(TempMsgBuff2, InStr(1, TempMsgBuff2, ":") - 2)
            
            
            GetChannel1 = Mid(Data, InStr(1, Data, "#"))
            'get channel2 finds the channel it came from
            GetChannel2 = Left(GetChannel1, InStr(1, GetChannel1, ":") - 2)
            
            'ok im proud of this part :P
            'basically each channel form's caption is the name of the channel which
            'makes them easy to identify. The part below searches all of the open forms
            'for a caption maching the name of the channel from the private message and
            'displays it in the approprate channel window
            
            For i = 1 To NumWindows
                If GetChannel2 = frm(i).Caption Then
                    frm(i).Text1.Text = frm(i).Text1.Text & "<" & InboundMessageName & ">" & InboundMessage
                End If
            Next i
    
        End If
    Else
    
    If InStr(1, Data, "353") Then
'get name list
'all of the msgvar's are just temporary strings for parsing the irc data
'same for the tempmsgbuff's

        MsgVar2 = Mid(Data, InStr(1, Data, "353"))
        MsgVar1 = Mid(MsgVar2, InStr(1, MsgVar2, ":") + 1)
        MsgVar3 = Left(MsgVar1, InStr(2, MsgVar1, ":") - 3)
        
        'split up nicks seperated by spaces
        words = Split(MsgVar3, " ")

            GetChannel1 = Mid(Data, InStr(1, Data, "353"))
            GetChannel2 = Mid(GetChannel1, InStr(1, GetChannel1, "#"))
            GetChannel3 = Left(GetChannel2, InStr(1, GetChannel2, ":") - 2)
            
'there was more data to parse for the nicks so i needed another getchannel variable.
'get channel3 is the name of the channel for the newly acquired nicks


'find form and add nicks to listbox
    
    For x = LBound(words) To UBound(words)
        For i = 1 To NumWindows
            If frm(i).Caption = GetChannel3 Then
                frm(i).List1.AddItem words(x)
            End If
        Next i
    Next x

End If

If InStr(1, Data, "PART") Then
'if someone leaves channel - this doesnt work yet but you get the idea from those above
        TempMsgBuff2 = Left(Data, InStr(1, Data, "!") - 1)
        LeavingNick = Mid(TempMsgBuff2, 2)
        
        GetChannel1 = Mid(Data, InStr(3, Data, "#") - 1)

        If LeavingNick <> Nick Then
        
            For i = 1 To NumWindows
                If GetChannel1 = frm(i).Caption Then
                    frm(i).Text1.Text = frm(i).Text1.Text & ">>> " & LeavingNick & " has left " & frm(i).Caption & vbCrLf
                End If
            Next i

        End If
End If

If InStr(1, Data, "QUIT") Then
    
        TempMsgBuff2 = Left(Data, InStr(1, Data, "!") - 1)
        LeavingNick = Mid(TempMsgBuff2, 2)
        
        GetChannel1 = Mid(Data, InStr(3, Data, ":") + 1)
        GetChannel2 = Left(GetChannel1, InStr(1, GetChannel1, vbCrLf) - 1)
        
        If LeavingNick <> Nick Then
        
               For i = 1 To NumWindows
                If GetChannel2 = frm(i).Caption Then
                    frm(i).Text1.Text = frm(i).Text1.Text & ">>> " & LeavingNick & " has Quit(Connection reset by peer)" & vbCrLf
                End If
            Next i
        End If
End If

If InStr(1, Data, "JOIN") Then
        'this works :P same ideas as above and says when someone has joined the channel
        TempMsgBuff2 = Left(Data, InStr(1, Data, "!") - 1)
        JoiningNick = Mid(TempMsgBuff2, 2)
        

        GetChannel1 = Mid(Data, InStr(3, Data, ":") + 1)
        GetChannel2 = Left(GetChannel1, InStr(1, GetChannel1, vbCrLf) - 1)
            
            
        If JoiningNick = Nick Then
            Exit Sub
        End If
    
            For z = 1 To NumWindows
                If GetChannel2 = frm(z).Caption Then
                    frm(z).Text1.Text = frm(z).Text1.Text & ">>> " & JoiningNick & " has joined" & vbCrLf
                    frm(z).List1.AddItem JoiningNick
                End If
            Next z
End If

End If
Error:
End Sub

