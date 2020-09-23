VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   6240
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
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
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   0
      Width           =   4815
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000008&
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
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   1455
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is the same as the form1 code

Private Sub Form_Unload(Cancel As Integer)
    Form1.Winsock1.SendData "PART " & Me.Caption & vbCrLf
    NumWindows = NumWindows - 1
End Sub

Private Sub Text1_Change()
    Text1.SelStart = Len(Text1)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Text2.Text = "" Then
            Exit Sub
        End If
        
        On Error GoTo Error:
            
        If Left(Text2.Text, 5) = "/join" Then
            Form1.Winsock1.SendData "JOIN " & Mid(Text2.Text, 6) & vbCrLf
            CurrentChannel = Mid(Text2.Text, 6)
            'Text1.Text = ">>> Chatting in:" & CurrentChannel & " as: " & Nick & vbCrLf
            Text2.Text = ""
            Connecting = False
            Call NewWindow
            Exit Sub
        End If
        
        If Left(Text2.Text, 4) = "/msg" Then
            TempMsgBuff = Right(Text2.Text, Len(Text2.Text) - 5)
            MsgVar1 = Mid(TempMsgBuff, InStr(1, TempMsgBuff, " "))
            MsgVar2 = Left(TempMsgBuff, Len(TempMsgBuff) - Len(MsgVar1))
            Form1.Winsock1.SendData "PRIVMSG " & MsgVar2 & " :" & MsgVar1 & vbCrLf
            Text2.Text = ""
        End If
        
        If Left(Text2.Text, 1) <> "/" Then
            Form1.Winsock1.SendData "PRIVMSG " & Me.Caption & " :" & Text2.Text & vbCrLf
            Text1.Text = Text1.Text & "<" & Nick & ">" & Text2.Text & vbCrLf
            Text2.Text = ""
        End If
        
    End If
Error:
        Exit Sub
End Sub
