VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   3885
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3630
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
      ForeColor       =   &H000080FF&
      Height          =   3855
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close Window"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'simple loading/saving of config file

Private Sub Form_Load()
Dim ConfigFile As String

    Open "irc.conf" For Input As #3
        Do While Not EOF(3)
            Line Input #3, ConfigFile
            Text1.Text = Text1.Text & ConfigFile & vbCrLf
        Loop
    Close #3

End Sub

Private Sub mnuSave_Click()
    Open "irc.conf" For Output As #2
        Print #2, Text1.Text
    Close #2
End Sub
