VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pass As Boolean
Dim Command As String

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
Winsock1.LocalPort = 23
Winsock1.Listen
Label1.Caption = ""
Dir1.Path = "C:\"
End Sub

Private Sub Winsock1_Close()
Winsock1.Close
Do Until Winsock1.State = sckClosed
DoEvents
Loop
Winsock1.LocalPort = 23
Winsock1.Listen
Dir1.Path = "C:\"
Pass = False
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
Do Until Winsock1.State = 7
DoEvents
Loop
Me.Caption = Winsock1.RemoteHostIP
Winsock1.SendData "Password: "
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Winsock1.GetData Data
If Asc(Data) = 13 Then
    Label1.Caption = Command
    If Pass = False Then
        If Command = "your_password" Then Pass = True: Winsock1.SendData vbCrLf & "welcome" & vbCrLf: Winsock1.SendData "C:\>" Else Winsock1.SendData "Password incorect!" & vbCrLf: Winsock1.SendData "Password: "
    Else
        If LCase(Command) = "cd.." Then
            If Dir1.Path <> "C:\" Then Dir1.Path = ".."
            If Dir1.Path <> "C:\" Then Winsock1.SendData UCase(Dir1.Path) & "\>" Else Winsock1.SendData "C:\>"
            Command = ""
            Exit Sub
        End If
        If LCase(Command) = "cd." Then
            Dir1.Path = "."
            If Dir1.Path <> "C:\" Then Winsock1.SendData UCase(Dir1.Path) & "\>" Else Winsock1.SendData "C:\>"
            Command = ""
            Exit Sub
        End If
        If LCase(Command) = "dir" Then
            Dim Lenght As Integer
            For I = 0 To Dir1.ListCount - 1
            Winsock1.SendData Dir1.List(I) & "     <DIR>" & vbCrLf
            Next
            For I = O To File1.ListCount
            Winsock1.SendData File1.List(I) & vbCrLf
            Next
            If Dir1.Path <> "C:\" Then Winsock1.SendData UCase(Dir1.Path) & "\>" Else Winsock1.SendData "C:\>"
            Command = ""
            Exit Sub
        End If
        If LCase(Left(Command, 4)) = "view" Then
            U = Right(Command, Len(Command) - 5)
            On Error GoTo err1
            If Dir1.Path = "C:\" Then
            Open "C:\" & U For Input As #1
            Do Until EOF(1)
            Line Input #1, O
            Winsock1.SendData O & vbCrLf
            Loop
            Close #1
            Else
            Open Dir1.Path & "\" & U For Input As #1
            Do Until EOF(1)
            Line Input #1, O
            Winsock1.SendData O & vbCrLf
            Loop
            Close #1
            End If
            If Dir1.Path <> "C:\" Then Winsock1.SendData UCase(Dir1.Path) & "\>" Else Winsock1.SendData "C:\>"
            Command = ""
            Exit Sub
err1:
            Winsock1.SendData Err.Description & vbCrLf
            If Dir1.Path <> "C:\" Then Winsock1.SendData UCase(Dir1.Path) & "\>" Else Winsock1.SendData "C:\>"
            Command = ""
            Exit Sub
        End If
        If LCase(Left(Command, 2)) = "cd" And LCase(Left(Command, 3)) <> "cd." And LCase(Left(Command, 3)) <> "cd\" And Len(Command) > 3 Then
        U = Right(Command, Len(Command) - 3)
            On Error GoTo err1
            If Dir1.Path <> "C:\" Then Dir1.Path = Dir1.Path & "\" & U Else Dir1.Path = Dir1.Path & U
            If Dir1.Path <> "C:\" Then Winsock1.SendData UCase(Dir1.Path) & "\>" Else Winsock1.SendData "C:\>"
            Command = ""
            Exit Sub
        End If
        If LCase(Command) = "cd\" Then
            Dir1.Path = "C:\"
            If Dir1.Path <> "C:\" Then Winsock1.SendData UCase(Dir1.Path) & "\>" Else Winsock1.SendData "C:\>"
            Command = ""
            Exit Sub
        End If
        If LCase(Command) = "quit" Then
            Winsock1.SendData "Bye Bye!!" & vbCrLf
            Winsock1_Close
            Command = ""
            Exit Sub
        End If
        If LCase(Command) = "help" Then
            Open App.Path & "\help.txt" For Input As #1
            Do Until EOF(1)
            Line Input #1, E
            Winsock1.SendData E & vbCrLf
            Loop
            Close #1
            If Dir1.Path <> "C:\" Then Winsock1.SendData UCase(Dir1.Path) & "\>" Else Winsock1.SendData "C:\>"
            Command = ""
            Exit Sub
        End If
        Winsock1.SendData "Wrong Command!" & vbCrLf & "Type help for help" & vbCrLf
        If Dir1.Path <> "C:\" Then Winsock1.SendData UCase(Dir1.Path) & "\>" Else Winsock1.SendData "C:\>"
    End If
Command = ""
Else
Command = Command & Data
End If
End Sub
