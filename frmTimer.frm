VERSION 5.00
Object = "{08E3F9AA-39A3-4CF1-A497-259D87D3790D}#1.0#0"; "COMPCONTROLS.OCX"
Begin VB.Form frmTimer 
   Caption         =   "Timer"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   6510
   Icon            =   "frmTimer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1800
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "Sto&p"
      Height          =   600
      Left            =   5280
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5160
      Top             =   1800
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   615
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtSec 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3600
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtHour 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   1575
   End
   Begin CompControler.CompControl CompControl1 
      Left            =   4560
      Top             =   1800
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileConfigure 
         Caption         =   "&Configure"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuFileChoice 
         Caption         =   "Choice00"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFileChoice 
         Caption         =   "Choice01"
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFileChoice 
         Caption         =   "Choice02"
         Index           =   2
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cSec As Integer
Dim cMin As Integer
Dim cHour As Integer
Private dSec As String
Private dMin As String
Private dHour As String
Private Finished As Boolean
Private TimeStart As String
Private TimeStop As String
Private PathIni As String
Private setLabel(2) As String
Private setHour(2) As String
Private setMin(2) As String
Private setSec(2) As String
Private CatChoice(2) As String

Private Sub cmdStart_Click()
txtSec.Locked = True
txtMin.Locked = True
txtHour.Locked = True
cSec = CInt(txtSec.Text)
cMin = CInt(txtMin.Text)
cHour = CInt(txtHour.Text)
PlaySound (TimeStart)

Timer1.Enabled = True

End Sub

Private Sub cmdStop_Click()
txtSec.Locked = False
txtMin.Locked = False
txtHour.Locked = False
Timer1.Enabled = False

End Sub

Private Sub Form_Load()
CheckINI

TimeStart = Dir1.Path & "\TimeStart.wav"
TimeStop = Dir1.Path & "\TimeStop.wav"

End Sub

Private Sub Form_Unload(Cancel As Integer)
'perform writeini herel...

End Sub

Private Sub mnuFileChoice_Click(Index As Integer)
txtHour.Text = setHour(Index)
txtMin.Text = setMin(Index)
txtSec.Text = setSec(Index)
frmTimer.SetFocus

Finished = False
cmdStart_Click

End Sub

Private Sub mnuFileConfigure_Click()
cmdStop_Click
frmTimer.Enabled = False
frmConfig.Show

End Sub

Private Sub mnuFileExit_Click()
Unload Me

End Sub

Private Sub Timer1_Timer()
If Finished = False Then
    CountDown
Else
    CountUp
End If

End Sub

Private Sub CountDown()
If cHour > 0 Then
    If cMin > 0 Then
        If cSec > 0 Then
            cSec = cSec - 1
            If cSec < 10 Then
                dSec = "0" & cSec
            Else
                dSec = cSec
            End If
            If cMin < 10 Then
                dMin = "0" & cMin
            Else
                dMin = cMin
            End If
            If cHour < 10 Then
                dHour = "0" & cHour
            Else
                dHour = cHour
            End If
            Me.Caption = "Timer " & dHour & ":" & dMin & ":" & dSec
            txtSec.Text = dSec
            txtMin.Text = dMin
            txtHour.Text = dHour
            Exit Sub
        End If
        If cSec = 0 Then
            cSec = 59
            cMin = cMin - 1
            dSec = cSec
            If cSec < 10 Then
                dSec = "0" & cSec
            Else
                dSec = cSec
            End If
            If cMin < 10 Then
                dMin = "0" & cMin
            Else
                dMin = cMin
            End If
            If cHour < 10 Then
                dHour = "0" & cHour
            Else
                dHour = cHour
            End If
            Me.Caption = "Timer " & dHour & ":" & dMin & ":" & dSec
            txtSec.Text = dSec
            txtMin.Text = dMin
            txtHour.Text = dHour
            Exit Sub
        End If
    End If
    If cMin = 0 Then
        dMin = "0" & cMin
        If cSec > 0 Then
            cSec = cSec - 1
            If cSec < 10 Then
                dSec = "0" & cSec
            Else
                dSec = cSec
            End If
            If cMin < 10 Then
                dMin = "0" & cMin
            Else
                dMin = cMin
            End If
            If cHour < 10 Then
                dHour = "0" & cHour
            Else
                dHour = cHour
            End If
            Me.Caption = "Timer " & dHour & ":" & dMin & ":" & dSec
            txtSec.Text = dSec
            txtMin.Text = dMin
            txtHour.Text = dHour
            Exit Sub
        End If
        If cSec = 0 Then
            cSec = 59
            cMin = 59
            cHour = cHour - 1
            If cSec < 10 Then
                dSec = "0" & cSec
            Else
                dSec = cSec
            End If
            If cMin < 10 Then
                dMin = "0" & cMin
            Else
                dMin = cMin
            End If
            If cHour < 10 Then
                dHour = "0" & cHour
            Else
                dHour = cHour
            End If
            Me.Caption = "Timer " & dHour & ":" & dMin & ":" & dSec
            txtSec.Text = dSec
            txtMin.Text = dMin
            txtHour.Text = dHour
            Exit Sub
        End If
    End If
End If
If cHour = 0 Then
    If cMin > 0 Then
        If cSec > 0 Then
            cSec = cSec - 1
            If cSec < 10 Then
                dSec = "0" & cSec
            Else
                dSec = cSec
            End If
            If cMin < 10 Then
                dMin = "0" & cMin
            Else
                dMin = cMin
            End If
            If cHour < 10 Then
                dHour = "0" & cHour
            Else
                dHour = cHour
            End If
            Me.Caption = "Timer " & dHour & ":" & dMin & ":" & dSec
            txtSec.Text = dSec
            txtMin.Text = dMin
            txtHour.Text = dHour
            Exit Sub
        End If
        If cSec = 0 Then
            cSec = 59
            cMin = cMin - 1
            If cSec < 10 Then
                dSec = "0" & cSec
            Else
                dSec = cSec
            End If
            If cMin < 10 Then
                dMin = "0" & cMin
            Else
                dMin = cMin
            End If
            If cHour < 10 Then
                dHour = "0" & cHour
            Else
                dHour = cHour
            End If
            Me.Caption = "Timer " & dHour & ":" & dMin & ":" & dSec
            txtSec.Text = dSec
            txtMin.Text = dMin
            txtHour.Text = dHour
            Exit Sub
        End If
    End If
    If cMin = 0 Then
        If cSec > 0 Then
            cSec = cSec - 1
            If cSec < 10 Then
                dSec = "0" & cSec
            Else
                dSec = cSec
            End If
            If cMin < 10 Then
                dMin = "0" & cMin
            Else
                dMin = cMin
            End If
            If cHour < 10 Then
                dHour = "0" & cHour
            Else
                dHour = cHour
            End If
            Me.Caption = "Timer " & dHour & ":" & dMin & ":" & dSec
            txtSec.Text = dSec
            txtMin.Text = dMin
            txtHour.Text = dHour
            Exit Sub
        End If
        If cSec = 0 Then
            Me.Caption = "Timer " & dHour & ":" & dMin & ":" & dSec
            PlaySound (TimeStop)
            frmTimeIsUp.Show
            Finished = True
            txtSec.Text = dSec
            txtMin.Text = dMin
            txtHour.Text = dHour
            Exit Sub
        End If
    End If
End If

End Sub

Private Sub CountUp()
If cSec < 59 Then
    cSec = cSec + 1
    If cSec < 10 Then
        dSec = "0" & cSec
    Else
        dSec = cSec
    End If
    If cMin < 10 Then
        dMin = "0" & cMin
    Else
        dMin = cMin
    End If
    If cHour < 10 Then
        dHour = "0" & cHour
    Else
        dHour = cHour
    End If
            
Else
    cSec = 0
    dSec = "00"
    If cMin < 59 Then
        cMin = cMin + 1
        If Finished = True Then
            PlaySound (TimeStop)
        End If
        If cSec < 10 Then
            dSec = "0" & cSec
        Else
            dSec = cSec
        End If
        If cMin < 10 Then
            dMin = "0" & cMin
        Else
            dMin = cMin
        End If
        If cHour < 10 Then
            dHour = "0" & cHour
        Else
            dHour = cHour
        End If
    Else
        cMin = 0
        dMin = "00"
        cHour = cHour + 1
        If cSec < 10 Then
            dSec = "0" & cSec
        Else
            dSec = cSec
        End If
        If cMin < 10 Then
            dMin = "0" & cMin
        Else
            dMin = cMin
        End If
        If cHour < 10 Then
            dHour = "0" & cHour
        Else
            dHour = cHour
        End If
    End If
End If
txtSec.Text = dSec
txtMin.Text = dMin
txtHour.Text = dHour
Me.Caption = "Timer " & dHour & ":" & dMin & ":" & dSec
End Sub

Private Sub txtHour_GotFocus()
txtHour.SelStart = 0
txtHour.SelLength = Len(txtHour.Text)

End Sub

Private Sub txtHour_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then    'Backspace key
    KeyAscii = 8
ElseIf IsNumeric(Chr(KeyAscii)) = False Or KeyAscii = 8 Then
    KeyAscii = 0
End If

End Sub

Private Sub txtMin_GotFocus()
txtMin.SelStart = 0
txtMin.SelLength = Len(txtMin.Text)

End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then    'Backspace key
    KeyAscii = 8
ElseIf IsNumeric(Chr(KeyAscii)) = False Then
    KeyAscii = 0
End If

End Sub

Private Sub txtSec_GotFocus()
txtSec.SelStart = 0
txtSec.SelLength = Len(txtSec.Text)

End Sub

Private Sub txtSec_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then    'Backspace key
    KeyAscii = 8
ElseIf IsNumeric(Chr(KeyAscii)) = False Then
    KeyAscii = 0
End If

End Sub

Public Sub CheckINI()
Dim i As Integer
PathIni = Dir1.Path & "\Timer.ini"
CatChoice(0) = "mnuFileChoice00"
CatChoice(1) = "mnuFileChoice01"
CatChoice(2) = "mnuFileChoice02"

For i = 0 To 2
    setLabel(i) = Readini(PathIni, CatChoice(i), "Label")
    mnuFileChoice(i).Caption = setLabel(i)
    setHour(i) = Readini(PathIni, CatChoice(i), "Hour")
    setMin(i) = Readini(PathIni, CatChoice(i), "Minute")
    setSec(i) = Readini(PathIni, CatChoice(i), "Second")
Next i

End Sub

Private Sub PlaySound(ByVal WhatSound As String)
If frmConfig.chkPlaySound = 1 Then CompControl1.PlayWAVFile (WhatSound)

End Sub

