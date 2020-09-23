VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure..."
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4845
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPlaySound 
      Caption         =   "Sound Enabled --- Uncheck this to disable sound"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.TextBox txtLabel 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtHour 
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   9
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   10
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox txtSec 
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   11
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox txtLabel 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtHour 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtSec 
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdKeep 
      Caption         =   "Keep"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtSec 
      Height          =   285
      Index           =   0
      Left            =   4080
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtHour 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtLabel 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   3360
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Name Choice 3 ------->"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   27
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Hours"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   26
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Minutes"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   25
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Seconds"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   24
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name Choice 2 ------->"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Hours"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Minutes"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   21
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Seconds"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   20
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Seconds"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   18
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Minutes"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   17
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Hours"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Name Choice 1 ------->"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PathIni As String
Private setLabel(2) As String
Private setHour(2) As String
Private setMin(2) As String
Private setSec(2) As String
Private CatChoice(3) As String
Public PlaySound As Boolean

Private Sub chkPlaySound_Click()
If chkPlaySound.Value = 0 Then
    PlaySound = False
    chkPlaySound.Caption = "Sound Disabled --- Check to play sound"
Else
    PlaySound = True
    chkPlaySound.Caption = "Sound Enabled --- Uncheck if problems with sound"
End If

End Sub

Private Sub cmdCancel_Click()
frmTimer.Enabled = True
Unload Me

End Sub

Private Sub cmdKeep_Click()
Dim i As Integer
For i = 0 To 2
    setLabel(i) = txtLabel(i).Text
    Debug.Print "setLabel(i) = " & setLabel(i)
    setHour(i) = txtHour(i).Text
    Debug.Print "sethour(i) = " & setHour(i)
    setMin(i) = txtMin(i).Text
    Debug.Print "setmin(i) = " & setMin(i)
    setSec(i) = txtSec(i).Text
    Debug.Print "setSec(i) = " & setSec(i)
    
Next i

i = 0
For i = 0 To 2
    Writeini PathIni, CatChoice(i), "Label", setLabel(i)
    frmTimer.mnuFileChoice(i).Caption = setLabel(i)
    Writeini PathIni, CatChoice(i), "Hour", setHour(i)
    Writeini PathIni, CatChoice(i), "Minute", setMin(i)
    Writeini PathIni, CatChoice(i), "Second", setSec(i)
Next i
Dim plysndTrue As String
Dim plysndFalse As String

Writeini PathIni, CatChoice(3), "PlaySound", PlaySound

frmTimer.CheckINI
cmdCancel_Click


End Sub

Private Sub cmdReset_Click()
Dim i As Integer
For i = 0 To 2
txtLabel(i).Text = setLabel(i)
txtHour(i).Text = setHour(i)
txtMin(i).Text = setMin(i)
txtSec(i).Text = setSec(i)

Next i

End Sub

Private Sub Form_Load()
Dim i As Integer
PathIni = Dir1.Path & "\Timer.ini"
CatChoice(0) = "mnuFileChoice00"
CatChoice(1) = "mnuFileChoice01"
CatChoice(2) = "mnuFileChoice02"
CatChoice(3) = "PlaySound"
For i = 0 To 2
    setLabel(i) = Readini(PathIni, CatChoice(i), "Label")
    setHour(i) = Readini(PathIni, CatChoice(i), "Hour")
    setMin(i) = Readini(PathIni, CatChoice(i), "Minute")
    setSec(i) = Readini(PathIni, CatChoice(i), "Second")
Next i
PlaySound = Readini(PathIni, CatChoice(3), "PlaySound")
If PlaySound = True Then
    chkPlaySound.Value = 1
    chkPlaySound.Caption = "Sound Enabled --- Uncheck if problems with sound"
Else
    chkPlaySound.Value = 0
    chkPlaySound.Caption = "Sound Disabled --- Check to play sound"
End If

cmdReset_Click

End Sub

Private Sub txtHour_GotFocus(Index As Integer)
txtHour(Index).SelStart = 0
txtHour(Index).SelLength = Len(txtHour(Index).Text)

End Sub

Private Sub txtHour_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then    'Backspace key
    KeyAscii = 8
ElseIf IsNumeric(Chr(KeyAscii)) = False Or KeyAscii = 8 Then
    KeyAscii = 0
End If

End Sub

Private Sub txtLabel_GotFocus(Index As Integer)
txtLabel(Index).SelStart = 0
txtLabel(Index).SelLength = Len(txtLabel(Index).Text)

End Sub

Private Sub txtMin_GotFocus(Index As Integer)
txtMin(Index).SelStart = 0
txtMin(Index).SelLength = Len(txtMin(Index).Text)

End Sub

Private Sub txtMin_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then    'Backspace key
    KeyAscii = 8
ElseIf IsNumeric(Chr(KeyAscii)) = False Or KeyAscii = 8 Then
    KeyAscii = 0
End If

End Sub

Private Sub txtSec_GotFocus(Index As Integer)
txtSec(Index).SelStart = 0
txtSec(Index).SelLength = Len(txtSec(Index).Text)

End Sub

Private Sub txtSec_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then    'Backspace key
    KeyAscii = 8
ElseIf IsNumeric(Chr(KeyAscii)) = False Or KeyAscii = 8 Then
    KeyAscii = 0
End If

End Sub
