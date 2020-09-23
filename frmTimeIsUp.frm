VERSION 5.00
Begin VB.Form frmTimeIsUp 
   Caption         =   "TIME IS UP!"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "frmTimeIsUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TIME IS UP!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmTimeIsUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

Unload Me

End Sub

