VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnHitung 
      Caption         =   "hitung"
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtNilaiX 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Y ="
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "X ="
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblHasil 
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnHitung_Click()
Dim x As Double
x = CDbl(txtNilaiX.Text)

Dim y As Double
    y = x ^ 2 + 3 * x + 2
    
lblHasil.Caption = CStr(y)

End Sub

