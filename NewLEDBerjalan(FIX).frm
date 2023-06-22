VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   14400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mulai"
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3120
      Top             =   2760
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   9015
      Begin VB.Shape Shape5 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   4080
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3120
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2160
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1200
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   240
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Zu Dzacky Haikal"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Simulasi LED Berjalan Dengan Visual Basic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   8895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Shape1.FillColor = vbWhite
Shape2.FillColor = vbWhite
Shape3.FillColor = vbWhite
Shape4.FillColor = vbWhite
Shape5.FillColor = vbWhite
End Sub

Private Sub Timer1_Timer()
Dim a As Integer
a = Timer1.Interval
Timer1.Interval = a + 100
Select Case a
Case 100
Shape5.FillColor = vbWhite
Shape1.FillColor = vbRed
Case 200
Shape1.FillColor = vbWhite
Shape2.FillColor = vbRed
Case 300
Shape2.FillColor = vbWhite
Shape3.FillColor = vbRed
Case 400
Shape3.FillColor = vbWhite
Shape4.FillColor = vbRed
Case 500
Shape4.FillColor = vbWhite
Shape5.FillColor = vbRed
Timer1.Interval = 100
End Select
Label1.Caption = Right(Label1.Caption, Len(Label1.Caption) - 1) & Left(Label1.Caption, 1)
End Sub
