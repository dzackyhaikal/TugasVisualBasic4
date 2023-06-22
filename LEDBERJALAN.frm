VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mulai"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   2280
   End
   Begin VB.Frame Frame1 
      Caption         =   "LED"
      Height          =   975
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   5415
      Begin VB.Shape Shape4 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   3480
         Shape           =   3  'Circle
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   1
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   0
         Left            =   4440
         Shape           =   3  'Circle
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   0
         Left            =   360
         Shape           =   3  'Circle
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Simulasi LED Berjalan Dengan Visual Basic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   5895
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
Shape5.Item.FillColor = vbWhite
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
