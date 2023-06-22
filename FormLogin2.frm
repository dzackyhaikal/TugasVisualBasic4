VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Zu Dzacky Haikal"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Form Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   675
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   4980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Sub Command1_Click()
 
 If Text1 = "admin" And Text2 = "admin" Then
Form2.Show 'Perintah Menampilkan Form 2
 Form1.Visible = False 'Menyembunyikan Form 1
 Unload Me 'Menutup Form 1
6  Else
MsgBox "User Name atau Password yang Anda Masukkan salah" _
& vbNewLine & "Silahkan Coba lagi !!", vbCritical, "Warning!!"
 Text1 = ""
Text2 = ""
 Text1.SetFocus
 End If
 End Sub
 Private Sub Command2_Click()
 Unload Me
 End Sub

