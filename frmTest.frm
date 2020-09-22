VERSION 5.00
Object = "*\AEasyButton.vbp"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EasyButton Test"
   ClientHeight    =   2325
   ClientLeft      =   2625
   ClientTop       =   1980
   ClientWidth     =   4290
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTest.frx":0000
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   ShowInTaskbar   =   0   'False
   Begin EasyX.EasyButton EasyButton5 
      Height          =   330
      Left            =   2970
      TabIndex        =   6
      Top             =   1440
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Caption         =   "&More..."
      Align           =   0
      BackColor       =   11451596
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   2701148
      AccessKey       =   "M"
   End
   Begin EasyX.EasyButton EasyButton1 
      Height          =   330
      Left            =   135
      TabIndex        =   2
      Top             =   630
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Caption         =   "&Picture"
      Align           =   0
      Picture         =   "frmTest.frx":5C04
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   2701148
      AccessKey       =   "P"
   End
   Begin EasyX.EasyButton EasyButton2 
      Height          =   330
      Left            =   135
      TabIndex        =   3
      Top             =   990
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Caption         =   "&No Picture"
      Align           =   0
      Picture         =   "frmTest.frx":7BD4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   2701148
      AccessKey       =   "N"
   End
   Begin EasyX.EasyButton EasyButton3 
      Height          =   330
      Left            =   135
      TabIndex        =   4
      Top             =   1350
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Caption         =   "&About"
      Align           =   0
      Picture         =   "frmTest.frx":9BA4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   2701148
      AccessKey       =   "A"
   End
   Begin EasyX.EasyButton EasyButton4 
      Height          =   330
      Left            =   135
      TabIndex        =   5
      Top             =   1890
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Caption         =   "&Exit"
      Align           =   0
      Picture         =   "frmTest.frx":BB74
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   2701148
      AccessKey       =   "E"
   End
   Begin EasyX.EasyButton EasyButton6 
      Height          =   330
      Left            =   2970
      TabIndex        =   7
      Top             =   1845
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Caption         =   "&How To"
      Align           =   0
      BackColor       =   11451596
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   2701148
      AccessKey       =   "H"
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "For best performance, it is a good idea to compile the UserControl like OCX."
      ForeColor       =   &H00400040&
      Height          =   555
      Left            =   1665
      TabIndex        =   1
      Top             =   675
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00D2D7DD&
      Index           =   1
      X1              =   0
      X2              =   288
      Y1              =   34
      Y2              =   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0029375C&
      Index           =   0
      X1              =   0
      X2              =   288
      Y1              =   33
      Y2              =   33
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Easy Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0029375C&
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EasyButton1_Click()

    Form3.Show 0

End Sub
Private Sub EasyButton2_Click()

    Form2.Show 1

End Sub
Private Sub EasyButton3_Click()

    EasyButton3.AboutBox

End Sub
Private Sub EasyButton4_Click()

    Unload Me
    End

End Sub
Private Sub EasyButton5_Click()

    Form4.Show 1

End Sub
Private Sub EasyButton6_Click()

    Form5.Show 1

End Sub
Private Sub Form_Resize()

    For X = 1 To ScaleWidth Step 100
        For Y = 1 To ScaleHeight Step 100
            PaintPicture Picture, X, Y
        Next
    Next

End Sub
