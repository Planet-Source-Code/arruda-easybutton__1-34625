VERSION 5.00
Begin VB.Form frmAboutBox 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About EasyButton"
   ClientHeight    =   1980
   ClientLeft      =   2355
   ClientTop       =   1620
   ClientWidth     =   3210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1980
      TabIndex        =   1
      Top             =   1530
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EasyButton"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   765
      TabIndex        =   2
      Top             =   135
      Width           =   2310
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmAboutBox.frx":0000
      Top             =   135
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   -90
      X2              =   3735
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   -90
      X2              =   3735
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by:    "
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   1305
      TabIndex        =   0
      Top             =   540
      Width           =   1860
   End
End
Attribute VB_Name = "frmAboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub CenterForm(Frm)
        
    Frm.Top = (Screen.Height / 2) - (Frm.Height / 2)
    Frm.Left = (Screen.Width / 2) - (Frm.Width / 2)

End Sub
Private Sub Command1_Click()
    
    Unload frmAboutBox
    Set frmAboutBox = Nothing

End Sub

Private Sub Form_Load()

    CenterForm Me
    Label1 = "Developed by:" & Chr(10)
    Label1 = Label1 & "Fausto Cruz Arruda" & Chr(10)
    Label1 = Label1 & "arruda@ sinainet.com.br" & Chr(10)
    Label1 = Label1 & "Londrina - Brazil"

End Sub


