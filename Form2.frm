VERSION 5.00
Object = "*\AEasyButton.vbp"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Without Pictures"
   ClientHeight    =   1620
   ClientLeft      =   3285
   ClientTop       =   2970
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin EasyX.EasyButton EasyButton1 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   675
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&OK"
      Align           =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      AccessKey       =   "O"
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -45
      TabIndex        =   2
      Top             =   405
      Width           =   4605
   End
   Begin EasyX.EasyButton EasyButton2 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1125
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Close"
      Align           =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      AccessKey       =   "C"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Without picture, with Caption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   90
      Width           =   4155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Accept Hot Keys and Mnemonics.    Various types of alignment for captions."
      Height          =   825
      Left            =   135
      TabIndex        =   0
      Top             =   675
      Width           =   2850
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EasyButton2_Click()

    Unload Me

End Sub
