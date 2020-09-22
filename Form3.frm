VERSION 5.00
Object = "*\AEasyButton.vbp"
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "With Pictures"
   ClientHeight    =   1605
   ClientLeft      =   3315
   ClientTop       =   3030
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   107
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   284
   ShowInTaskbar   =   0   'False
   Begin EasyX.EasyButton EasyButton1 
      Height          =   630
      Left            =   90
      TabIndex        =   1
      Top             =   855
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1111
      Caption         =   ""
      Align           =   0
      Picture         =   "Form3.frx":2910
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
   End
   Begin EasyX.EasyButton EasyButton2 
      Height          =   630
      Left            =   810
      TabIndex        =   2
      Top             =   855
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1111
      Caption         =   ""
      Align           =   0
      Picture         =   "Form3.frx":4A40
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
   End
   Begin EasyX.EasyButton EasyButton3 
      Height          =   630
      Left            =   1530
      TabIndex        =   3
      Top             =   855
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1111
      Caption         =   ""
      Align           =   0
      Picture         =   "Form3.frx":6B70
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
   End
   Begin EasyX.EasyButton EasyButton4 
      Height          =   630
      Left            =   2205
      TabIndex        =   4
      Top             =   855
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1111
      Caption         =   ""
      Align           =   0
      Picture         =   "Form3.frx":8CA0
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
   End
   Begin EasyX.EasyButton EasyButton5 
      Height          =   630
      Left            =   3510
      TabIndex        =   5
      Top             =   855
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1111
      Caption         =   ""
      Align           =   0
      Picture         =   "Form3.frx":ADD0
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
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "With picture, no caption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00655450&
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00655450&
      Index           =   0
      X1              =   0
      X2              =   288
      Y1              =   36
      Y2              =   36
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DFD5CC&
      Index           =   1
      X1              =   0
      X2              =   288
      Y1              =   37
      Y2              =   37
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EasyButton5_Click()

    Unload Me

End Sub

Private Sub Form_Resize()
    
    For X = 1 To ScaleWidth Step 100
        For Y = 1 To ScaleHeight Step 100
            PaintPicture Picture, X, Y
        Next
    Next

End Sub


