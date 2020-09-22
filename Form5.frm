VERSION 5.00
Object = "*\AEasyButton.vbp"
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "How to draw a picture"
   ClientHeight    =   5505
   ClientLeft      =   2985
   ClientTop       =   1845
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   367
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   Begin EasyX.EasyButton EasyButton1 
      Height          =   525
      Left            =   3915
      TabIndex        =   9
      Top             =   4860
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   926
      Caption         =   ""
      Align           =   0
      Picture         =   "Form5.frx":2910
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
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -90
      TabIndex        =   7
      Top             =   2700
      Width           =   5595
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -90
      TabIndex        =   6
      Top             =   4680
      Width           =   5595
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   270
      TabIndex        =   8
      Top             =   2385
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Height          =   1635
      Left            =   1980
      Top             =   2880
      Width           =   1635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   4
      X1              =   75
      X2              =   135
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Same Picture"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   5
      Top             =   4365
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mask"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   3
      Left            =   4140
      TabIndex        =   4
      Top             =   4050
      Width           =   390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   3
      X1              =   213
      X2              =   273
      Y1              =   279
      Y2              =   279
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MouseDown"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   450
      TabIndex        =   3
      Top             =   3735
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   2
      X1              =   93
      X2              =   153
      Y1              =   258
      Y2              =   258
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MouseOver"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   4140
      TabIndex        =   2
      Top             =   3420
      Width           =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   1
      X1              =   213
      X2              =   273
      Y1              =   237
      Y2              =   237
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MouseOut"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   585
      TabIndex        =   1
      Top             =   3015
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   0
      X1              =   93
      X2              =   153
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   2160
      Picture         =   "Form5.frx":5FB0
      Top             =   3015
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2220
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   5055
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EasyButton1_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    Label1 = "You can make buttons in any format." & Chr(10)
    Label1 = Label1 & "(The secret is the picture)" & Chr(10) & Chr(10)
    Label1 = Label1 & "How to make a picture:" & Chr(10)
    Label1 = Label1 & "     You must draw 4 images of same size in a only image." & Chr(10)
    Label1 = Label1 & "     The first image is 'mouse out'." & Chr(10)
    Label1 = Label1 & "     Second image is 'mouse over'." & Chr(10)
    Label1 = Label1 & "     Third is 'mouse down'" & Chr(10)
    Label1 = Label1 & "     Fourth is a mask. The mask define the clickable area." & Chr(10)
    Label3 = "See image below"

End Sub
Private Sub Form_Resize()
    
    For X = 1 To ScaleWidth Step 100
        For Y = 1 To ScaleHeight Step 100
            PaintPicture Picture, X, Y
        Next
    Next

End Sub


