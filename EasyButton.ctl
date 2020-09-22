VERSION 5.00
Begin VB.UserControl EasyButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
   ScaleHeight     =   47
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
   ToolboxBitmap   =   "EasyButton.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1530
      Top             =   1620
   End
End
Attribute VB_Name = "EasyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_DEFAULT = DT_CENTER Or DT_VCENTER

Public Enum Alignment
    [CenterCenter]
    [CenterTop]
    [CenterBottom]
    [LeftCenter]
    [LeftTop]
    [LeftBottom]
    [RightCenter]
    [RightTop]
    [RightBottom]
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim Rt As RECT
Dim Pt As POINTAPI

Dim DC As Long
Dim Obj As Long
Dim MouseOver As Boolean
Dim MouseButton As Integer
Dim ButtonState As Integer
Dim PtIn As Boolean
Dim PicHeight As Integer
Dim PicWidth As Integer
Dim Pic As StdPicture
Dim mCaption As String
Dim mAlign As Alignment
Dim HasPicture As Boolean

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552

    frmAboutBox.Show 1

End Sub
Public Property Get Caption() As String

    Caption = mCaption

End Property
Public Property Get Enabled() As Boolean

    Enabled = UserControl.Enabled

End Property
Public Property Let Caption(ByVal newCaption As String)

    mCaption = newCaption
    PropertyChanged "Caption"
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If
    
End Property
Public Property Let AccessKey(ByVal newKey As String)

    UserControl.AccessKeys() = newKey
    PropertyChanged "AccessKey"
    
End Property
Public Property Let Enabled(ByVal newEnabled As Boolean)

    UserControl.Enabled() = newEnabled
    PropertyChanged "Enabled"
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If
    
End Property
Public Property Let Align(ByVal newAlign As Alignment)

    mAlign = newAlign
    PropertyChanged "Align"
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If

End Property
Private Sub DrawMouseOut()
    
    If Not HasPicture Then
        Cls
        UserControl.Line (0, 0)-Step(ScaleWidth - 1, ScaleHeight - 1), &HCFCFCF, B
    Else
        BitBlt hdc, 0, 0, PicWidth, (PicHeight / 4), DC, 0, 0, vbSrcCopy
    End If
    If Trim(mCaption) <> "" Then
        Rt.Left = 0
        Rt.Top = 0
        Rt.Bottom = ScaleHeight
        Rt.Right = ScaleWidth
        OldColor = UserControl.ForeColor
        If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
        DrawText hdc, Trim(mCaption), Len(Trim(mCaption)), Rt, GetAlign(mAlign) Or DT_NOCLIP Or DT_SINGLELINE
        UserControl.ForeColor = OldColor
    End If
    Refresh
    ButtonState = 0

End Sub
Private Sub DrawUp()
    
    If Not HasPicture Then
        Cls
        UserControl.Line (0, 0)-Step(ScaleWidth - 1, 0), vb3DHighlight
        UserControl.Line (0, 0)-Step(0, ScaleHeight - 1), vb3DHighlight
        UserControl.Line (0, ScaleHeight - 1)-Step(ScaleWidth, 0), vb3DDKShadow
        UserControl.Line (ScaleWidth - 1, 0)-Step(0, ScaleHeight - 1), vb3DDKShadow
    Else
        BitBlt hdc, 0, 0, PicWidth, (PicHeight / 4), DC, 0, (PicHeight / 4), vbSrcCopy
    End If
    
    If Trim(mCaption) <> "" Then
        Rt.Left = 0
        Rt.Top = 0
        Rt.Bottom = ScaleHeight
        Rt.Right = ScaleWidth
        OldColor = UserControl.ForeColor
        If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
        DrawText hdc, Trim(mCaption), Len(Trim(mCaption)), Rt, GetAlign(mAlign) Or DT_NOCLIP Or DT_SINGLELINE
        UserControl.ForeColor = OldColor
    End If
    Refresh
    ButtonState = 1
    
End Sub
Private Sub DrawDown()
    
    If Not HasPicture Then
        Cls
        UserControl.Line (0, 0)-Step(ScaleWidth - 1, 0), vb3DDKShadow
        UserControl.Line (0, 0)-Step(0, ScaleHeight - 1), vb3DDKShadow
        UserControl.Line (0, ScaleHeight - 1)-Step(ScaleWidth, 0), vb3DHighlight
        UserControl.Line (ScaleWidth - 1, 0)-Step(0, ScaleHeight - 1), vb3DHighlight
    Else
        BitBlt hdc, 0, 0, PicWidth, (PicHeight / 4), DC, 0, (PicHeight / 4) * 2, vbSrcCopy
    End If
    
    If Trim(mCaption) <> "" Then
        Rt.Left = 1
        Rt.Top = 1
        Rt.Bottom = ScaleHeight + 1
        Rt.Right = ScaleWidth + 1
        OldColor = UserControl.ForeColor
        If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
        DrawText hdc, Trim(mCaption), Len(Trim(mCaption)), Rt, GetAlign(mAlign) Or DT_NOCLIP Or DT_SINGLELINE
        UserControl.ForeColor = OldColor
    End If
    Refresh
    ButtonState = 2
    
End Sub
Private Function GetAlign(ByVal Alng As Alignment) As Long
    
    Select Case Alng
        Case 0: GetAlign = DT_CENTER Or DT_VCENTER
        Case 1: GetAlign = DT_CENTER Or DT_TOP
        Case 2: GetAlign = DT_CENTER Or DT_BOTTOM
        Case 3: GetAlign = DT_LEFT Or DT_VCENTER
        Case 4: GetAlign = DT_LEFT Or DT_TOP
        Case 5: GetAlign = DT_LEFT Or DT_BOTTOM
        Case 6: GetAlign = DT_RIGHT Or DT_VCENTER
        Case 7: GetAlign = DT_RIGHT Or DT_TOP
        Case 8: GetAlign = DT_RIGHT Or DT_BOTTOM
    End Select

End Function
Private Function IsActiveWindow() As Boolean

    On Error Resume Next
    If GetActiveWindow() <> UserControl.Parent.hwnd Then
        Timer1.Enabled = False
        DrawMouseOut
        ButtonState = 0
        IsActiveWindow = False
    Else
        IsActiveWindow = True
    End If
    DoEvents
    
End Function
Private Sub Timer1_Timer()
    
    If Not IsActiveWindow Then Exit Sub
    GetCursorPos Pt
    ScreenToClient hwnd, Pt
    MouseOver = Not ((Pt.X < 0) Or (Pt.X > ScaleWidth) Or (Pt.Y < 0) Or (Pt.Y > ScaleHeight))
    If HasPicture Then
        If Not PtIn Then MouseOver = False
    End If
    If MouseOver Then
        If MouseButton = 1 Then
            If ButtonState <> 2 Then
                DrawDown
                ButtonState = 2
            End If
        Else
            If ButtonState <> 1 Then
                DrawUp
                ButtonState = 1
            End If
        End If
    Else
        If MouseButton = 1 Then
            If ButtonState <> 1 Then
                DrawUp
                ButtonState = 1
            End If
        Else
            Timer1.Enabled = False
            If ButtonState <> 0 Then
                DrawMouseOut
                ButtonState = 0
            End If
        End If
    End If

End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

    RaiseEvent Click

End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Not IsActiveWindow Then Exit Sub
    PtIn = (GetPixel(DC, X, Y + ((PicHeight / 4) * 3)) = 0)
    Timer1.Enabled = True
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Paint()

    DoEvents

End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    mAlign = PropBag.ReadProperty("Align", DT_DEFAULT)
    mCaption = PropBag.ReadProperty("Caption", "Command")
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HE0E0E0)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.AccessKeys = PropBag.ReadProperty("AccessKey", "")

    If UserControl.Ambient.UserMode Then
        DeleteObject Obj
        DeleteDC DC
        DC = CreateCompatibleDC(hdc)
        Obj = SelectObject(DC, Picture)
    End If
    
    Rt.Left = 0
    Rt.Top = 0
    Rt.Right = ScaleWidth
    Rt.Bottom = ScaleHeight
    
    If Not HasPicture Then DrawMouseOut
    If Not UserControl.Enabled Then Exit Sub
    
    If Trim(mCaption) <> "" Then
        OldColor = UserControl.ForeColor
        If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
        DrawText hdc, Trim(mCaption), Len(Trim(mCaption)), Rt, GetAlign(mAlign) Or DT_NOCLIP Or DT_SINGLELINE
        UserControl.ForeColor = OldColor
        X = InStr(1, mCaption, "&")
        If X > 0 Then UserControl.AccessKeys = Mid(mCaption, X + 1, 1)
    End If

End Sub
Private Sub UserControl_Click()
    
    If Not PtIn And HasPicture Then Exit Sub
    RaiseEvent Click

End Sub
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    
    BackColor = UserControl.BackColor

End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If
    
End Property
Private Sub UserControl_DblClick()
    
    RaiseEvent DblClick

End Sub
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    
    Set Font = UserControl.Font

End Property
Public Property Set Font(ByVal New_Font As Font)
    
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If
    
End Property
Public Property Get ForeColor() As OLE_COLOR
    
    ForeColor = UserControl.ForeColor
    
End Property
Public Property Get Align() As Alignment
    
    Align = mAlign
    
End Property
Public Property Get AccessKey() As String
    
    AccessKey = UserControl.AccessKeys
    
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If
    
End Property
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    
    hwnd = UserControl.hwnd

End Property
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyDown(KeyCode, Shift)

End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    
    RaiseEvent KeyPress(KeyAscii)

End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyUp(KeyCode, Shift)
    
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    PtIn = (GetPixel(DC, X, Y + ((PicHeight / 4) * 3)) = 0)
    If Not PtIn And HasPicture Then Exit Sub
    MouseButton = 1
    If Not Timer1.Enabled Then Timer1.Enabled = True
    
End Sub
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    
    Set MouseIcon = UserControl.MouseIcon

End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"

End Property
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    
    MousePointer = UserControl.MousePointer

End Property
Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"

End Property
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If MouseButton = 1 And PtIn Then
        DrawUp
        ButtonState = 1
    End If
    MouseButton = 0
    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    
    Set Picture = UserControl.Picture

End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
    If UserControl.Picture <> 0 Then
        PicHeight = ScaleY(New_Picture.Height, 8, 3)
        PicWidth = ScaleX(New_Picture.Width, 8, 3)
        Height = PicHeight / 4
        Width = ScaleX(PicWidth, 3, 1)
        If Trim(mCaption) <> "" Then
            Rt.Left = 0
            Rt.Top = 0
            Rt.Bottom = ScaleHeight
            Rt.Right = ScaleWidth
            OldColor = UserControl.ForeColor
            If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
            DrawText hdc, Trim(mCaption), Len(Trim(mCaption)), Rt, GetAlign(mAlign) Or DT_NOCLIP Or DT_SINGLELINE
            UserControl.ForeColor = OldColor
        End If
        HasPicture = True
    Else
        HasPicture = False
    End If
    
End Property
Private Sub UserControl_InitProperties()
    
    Set UserControl.Font = Ambient.Font
    
End Sub
Private Sub UserControl_Resize()
    
    If UserControl.Picture <> 0 Then
        Height = ScaleY(PicHeight, 3, 1) / 4
        Width = ScaleX(PicWidth, 3, 1)
    Else
        Cls
        DrawMouseOut
    End If

End Sub
Private Sub UserControl_Terminate()
    
    DoEvents
    DeleteObject Obj
    DeleteDC DC
    
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", mCaption, "Command")
    Call PropBag.WriteProperty("Align", mAlign, DT_DEFAULT)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HE0E0E0)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("AccessKey", UserControl.AccessKeys, "")
    
End Sub
