VERSION 5.00
Begin VB.UserControl XPcmd 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   DefaultCancel   =   -1  'True
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   64
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ToolboxBitmap   =   "XPCmd.ctx":0000
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   7305
      ScaleHeight     =   1455
      ScaleWidth      =   2535
      TabIndex        =   1
      Top             =   4305
      Width           =   2535
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   7005
      Picture         =   "XPCmd.ctx":0312
      ScaleHeight     =   1440
      ScaleWidth      =   7560
      TabIndex        =   0
      Top             =   2355
      Width           =   7560
   End
End
Attribute VB_Name = "XPcmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'original ideal from Teh Ming Han (teh_minghan@hotmail.com)
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINT_API
    x As Long
    y As Long
End Type

Const RGN_AND = 1&
Const RGN_OR = 2&
Const RGN_XOR = 3&
Const RGN_DIFF = 4&
Const RGN_COPY = 5&
Const DT_CENTER = &H1
Const DT_SINGLELINE = &H20
Const DT_VCENTER = &H4

Dim m_txtRect  As RECT
Dim m_Font As Font
Dim m_ForeColor As OLE_COLOR
Dim m_Value As Boolean
Dim m_sCaption  As String
Dim lwFontAlign As Long

Event Click()
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_UserMemId = -607
Event MouseLeave()

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
    lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    Enabled = True
    m_Value = False
    m_ForeColor = vbWhite
    Set Font = UserControl.Font
    m_sCaption = "OK"
    UserControl_Paint
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DrawButton 1
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If UserControl.Enabled = False Then Exit Sub
        If x >= 0 And x <= ScaleWidth And _
           y >= 0 And y <= ScaleHeight Then
            ' Make all messages get sent to the UserControl for a while
            SetCapture hWnd
            If Button = 1 Then
                DrawButton 1
            Else
                DrawButton 2
            End If
            RaiseEvent MouseMove(Button, Shift, x, y)
        Else
            ' Cursor went outside of the control. Release messages to be sent to wherever.
            DrawButton 0
            ReleaseCapture
            RaiseEvent MouseLeave
        End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DrawButton 0
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Paint()
    UserControl.ForeColor = m_ForeColor
    If Enabled = False Then
         UserControl.ForeColor = RGB(125, 125, 125)
    End If
    DrawButton 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
    m_sCaption = PropBag.ReadProperty("Caption", "")
    Set Font = PropBag.ReadProperty("Font", UserControl.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", UserControl.ForeColor)
    m_Value = PropBag.ReadProperty("Value", False)
    UserControl.ForeColor = m_ForeColor
    UserControl_Paint
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    UserControl_Paint
End Property

Private Sub UserControl_Resize()
    Dim hRgn As Long
    hRgn = CreateRoundRectRgn(2, 2, ScaleWidth - 2, ScaleHeight - 2, 2, 2)
    SetWindowRgn hWnd, hRgn, True
    UserControl_Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Caption", m_sCaption, "")
    Call PropBag.WriteProperty("Font", m_Font, UserControl.Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, UserControl.Ambient.ForeColor)
    Call PropBag.WriteProperty("Value", m_Value, False)
    UserControl_Paint
End Sub

Public Property Let Caption(ByVal NewCaption As String)
Attribute Caption.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    m_sCaption = NewCaption
    PropertyChanged "Caption"
    UserControl_Paint
End Property

Public Property Get Caption() As String
    Caption = m_sCaption
End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    Set UserControl.Font = vNewFont
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    UserControl.ForeColor = m_ForeColor
    UserControl_Resize
End Property

Sub DrawCaption(state As Boolean)
    If state Then
        SetRect m_txtRect, 3, 3, ScaleWidth - 1, ScaleHeight - 1
    Else
        SetRect m_txtRect, 1, 1, ScaleWidth - 1, ScaleHeight - 1
    End If
    lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    DrawText hdc, m_sCaption, -1, m_txtRect, lwFontAlign
End Sub

Public Property Let Value(bValue As Boolean)
On Error GoTo VHandler
    m_Value = bValue
    PropertyChanged "Value"
    If bValue Then RaiseEvent Click
VHandler:
End Property

Public Property Get Value() As Boolean
On Error GoTo VHandler
    Value = m_Value
    Refresh
    Exit Property
VHandler:
End Property

Private Sub DrawButton(z As Integer)
    Dim brx, bry, bw, bh As Integer
    Dim Py1, Py2, Px1, Px2, pW, pH As Integer
    pW = 3
    pH = 3
    Px1 = 3
    Py1 = 3
    brx = ScaleWidth - pW
    bry = ScaleHeight - pH
    bw = ScaleWidth - (pW * 2)
    bh = ScaleHeight - (pH * 2)
    Select Case z
        Case 0:
            Pic.PaintPicture PicButton, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, 0, 0, PicButton.ScaleWidth, PicButton.ScaleHeight
        Case 1:
            Pic.PaintPicture PicButton, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, PicButton.ScaleWidth / 3, 0, PicButton.ScaleWidth, PicButton.ScaleHeight
        Case 2:
            Pic.PaintPicture PicButton, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, (PicButton.ScaleWidth / 3) * 2, 0, PicButton.ScaleWidth, PicButton.ScaleHeight
    End Select
    Pic.Picture = Pic.Image
    Pic.Refresh
    Py2 = Pic.Height - Py1
    Px2 = (Pic.Width / 3) - Px1
    PaintPicture Pic, 0, 0, pW, pH, 0, 0, pW, pH
    PaintPicture Pic, brx, 0, pW, pH, Px2, 0, pW, pH
    PaintPicture Pic, brx, bry, pW, pH, Px2, Py2, pW, pH
    PaintPicture Pic, 0, bry, pW, pH, 0, Py2, pW, pH
    PaintPicture Pic, Px1, 0, bw, pH, Px1, 0, Px2 - pW, pH
    PaintPicture Pic, brx, Py1, pW, bh, Px2, Py1, pW, Py2 - pH
    PaintPicture Pic, 0, Py1, pW, bh, 0, Py1, pW, Py2 - pH
    PaintPicture Pic, Px1, bry, bw, pH, Px1, Py2, Px2 - pW, pH
    PaintPicture Pic, Px1, Py1, bw, bh, Px1, Py1, Px2 - pW, Py2 - pH
    If z = 1 Then
        DrawCaption True
    Else
        DrawCaption False
    End If
 End Sub
