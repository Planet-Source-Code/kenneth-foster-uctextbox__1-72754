VERSION 5.00
Begin VB.UserControl XPTextbox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   197
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   4890
      Width           =   4065
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4065
      Width           =   4020
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      Top             =   3480
      Width           =   3990
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Top             =   2970
      Width           =   3915
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   2640
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   3885
   End
End
Attribute VB_Name = "XPTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Public Enum eScrollBars
   None = 0
   Horizontal = 1
   Vertical = 2
   Both = 3
End Enum

Public Enum eAlign
   [Left Justify] = 0
   [Right Justify] = 1
   Center = 2
End Enum
'------------------------------------------------------------
'draw and set rectangular area of the control
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

'draw by pixel or by line
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

'select and delete created objects
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'create regions of pixels and remove them to make the control transparent
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Const RGN_DIFF As Long = 4

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Const m_def_Alignment = 0
Private Const m_def_BackColor = vbWhite

Private m_Alignment As Integer
Private m_BackColor As OLE_COLOR

Private rc As RECT
Private W As Long, H As Long
Private regMain As Long, rgn1 As Long
Private r As Long, l As Long, t As Long, B As Long
Private m_MultiLine As Boolean
Private m_ScrollBars As Integer

Private Sub DrawButton()
Dim pt As POINTAPI, Pen As Long, hPen As Long
Dim I As Long, ColorR As Long, ColorG As Long, ColorB As Long
Dim hBrush As Long

  With UserControl
  
    hBrush = CreateSolidBrush(RGB(0, 60, 116))
    FrameRect UserControl.hDC, rc, hBrush
    DeleteObject hBrush
    
    'Left top corner
    SetPixel .hDC, l, t + 1, RGB(122, 149, 168)
    SetPixel .hDC, l + 1, t + 1, RGB(37, 87, 131)
    SetPixel .hDC, l + 1, t, RGB(122, 149, 168)
    
    'right top corner
    SetPixel .hDC, r - 1, t, RGB(122, 149, 168)
    SetPixel .hDC, r - 1, t + 1, RGB(37, 87, 131)
    SetPixel .hDC, r, t + 1, RGB(122, 149, 168)
    
    'left bottom corner
    SetPixel .hDC, l, B - 2, RGB(122, 149, 168)
    SetPixel .hDC, l + 1, B - 2, RGB(37, 87, 131)
    SetPixel .hDC, l + 1, B - 1, RGB(122, 149, 168)
    
    'right bottom corner
    SetPixel .hDC, r, B - 2, RGB(122, 149, 168)
    SetPixel .hDC, r - 1, B - 2, RGB(37, 87, 131)
    SetPixel .hDC, r - 1, B - 1, RGB(122, 149, 168)
  End With
  
  DeleteObject regMain
  regMain = CreateRectRgn(0, 0, W, H)
  rgn1 = CreateRectRgn(0, 0, 1, 1)            'Left top coner
  CombineRgn regMain, regMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  rgn1 = CreateRectRgn(0, H - 1, 1, H)      'Left bottom corner
  CombineRgn regMain, regMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  rgn1 = CreateRectRgn(W - 1, 0, W, 1)      'Right top corner
  CombineRgn regMain, regMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  rgn1 = CreateRectRgn(W - 1, H - 1, W, H) 'Right bottom corner
  CombineRgn regMain, regMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  SetWindowRgn UserControl.hWnd, regMain, True
End Sub

Private Sub UserControl_InitProperties()
   Set Text1.Font = Ambient.Font
   m_BackColor = m_def_BackColor
   m_Alignment = m_def_Alignment
End Sub

Private Sub UserControl_Resize()
  GetClientRect UserControl.hWnd, rc
  With rc
    r = .Right - 1: l = .Left: t = .Top: B = .Bottom
    W = .Right: H = .Bottom
  End With
  
   Text1.Visible = False
   Text2.Visible = False
   Text3.Visible = False
   Text4.Visible = False
   Text5.Visible = False
   
   Select Case m_ScrollBars
      Case 0      'none
         If m_MultiLine = True Then
            Text1.Left = 3
            Text1.Top = 2
            Text1.Width = UserControl.ScaleWidth - 6
            Text1.Height = UserControl.ScaleHeight - 4
            Text1.Visible = True
         Else           'single line
            Text2.Left = 3
            Text2.Top = 2
            Text2.Width = UserControl.ScaleWidth - 6
            Text2.Height = UserControl.ScaleHeight - 4
            Text2.Visible = True
            m_MultiLine = False
         End If
      Case 1          'horizontal
        Text3.Left = 3
        Text3.Top = 2
        Text3.Width = UserControl.ScaleWidth - 6
        Text3.Height = UserControl.ScaleHeight - 4
        Text3.Visible = True
        m_MultiLine = True
      Case 2          'vertical
        Text4.Left = 3
        Text4.Top = 2
        Text4.Width = UserControl.ScaleWidth - 6
        Text4.Height = UserControl.ScaleHeight - 4
        Text4.Visible = True
        m_MultiLine = True
      Case 3          'both
        Text5.Left = 3
        Text5.Top = 2
        Text5.Width = UserControl.ScaleWidth - 6
        Text5.Height = UserControl.ScaleHeight - 4
        Text5.Visible = True
        m_MultiLine = True
   End Select
   
 ' UserControl.Cls
  UserControl.Refresh
  DrawButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
      Set Text1.Font = .ReadProperty("Font", Ambient.Font)
      Set Text2.Font = .ReadProperty("Font", Ambient.Font)
      Set Text3.Font = .ReadProperty("Font", Ambient.Font)
      Set Text4.Font = .ReadProperty("Font", Ambient.Font)
      Set Text5.Font = .ReadProperty("Font", Ambient.Font)
      Text1.ForeColor = .ReadProperty("ForeColor", vbButtonText)
      Text1.MaxLength = .ReadProperty("MaxLength", 0)
      Text1.Locked = .ReadProperty("Locked", False)
      Text1.Enabled = .ReadProperty("Enabled", True)
      Text1.Text = .ReadProperty("Text", "")
      Text2.ForeColor = .ReadProperty("ForeColor", vbButtonText)
      Text2.MaxLength = .ReadProperty("MaxLength", 0)
      Text2.Locked = .ReadProperty("Locked", False)
      Text2.Enabled = .ReadProperty("Enabled", True)
      Text2.Text = .ReadProperty("Text", "")
      Text3.ForeColor = .ReadProperty("ForeColor", vbButtonText)
      Text3.MaxLength = .ReadProperty("MaxLength", 0)
      Text3.Locked = .ReadProperty("Locked", False)
      Text3.Enabled = .ReadProperty("Enabled", True)
      Text3.Text = .ReadProperty("Text", "")
      Text4.ForeColor = .ReadProperty("ForeColor", vbButtonText)
      Text4.MaxLength = .ReadProperty("MaxLength", 0)
      Text4.Locked = .ReadProperty("Locked", False)
      Text4.Enabled = .ReadProperty("Enabled", True)
      Text4.Text = .ReadProperty("Text", "")
      Text5.ForeColor = .ReadProperty("ForeColor", vbButtonText)
      Text5.MaxLength = .ReadProperty("MaxLength", 0)
      Text5.Locked = .ReadProperty("Locked", False)
      Text5.Enabled = .ReadProperty("Enabled", True)
      Text5.Text = .ReadProperty("Text", "")
      MultiLine = .ReadProperty("MultiLine", False)
      ScrollBars = .ReadProperty("ScrollBars", 0)
      BackColor = .ReadProperty("BackColor", m_def_BackColor)
      Alignment = .ReadProperty("Alignment", m_def_Alignment)
   End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Font", Text1.Font, Ambient.Font
      .WriteProperty "ForeColor", Text1.ForeColor, vbButtonText
      .WriteProperty "MaxLength", Text1.MaxLength, 0
      .WriteProperty "Locked", Text1.Locked, False
      .WriteProperty "Enabled", Text1.Enabled, True
      .WriteProperty "Text", Text1.Text, ""
      .WriteProperty "Font", Text2.Font, Ambient.Font
      .WriteProperty "ForeColor", Text2.ForeColor, vbButtonText
      .WriteProperty "MaxLength", Text2.MaxLength, 0
      .WriteProperty "Locked", Text2.Locked, False
      .WriteProperty "Enabled", Text2.Enabled, True
      .WriteProperty "Text", Text2.Text, ""
      .WriteProperty "Font", Text3.Font, Ambient.Font
      .WriteProperty "ForeColor", Text3.ForeColor, vbButtonText
      .WriteProperty "MaxLength", Text3.MaxLength, 0
      .WriteProperty "Locked", Text3.Locked, False
      .WriteProperty "Enabled", Text3.Enabled, True
      .WriteProperty "Text", Text3.Text, ""
      .WriteProperty "Font", Text4.Font, Ambient.Font
      .WriteProperty "ForeColor", Text4.ForeColor, vbButtonText
      .WriteProperty "MaxLength", Text4.MaxLength, 0
      .WriteProperty "Locked", Text4.Locked, False
      .WriteProperty "Enabled", Text4.Enabled, True
      .WriteProperty "Text", Text4.Text, ""
      .WriteProperty "Font", Text5.Font, Ambient.Font
      .WriteProperty "ForeColor", Text5.ForeColor, vbButtonText
      .WriteProperty "MaxLength", Text5.MaxLength, 0
      .WriteProperty "Locked", Text5.Locked, False
      .WriteProperty "Enabled", Text5.Enabled, True
      .WriteProperty "Text", Text5.Text, ""
      .WriteProperty "MultiLine", m_MultiLine, False
      .WriteProperty "ScrollBars", m_ScrollBars, 0
      .WriteProperty "BackColor", m_BackColor, m_def_BackColor
      .WriteProperty "Alignment", m_Alignment, m_def_Alignment
   End With
End Sub

Public Property Get Alignment() As eAlign
   Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal NewAlignment As eAlign)
   m_Alignment = NewAlignment
   If m_Alignment = 1 Or m_Alignment = 2 Then
      If m_ScrollBars = 1 Or m_ScrollBars = 3 Then m_ScrollBars = 2
   End If
   Text1.Alignment = m_Alignment
   Text2.Alignment = m_Alignment
   Text3.Alignment = m_Alignment
   Text4.Alignment = m_Alignment
   Text5.Alignment = m_Alignment
   PropertyChanged "Alignment"
   UserControl_Resize
End Property

Public Property Get BackColor() As OLE_COLOR
   BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
   m_BackColor = NewBackColor
   Text1.BackColor = NewBackColor
   Text2.BackColor = NewBackColor
   Text3.BackColor = NewBackColor
   Text4.BackColor = NewBackColor
   Text5.BackColor = NewBackColor
   UserControl.BackColor = NewBackColor
   PropertyChanged "BackColor"
   DrawButton
End Property

Public Property Get Font() As Font
    Set Font = Text1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    Set Text2.Font = New_Font
    Set Text3.Font = New_Font
    Set Text4.Font = New_Font
    Set Text5.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Text1.ForeColor = New_ForeColor
    Text2.ForeColor = New_ForeColor
    Text3.ForeColor = New_ForeColor
    Text4.ForeColor = New_ForeColor
    Text5.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get MaxLength() As Single
    MaxLength = Text1.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Single)
    Text1.MaxLength = New_MaxLength
    Text2.MaxLength = New_MaxLength
    Text3.MaxLength = New_MaxLength
    Text4.MaxLength = New_MaxLength
    Text5.MaxLength = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Public Property Get MultiLine() As Boolean
    MultiLine = m_MultiLine
End Property

Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
    m_MultiLine = New_MultiLine
    PropertyChanged "MultiLine"
    UserControl_Resize
End Property

Public Property Get Locked() As Boolean
    Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Text1.Locked = New_Locked
    Text2.Locked = New_Locked
    Text3.Locked = New_Locked
    Text4.Locked = New_Locked
    Text5.Locked = New_Locked
    PropertyChanged "Locked"
End Property

Public Property Get Enabled() As Boolean
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Text1.Enabled = New_Enabled
    Text2.Enabled = New_Enabled
    Text3.Enabled = New_Enabled
    Text4.Enabled = New_Enabled
    Text5.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Text() As String
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text = New_Text
    Text2.Text = New_Text
    Text3.Text = New_Text
    Text4.Text = New_Text
    Text5.Text = New_Text
    PropertyChanged "Text"
    UserControl_Resize
End Property

Public Property Get ScrollBars() As eScrollBars
    ScrollBars = m_ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As eScrollBars)
    m_ScrollBars = New_ScrollBars
    PropertyChanged "ScrollBars"
    UserControl_Resize
End Property

