VERSION 5.00
Begin VB.UserControl DPMCCombo 
   AutoRedraw      =   -1  'True
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   BeginProperty Font 
      Name            =   "Marlett"
      Size            =   9.75
      Charset         =   2
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   117
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "DPMCCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Email Me: zhujy@samling.com.my
'Welcome to visit our company WebSite at:
  'www.Samling.com.my
  'www.Samling.com.cn
'Samling Group---One of a leading and largest WoodBased Industry company in Asia
'Copyright only credited to Author

Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As Long, ByVal hpal As Long, ByRef lpcolorref As Long)
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cY As Long, ByVal fuFlags As Long) As Long

Const DST_COMPLEX = &H0
Const DST_TEXT = &H1
Const DST_PREFIXTEXT = &H2
Const DST_ICON = &H3
Const DST_BITMAP = &H4

Const DSS_NORMAL = &H0
Const DSS_UNION = &H10
Const DSS_DISABLED = &H20
Const DSS_MONO = &H80
Const DSS_RIGHT = &H8000
Const SM_CXHTHUMB = 10

Const DT_BOTTOM = &H8
Const DT_CENTER = &H1
Const DT_LEFT = &H0
Const DT_NOCLIP = &H100
Const DT_NOPREFIX = &H800
Const DT_RIGHT = &H2
Const DT_SINGLELINE = &H20
Const DT_TOP = &H0
Const DT_VCENTER = &H4
Const DT_WORDBREAK = &H10
Const m_def_IconSizeWidth = 16
Const m_def_IconSizeHeight = 16
Const m_def_Enabled = True
Const m_def_ShowIcon = True
Const m_def_Style = 0
Const m_def_FocusColor = &HC00000

Const m_def_Text = ""
Const m_def_BorderColor = &HFF8080
Const m_def_BorderColorOver = &H80FF&
Const m_def_BorderColorDown = &HFF&
Const m_def_BgColor = &HFFFFFF
Const m_def_BgColorOver = &HFFFFFF
Const m_def_BgColorDown = &HFFFFFF
Const m_def_ButtonBgColor = &HFFC0C0
Const m_def_ButtonBgColorOver = &HD0B0B0
Const m_def_ButtonBgColorDown = &HFFC0C0
Const CRdf_BackNormal = &HC8FFFF
Const CRdf_BackSelected = &HC000&
Const CRdf_BackSelectedG1 = &HC000&
Const CRdf_BackSelectedG2 = &HE0E0E0
Const CRdf_BoxBorder = vbHighlightText
Const CRdf_BoxOffset = 23
Const CRdf_BoxRadius = 20
Const m_def_nGridWidth = 25
Const m_def_nGridHeight = 25
Const CRdf_SelectMode = 0
Const CRdf_SelectModeStyle = 0
Const df_SelectControlType = 1
Const m_def_MinListHeight = 2000
Const m_def_BoundColumns = "0"

Dim m_ColumnHeaders             As Boolean
Dim m_BorderColor               As OLE_COLOR
Dim m_BorderColorOver           As OLE_COLOR
Dim m_BorderColorDown           As OLE_COLOR
Dim m_BgColor                   As OLE_COLOR
Dim m_BgColorOver               As OLE_COLOR
Dim m_BgColorDown               As OLE_COLOR
Dim m_ButtonBgColor             As OLE_COLOR
Dim m_ButtonBgColorOver         As OLE_COLOR
Dim m_ButtonBgColorDown         As OLE_COLOR


Dim m_FocusColor                As OLE_COLOR
Dim UsrRect                     As RECT
Dim ButtRect                    As RECT
Dim Ret                         As Long
Dim CrlRet                      As Long
Dim IsMOver                     As Boolean
Dim IsMDown                     As Boolean
Dim IsButtDown                  As Boolean
Dim IsCrlOver                   As Boolean
Dim Clicked                     As Boolean
Dim InFocus                     As Boolean
Dim m_DropListEnabled           As Boolean
Dim m_Enabled                   As Boolean
Dim m_Icon                      As StdPicture
Dim M_IconOK                    As Boolean
Dim m_ShowIcon                  As Boolean
Dim m_IconSizeWidth             As Long
Dim m_IconSizeHeight            As Long
Dim m_SelectControl             As UserControlType
Dim m_Text                      As String
Dim m_MaxLength                 As Integer

'Overall Usercontrol Declare----------------------------------------------------------------------------
Public Enum UserControlType
            [XPMultiCombo] = 0
End Enum

Public Enum pbcStyle
    [pbxp] = 0
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Event Click()
Event MouseOver()
Event MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event Change()
Event DropList()
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

Public Sub DrawControl(ByVal StatusType As Long)
   Dim CurFontName As String
   Dim Brsh As Long, Clr As Long
   Dim lx As Long, ty As Long
   Dim rx As Long, by As Long
   Dim xx As Long
   Dim sh As Long
   Dim textline As Long
   Dim align As Long
   Dim lr As Long
   
   lx = ScaleLeft: ty = ScaleTop
   rx = ScaleWidth: by = ScaleHeight
      
   On Error Resume Next
   SetRect UsrRect, 0, 0, rx, by
   If Not m_Enabled Then StatusType = 0
   Cls
Select Case m_SelectControl
  Case 0
   Select Case StatusType
   
      '## Draw Button Normal--No Focus,No Mouse Event
      Case 0
         
               Call SetRect(UsrRect, 0, 0, rx - by, by)
               OleTranslateColor m_BgColor, ByVal 0&, Clr
               Brsh = CreateSolidBrush(Clr)
               FillRect hdc, UsrRect, Brsh
               DeleteObject Brsh
               
                  Call SetRect(ButtRect, rx - by, 0, rx, by)
                  OleTranslateColor m_ButtonBgColor, ByVal 0&, Clr
                  Brsh = CreateSolidBrush(Clr)
                  FillRect hdc, ButtRect, Brsh
                  DeleteObject Brsh
          
         Call SetRect(ButtRect, rx - by, 0, rx, by)
        If InFocus Then
            OleTranslateColor m_FocusColor, ByVal 0&, Clr
         Else
            OleTranslateColor m_BorderColor, ByVal 0&, Clr
         End If
         Brsh = CreateSolidBrush(Clr)
         FrameRect hdc, UsrRect, Brsh
         DeleteObject Brsh
         
        SetRect UsrRect, 0, 0, rx, by
        If InFocus Then
            OleTranslateColor m_FocusColor, ByVal 0&, Clr
         Else
            OleTranslateColor m_BorderColor, ByVal 0&, Clr
         End If
         Brsh = CreateSolidBrush(Clr)
         FrameRect hdc, UsrRect, Brsh
         DeleteObject Brsh
         
         If m_ShowIcon Then
            If Not m_Enabled Then
               lr = DrawState(hdc, 0, 0, m_Icon, 0, rx - (by + m_IconSizeWidth) / 2, (by / 2) - (m_IconSizeHeight / 2), m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_DISABLED)
            Else
               lr = DrawState(hdc, 0, 0, m_Icon, 0, rx - (by + m_IconSizeWidth) / 2, (by / 2) - (m_IconSizeHeight / 2), m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_NORMAL)
            End If
                    
         End If
         
         '## Draw Button Over
      Case 1

         SetRect UsrRect, 0, 0, rx, by
         
               OleTranslateColor m_BgColorOver, ByVal 0&, Clr
               Brsh = CreateSolidBrush(Clr)
               FillRect hdc, UsrRect, Brsh
               DeleteObject Brsh
              
                  Call SetRect(ButtRect, rx - by, 0, rx, by)
                  OleTranslateColor m_ButtonBgColorOver, ByVal 0&, Clr
                  Brsh = CreateSolidBrush(Clr)
                  FillRect hdc, ButtRect, Brsh
                  DeleteObject Brsh
         
         Call SetRect(ButtRect, rx - by, 0, rx, by)
         OleTranslateColor m_BorderColorOver, ByVal 0&, Clr
         Brsh = CreateSolidBrush(Clr)
         FrameRect hdc, ButtRect, Brsh
         DeleteObject Brsh
         
         SetRect UsrRect, 0, 0, rx, by
         OleTranslateColor m_BorderColorOver, ByVal 0&, Clr
         Brsh = CreateSolidBrush(Clr)
         FrameRect hdc, UsrRect, Brsh
         DeleteObject Brsh
         If m_ShowIcon Then
            Brsh = CreateSolidBrush(RGB(136, 141, 157))
            lr = DrawState(hdc, Brsh, 0, m_Icon, 0, rx - (by + m_IconSizeWidth) / 2 + 1.5, (by / 2) - (m_IconSizeHeight / 2) + 1.5, m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_MONO)
            DeleteObject Brsh
            lr = DrawState(hdc, 0, 0, m_Icon, 0, rx - (by + m_IconSizeWidth) / 2 - 1.5, (by / 2) - (m_IconSizeHeight / 2) - 1.5, m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_NORMAL)
         End If
         'Draw Button Down
      Case 2
         
         SetRect UsrRect, 0, 0, rx, by
        
               OleTranslateColor m_BgColorDown, ByVal 0&, Clr
               Brsh = CreateSolidBrush(Clr)
               FillRect hdc, UsrRect, Brsh
               DeleteObject Brsh
                  Call SetRect(ButtRect, rx - by, 0, rx, by)
                  OleTranslateColor m_ButtonBgColorDown, ByVal 0&, Clr
                  Brsh = CreateSolidBrush(Clr)
                  FillRect hdc, ButtRect, Brsh
                  DeleteObject Brsh
                  SetRectEmpty ButtRect
     
         Call SetRect(ButtRect, rx - by, 0, rx, by)
        
            OleTranslateColor m_BorderColorDown, ByVal 0&, Clr
        
         Brsh = CreateSolidBrush(Clr)
         FrameRect hdc, ButtRect, Brsh
         DeleteObject Brsh
        
         SetRect UsrRect, 0, 0, rx, by
         OleTranslateColor m_BorderColorDown, ByVal 0&, Clr
         Brsh = CreateSolidBrush(Clr)
         FrameRect hdc, UsrRect, Brsh
         DeleteObject Brsh
         If m_ShowIcon Then
            lr = DrawState(hdc, 0, 0, m_Icon, 0, rx - (by + m_IconSizeWidth) / 2, (by / 2) - (m_IconSizeHeight / 2), m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_NORMAL)
         End If
    
   End Select
If Not m_ShowIcon Then
    CurFontName = Font.Name
    Font.Name = "Marlett"
    If by <> 0 Then
        Call DrawText(hdc, "6", 1&, ButtRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
    End If
    Font.Name = CurFontName
End If
  
End Select
    
 '## Hide Listview's column Header?
 'Note : Better put those rungs in other Routine either showpopup or load_rs_to_lsw.
 'Because DrawControl is Executed before them.
 If m_ColumnHeaders = False Then
        HideColumnHeaders = True
 Else
        HideColumnHeaders = False
 End If
End Sub
Public Sub RefreshControl()

   If IsCrlOver Then
      Call DrawControl(2)
   Else
      Call DrawControl(0)
   End If

End Sub
Private Sub Text1_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Public Sub OLEDrag()
    Text1.OLEDrag
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Text1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub Text1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub Text1_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub Text1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub
Private Sub Text1_Change()

    m_Text = Text1.Text
    
    RaiseEvent Change

End Sub

Private Sub Text1_GotFocus()
If m_Enabled Then
    InFocus = True
    Call RefreshControl
End If
End Sub

Private Sub Text1_LostFocus()

    InFocus = False
    Call RefreshControl

End Sub

Private Sub UserControl_Initialize()

Call DrawControl(0)

End Sub

Private Sub UserControl_LostFocus()

    InFocus = False
Call RefreshControl

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If m_Enabled Then
    RaiseEvent MouseMove(Button, Shift, X, Y)
     If Not IsButtDown Then
        UserControl_MouseOut Button, Shift, X, Y
     End If
  End If
  
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If m_Enabled Then
   IsButtDown = False
   Call DrawControl(0)
   Select Case m_SelectControl
    Case 0
      If (X >= ButtRect.Left And X <= ButtRect.Right) And (Y >= ButtRect.Top And Y <= ButtRect.Bottom) Then
       Select Case m_SelectControl
          Case 0
              Call XPComboShow(1)
          End Select
      End If
   End Select
  End If
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    '## Single Line Textbox
      Text1.Move 2, (ScaleHeight / 2) - (Text1.Height / 2), ScaleWidth - ScaleHeight - 4
      
      Call RefreshControl
   
End Sub

Function UserControl_MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)

   IsCrlOver = False
   IsMOver = False
Select Case m_SelectControl
  Case 0 '## Case XPMCCombo
     If (X >= ButtRect.Left And X <= ButtRect.Right) And (Y >= ButtRect.Top And Y <= ButtRect.Bottom) Then
        If IsMOver = False Then
            
            IsMOver = True
            Ret = SetCapture(UserControl.hWnd)
            
        RaiseEvent MouseOver
        Call DrawControl(1)
        End If
    Else
        IsMOver = False
        Ret = ReleaseCapture()


    End If
    
      If (X >= 0 And X <= ScaleWidth) And (Y >= 0 And Y <= ScaleHeight) Then
        
        If IsCrlOver = False Then
           
            IsCrlOver = True
            CrlRet = SetCapture(UserControl.hWnd)
       
        RaiseEvent MouseOver
        Call DrawControl(1)
        End If
      Else
        IsMOver = False
        IsCrlOver = False
        CrlRet = ReleaseCapture()
        RaiseEvent MouseOut(Button, Shift, X, Y)
        Call DrawControl(0)
     End If
     
  End Select
   
End Function
Public Property Get Enabled() As Boolean

   Enabled = m_Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

   m_Enabled = New_Enabled
   PropertyChanged "Enabled"
   Call RefreshControl

End Property
Public Property Get ShowIcon() As Boolean
   ShowIcon = m_ShowIcon
End Property

Public Property Let ShowIcon(ByVal New_ShowIcon As Boolean)
   
   m_ShowIcon = New_ShowIcon
   PropertyChanged "ShowIcon"
   RefreshControl
  
End Property
Public Property Get Icon() As StdPicture

   Set Icon = m_Icon

End Property

Public Property Set Icon(ByVal New_Icon As StdPicture)

   Set m_Icon = New_Icon

   
   PropertyChanged "Icon"
   Call RefreshControl

End Property
Public Property Get IconSizeWidth() As Long

   IconSizeWidth = m_IconSizeWidth

End Property

Public Property Let IconSizeWidth(ByVal New_IconSizeWidth As Long)

   m_IconSizeWidth = New_IconSizeWidth
   PropertyChanged "IconSizeWidth"
   Call RefreshControl

End Property

Public Property Get IconSizeHeight() As Long

   IconSizeHeight = m_IconSizeHeight

End Property

Public Property Let IconSizeHeight(ByVal New_IconSizeHeight As Long)

   m_IconSizeHeight = New_IconSizeHeight
   PropertyChanged "IconSizeHeight"
   Call RefreshControl

End Property

Public Property Get BorderColor() As OLE_COLOR

    BorderColor = m_BorderColor

End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)

    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    Call RefreshControl

End Property

Public Property Get BorderColorOver() As OLE_COLOR

    BorderColorOver = m_BorderColorOver

End Property

Public Property Let BorderColorOver(ByVal New_BorderColorOver As OLE_COLOR)

    m_BorderColorOver = New_BorderColorOver
    PropertyChanged "BorderColorOver"
   Call RefreshControl

End Property

Public Property Get BorderColorDown() As OLE_COLOR

    BorderColorDown = m_BorderColorDown

End Property

Public Property Let BorderColorDown(ByVal New_BorderColorDown As OLE_COLOR)

    m_BorderColorDown = New_BorderColorDown
    PropertyChanged "BorderColorDown"
    Call RefreshControl

End Property

Public Property Get BgColor() As OLE_COLOR

    BgColor = m_BgColor

End Property

Public Property Let BgColor(ByVal New_BgColor As OLE_COLOR)

    m_BgColor = New_BgColor
    PropertyChanged "BgColor"
   Call RefreshControl

End Property

Public Property Get BgColorOver() As OLE_COLOR

    BgColorOver = m_BgColorOver

End Property

Public Property Let BgColorOver(ByVal New_BgColorOver As OLE_COLOR)

    m_BgColorOver = New_BgColorOver
    PropertyChanged "BgColorOver"
    Call RefreshControl

End Property

Public Property Get BgColorDown() As OLE_COLOR

    BgColorDown = m_BgColorDown

End Property

Public Property Let BgColorDown(ByVal New_BgColorDown As OLE_COLOR)

    m_BgColorDown = New_BgColorDown
    PropertyChanged "BgColorDown"
    Call RefreshControl

End Property

Public Property Get ButtonBgColor() As OLE_COLOR

    ButtonBgColor = m_ButtonBgColor

End Property

Public Property Let ButtonBgColor(ByVal New_ButtonBgColor As OLE_COLOR)

    m_ButtonBgColor = New_ButtonBgColor
    PropertyChanged "ButtonBgColor"
    Call RefreshControl

End Property

Public Property Get ButtonBgColorOver() As OLE_COLOR

    ButtonBgColorOver = m_ButtonBgColorOver

End Property

Public Property Let ButtonBgColorOver(ByVal New_ButtonBgColorOver As OLE_COLOR)

    m_ButtonBgColorOver = New_ButtonBgColorOver
    PropertyChanged "ButtonBgColorOver"
    Call RefreshControl

End Property

Public Property Get ButtonBgColorDown() As OLE_COLOR

    ButtonBgColorDown = m_ButtonBgColorDown

End Property

Public Property Let ButtonBgColorDown(ByVal New_ButtonBgColorDown As OLE_COLOR)

    m_ButtonBgColorDown = New_ButtonBgColorDown
    PropertyChanged "ButtonBgColorDown"
   Call RefreshControl

End Property
Public Property Get ColumnHeaders() As Boolean
    ColumnHeaders = m_ColumnHeaders
End Property

Public Property Let ColumnHeaders(ByVal New_ColumnHeaders As Boolean)
    m_ColumnHeaders = New_ColumnHeaders
    
    PropertyChanged "ColumnHeaders"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_Enabled = m_def_Enabled
    m_ShowIcon = m_def_ShowIcon
    m_IconSizeWidth = m_def_IconSizeWidth
    m_IconSizeHeight = m_def_IconSizeHeight
    m_BorderColor = m_def_BorderColor
    m_BorderColorOver = m_def_BorderColorOver
    m_BorderColorDown = m_def_BorderColorDown
    m_BgColor = m_def_BgColor
    m_BgColorOver = m_def_BgColorOver
    m_BgColorDown = m_def_BgColorDown
    m_ButtonBgColor = m_def_ButtonBgColor
    m_ButtonBgColorOver = m_def_ButtonBgColorOver
    m_ButtonBgColorDown = m_def_ButtonBgColorDown
    m_Text = m_def_Text
    m_ColumnHeaders = True
    m_ListHeight = 3070
    m_FocusColor = m_def_FocusColor
    m_DropListEnabled = True
     
    Call RefreshControl
  
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    IniLat = Height / 15
    IniLung = Width / 15
    
    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    m_ShowIcon = PropBag.ReadProperty("ShowIcon", m_def_ShowIcon)
    m_IconSizeWidth = PropBag.ReadProperty("IconSizeWidth", m_def_IconSizeWidth)
    m_IconSizeHeight = PropBag.ReadProperty("IconSizeHeight", m_def_IconSizeHeight)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_BorderColorOver = PropBag.ReadProperty("BorderColorOver", m_def_BorderColorOver)
    m_BorderColorDown = PropBag.ReadProperty("BorderColorDown", m_def_BorderColorDown)
    m_BgColor = PropBag.ReadProperty("BgColor", m_def_BgColor)
    m_BgColorOver = PropBag.ReadProperty("BgColorOver", m_def_BgColorOver)
    m_BgColorDown = PropBag.ReadProperty("BgColorDown", m_def_BgColorDown)
    m_ButtonBgColor = PropBag.ReadProperty("ButtonBgColor", m_def_ButtonBgColor)
    m_ButtonBgColorOver = PropBag.ReadProperty("ButtonBgColorOver", m_def_ButtonBgColorOver)
    m_ButtonBgColorDown = PropBag.ReadProperty("ButtonBgColorDown", m_def_ButtonBgColorDown)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_NrColVisible = PropBag.ReadProperty("NrColVisible", 1)
    m_ListHeight = PropBag.ReadProperty("ListHeight", 3070)
    m_ListWidth = PropBag.ReadProperty("ListWidth", "100")
    m_BoundColumns = PropBag.ReadProperty("m_BoundColumns", "0")
    m_FocusColor = PropBag.ReadProperty("FocusColor", m_def_FocusColor)
    m_MaxLength = PropBag.ReadProperty("TextMaxLength", 0)
    Text1.MaxLength = PropBag.ReadProperty("TextMaxLength", 0)
    Text1.Enabled = PropBag.ReadProperty("Text_Enabled", True)
    Text1.Locked = PropBag.ReadProperty("Text_Locked", False)
    m_DropListEnabled = PropBag.ReadProperty("DropListEnabled", True)
    m_ColumnHeaders = PropBag.ReadProperty("ColumnHeaders", True)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    
    Call RefreshControl

End Sub

Private Sub UserControl_Show()
   If m_SelectControl = XPMultiCombo Then
      Text1.Text = ""
   End If
   
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("ShowIcon", m_ShowIcon, m_def_ShowIcon)
    Call PropBag.WriteProperty("IconSizeWidth", m_IconSizeWidth, m_def_IconSizeWidth)
    Call PropBag.WriteProperty("IconSizeHeight", m_IconSizeHeight, m_def_IconSizeHeight)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderColorOver", m_BorderColorOver, m_def_BorderColorOver)
    Call PropBag.WriteProperty("BorderColorDown", m_BorderColorDown, m_def_BorderColorDown)
    Call PropBag.WriteProperty("BgColor", m_BgColor, m_def_BgColor)
    Call PropBag.WriteProperty("BgColorOver", m_BgColorOver, m_def_BgColorOver)
    Call PropBag.WriteProperty("BgColorDown", m_BgColorDown, m_def_BgColorDown)
    Call PropBag.WriteProperty("ButtonBgColor", m_ButtonBgColor, m_def_ButtonBgColor)
    Call PropBag.WriteProperty("ButtonBgColorOver", m_ButtonBgColorOver, m_def_ButtonBgColorOver)
    Call PropBag.WriteProperty("ButtonBgColorDown", m_ButtonBgColorDown, m_def_ButtonBgColorDown)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("TextMaxLength", m_MaxLength, 0)
    Call PropBag.WriteProperty("FocusColor", m_FocusColor, m_def_FocusColor)
    Call PropBag.WriteProperty("NrColVisible", m_NrColVisible, 1)
    Call PropBag.WriteProperty("ListHeight", m_ListHeight, 3070)
    Call PropBag.WriteProperty("ListWidth", m_ListWidth, "100")
    Call PropBag.WriteProperty("BoundColumns", m_BoundColumns, "0")
    Call PropBag.WriteProperty("ColumnHeaders", m_ColumnHeaders, True)
    Call PropBag.WriteProperty("Text_Enabled", Text1.Enabled, True)
    Call PropBag.WriteProperty("Text_Locked", Text1.Locked, False)
    Call PropBag.WriteProperty("DropListEnabled", m_DropListEnabled, True)
   
End Sub

Private Sub UserControl_Click()

  If m_Enabled Then
    RaiseEvent Click
  End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  If m_Enabled Then
    RaiseEvent KeyDown(KeyCode, Shift)
    Call DrawControl(1)
  End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  If m_Enabled Then
     
    RaiseEvent KeyPress(KeyAscii)
  End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   If m_Enabled Then
    RaiseEvent KeyUp(KeyCode, Shift)
    Call DrawControl(0)
   End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

RaiseEvent MouseDown(Button, Shift, X, Y)
Select Case m_SelectControl
  Case 0
    If (X >= ButtRect.Left And X <= ButtRect.Right) And (Y >= ButtRect.Top And Y <= ButtRect.Bottom) Then
      IsButtDown = True
      Call DrawControl(2)
    Else
       IsButtDown = False
    End If

 End Select
End Sub

Public Sub XPComboShow(Show As Integer)
 If Show = 0 And IsWindowVisible(frmpopup.hWnd) = 0 Then
    GoTo ShowDropDown_Exit
 ElseIf Show = 1 And (IsWindowVisible(frmpopup.hWnd) <> 0 Or m_DropListEnabled = False) Then
    GoTo ShowDropDown_Exit
 End If
 
  If Show Then
  
  Dim ClrPos As RECT
  Dim crx As Long
   
    Call GetWindowRect(hWnd, ClrPos)
   
    Call RefreshControl
   
  RaiseEvent DropList
  Load frmpopup
    With frmpopup
        
        .Left = ClrPos.Left * Screen.TwipsPerPixelX
        .Top = ClrPos.Bottom * Screen.TwipsPerPixelY
        .BackColor = m_BgColor
        .Selectedtext = Text1.Text
        If (.Top + .Height) > Screen.Height Then
            .Top = ClrPos.Top * Screen.TwipsPerPixelY - .Height
        End If
        
        'Check whether ColumnHeader is on or Hide
        If m_ColumnHeaders = False Then
        .Height = m_ListHeight - 270
        Else
        .Height = m_ListHeight + 10
        End If
        
        If m_ListHeight <= 2000 Then
           m_ListHeight = 2000
        End If
        
  
      .lsw.Width = iTotalCol * Screen.TwipsPerPixelY
      .lsw.Height = m_ListHeight
      .Width = .lsw.Width
      .lsw.BackColor = m_BgColor
      .isclick = False
     
        'Modal Mode
        .Show 1
        
        If .isclick Then
            Text1.Text = .Selectedtext
            
        End If
        
         Unload frmpopup
     Set frmpopup = Nothing
          
    Call SetRectEmpty(ClrPos)
       End With
        IsMDown = False
        IsCrlOver = False
        Call RefreshControl
        
       End If

ShowDropDown_Exit:
    Exit Sub
End Sub

'## ControlType Select
Public Property Get SelectControl() As UserControlType
    SelectControl = m_SelectControl
End Property

Public Property Let SelectControl(ByVal New_SelectControl As UserControlType)
    m_SelectControl = New_SelectControl
    PropertyChanged "SelectControl"
End Property
' Font
Public Property Get Font() As StdFont
    Set Font = Text1.Font
End Property
Public Property Set Font(New_Font As StdFont)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
   Call RefreshControl
End Property
Public Property Get TextMaxLength() As Integer
   TextMaxLength = m_MaxLength
End Property

Public Property Let TextMaxLength(ByVal New_TextMaxLength As Integer)
   
   m_MaxLength = New_TextMaxLength
   
   Text1.MaxLength = New_TextMaxLength
   
   PropertyChanged "TextMaxLength"
   
End Property
Public Property Get Text() As String

    Text = m_Text

End Property

Public Property Let Text(ByVal New_Text As String)
    If m_SelectControl = XPMultiCombo Then
       Text1.Text = Left(New_Text, m_MaxLength)
    Else
       Text1.Text = New_Text
    End If
  
    PropertyChanged "Text"

End Property

Public Property Get Text_Enabled() As Boolean
    Text_Enabled = Text1.Enabled
End Property
Public Property Let Text_Enabled(ByVal Text_New_Enabled As Boolean)
    Text1.Enabled() = Text_New_Enabled
    PropertyChanged "Text_Enabled"
End Property

Public Property Get Text_Locked() As Boolean
    Text_Locked = Text1.Locked
End Property
Public Property Let Text_Locked(ByVal Text_New_Locked As Boolean)
    Text1.Locked() = Text_New_Locked
    PropertyChanged "Text_Locked"
End Property

Public Property Get NrColVisible() As Long
    NrColVisible = m_NrColVisible
End Property

Public Property Let NrColVisible(New_NrColVisible As Long)
    m_NrColVisible = New_NrColVisible
    PropertyChanged "NrColVisible"
End Property
Public Property Get ListHeight() As Long
    ListHeight = m_ListHeight
End Property
Public Property Let ListHeight(New_ListHeight As Long)
    m_ListHeight = New_ListHeight
    PropertyChanged "ListHeight"
End Property
Public Property Get ListWidth() As String
    ListWidth = m_ListWidth
End Property
Public Property Let ListWidth(New_ListWidth As String)
    m_ListWidth = New_ListWidth
    PropertyChanged "ListWidth"
End Property
Public Property Get BoundColumns() As String
    BoundColumns = m_BoundColumns
End Property
Public Property Let BoundColumns(New_BoundColumns As String)
    m_BoundColumns = New_BoundColumns
    PropertyChanged "BoundColumns"
End Property
Public Property Get DropListEnabled() As Boolean
    DropListEnabled = m_DropListEnabled
End Property

Public Property Let DropListEnabled(ByVal New_DropListEnabled As Boolean)
    m_DropListEnabled = New_DropListEnabled
    
    PropertyChanged "DropListEnabled"
End Property

Public Property Get FocusColor() As OLE_COLOR

    FocusColor = m_FocusColor

End Property

Public Property Let FocusColor(ByVal New_FocusColor As OLE_COLOR)

    m_FocusColor = New_FocusColor
    PropertyChanged "FocusColor"
    Call RefreshControl

End Property
Public Sub Load_rs_to_lsw(ByVal lswcbo_rs As Recordset)
      
       
    Dim vbook
    Dim chk_book As Boolean
    Dim rs_opened As Boolean
    Dim col_length As Integer
    Dim itemx
    Dim I As Integer
    Dim col_turn As Integer
    Dim intCount As Integer
    
     If lswcbo_rs.State = 0 Then
            rs_opened = True
            lswcbo_rs.Open
     Else
            rs_opened = False
            
     End If
     
    '-------------------------------------------------
    '# Deal Users input worng numbers of Bounding Columns
        
        'Clear 0,numbers of Bounding Columns
        NumBounds = 0
        'Calculate how many Fields in Recordset
        intCount = lswcbo_rs.Fields.Count
        
  
        Dim lWid() As Long
        Dim substr() As String
        Dim SubStrCount As Integer
        SubStrCount = 0
        ReDim substr(0 To 10) As String
        SubStrCount = DespartireSTR(substr(), m_ListWidth, ";")
        
        Dim strsplit() As String
        Dim StrBoundColumns As Integer
        StrBoundColumns = 0
        ReDim strsplit(0 To 10) As String
        StrBoundColumns = DespartireSTR(strsplit(), m_BoundColumns, ";")
      
        Dim m As Integer
        Dim intsplit() As Integer
        ReDim intsplit(0 To 10) As Integer
                 
   '# Check whether user set visible bounding columns are
   '  over total fields (XPMCCombo1.NrColVisible = 4 but intCount is
   '  only 3)
    If m_NrColVisible > 0 Then
     If m_NrColVisible >= intCount Then
     NumBounds = intCount
     Else
     NumBounds = m_NrColVisible
     End If
    Else
     NumBounds = 1
    End If
    
   '# Check whether user set visible bounding columns are
   '  over total fields (XPMCCombo1.ListWidth = "200;1800;1000;1000",StrBoundColumns=4
   '  but intCount has only 3,intCount=3)
    If StrBoundColumns > 0 Then
        If StrBoundColumns >= intCount Then
           StrBoundColumns = intCount
        End If
    Else
        StrBoundColumns = 1
    End If
    
   '# converter string to interger
    For m = 1 To StrBoundColumns
        intsplit(m) = CInt(Val(strsplit(m)))
       '# Override the fault when user input illegal format
       '  such as XPMCCombo1.BoundColumns = "2;0;4;" but actual total fields only
       '  have 3(that is:intcount=3)
        If intsplit(m) > intCount Then intsplit(m) = intCount
    Next m
        
    
    If NumBounds >= StrBoundColumns Then
        NumBounds = StrBoundColumns
    End If
             
        Dim iCt As Integer
        iCt = NumBounds - 1
        
        ReDim lWid(0 To iCt)
        
        For I = 1 To iCt
                
                lWid(I) = Val(substr(I + 1))
                
        Next
              
                
  '-------------------------------------------------
    With lswcbo_rs
        
        If Check_bookmarkable(lswcbo_rs) = True Then
            chk_book = True
            vbook = .Bookmark
        Else
            chk_book = False
        End If
        
        
                
        frmpopup.lsw.ColumnHeaders.Clear
                
        
         
        For I = 0 To iCt
            If I <> 0 Then
             frmpopup.lsw.ColumnHeaders.Add , , .Fields(intsplit(I + 1)).Name, lWid(I)
                
            Else
            '# First Column
             frmpopup.lsw.ColumnHeaders.Add , , .Fields(intsplit(1)).Name, IniLung
            
            End If
        Next
        
        frmpopup.lsw.ListItems.Clear
        
        .MoveFirst
        Do Until .EOF
            If IsNull(.Fields(intsplit(1))) = True Then
              Set itemx = frmpopup.lsw.ListItems.Add(, , "-")
            Else
              If Len(Trim$(.Fields(intsplit(1)))) = 0 Then
                 Set itemx = frmpopup.lsw.ListItems.Add(, , "-")
              Else
                   Set itemx = frmpopup.lsw.ListItems.Add(, , .Fields(intsplit(1)))
              End If
             End If
                If iCt > 0 Then
                   Dim h As Integer

                    For h = 1 To iCt
                      If IsNull(.Fields(intsplit(h + 1))) = True Then
                         itemx.SubItems(h) = "-"
                      Else
                         If Len(Trim$(.Fields(intsplit(h + 1)))) = 0 Then
                            itemx.SubItems(h) = "-"
                         Else
                           itemx.SubItems(h) = .Fields(intsplit(h + 1))
                         End If
                      End If
                    
                    Next
  
                End If
                
            .MoveNext
        Loop
       
        '---------------------------------------------------
        If chk_book = True Then .Bookmark = vbook
        If rs_opened = True Then .Close
        '---------------------------------------------------
        End With
        
    lCol = ListView_GetColumnWidth(frmpopup.lsw.hWnd, 0)
    iSubCol = 0
    If NumBounds >= 2 Then
       Dim N As Long
       For N = 1 To NumBounds - 1
           iSubCol = ListView_GetColumnWidth(frmpopup.lsw.hWnd, N) + iSubCol
       Next
    Else
           iSubCol = 0
    End If
   
 
    'If isHScrollbarvisible = True Then
    If NoOfRecs(lswcbo_rs) <= 13 And NoOfRecs(lswcbo_rs) > 0 Then
       iTotalCol = iSubCol + lCol 'Do not get VScrollbar
    'Else
    ElseIf NoOfRecs(lswcbo_rs) > 13 Then
       iTotalCol = iSubCol + lCol + 16 'Add 16 if got VScrollbar
     End If
  
End Sub

Private Function DespartireSTR(SubStrs() As String, ByVal SrcStr As String, _
                               ByVal Delimiter As String) As Integer
      ReDim SubStrs(0) As String
      Dim CurPos As Long
      Dim NextPos As Long
      Dim DelLen As Integer
      Dim nCount As Integer
      Dim TStr As String
      CurPos = 0
      NextPos = 0
      DelLen = 0
      nCount = 0
      TStr = ""
      SrcStr = Delimiter & SrcStr & Delimiter
      DelLen = Len(Delimiter)
      nCount = 0
      CurPos = 1
      NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)
      Do Until NextPos = 0
         TStr = Mid$(SrcStr, CurPos + DelLen, NextPos - CurPos - DelLen)
         nCount = nCount + 1
         ReDim Preserve SubStrs(nCount) As String
         SubStrs(nCount) = TStr
         CurPos = NextPos
         NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)
      Loop

      DespartireSTR = nCount
      
   End Function

Private Function Check_bookmarkable(chk_rs As Recordset) As Boolean
    If chk_rs.EOF = True Or chk_rs.BOF = True Then Check_bookmarkable = False Else Check_bookmarkable = True
End Function

Private Function NoOfRecs(Rs As Recordset) As Long

On Error Resume Next
    If Rs Is Nothing Then
        NoOfRecs = 0
    Else
        NoOfRecs = Rs.RecordCount
    End If

End Function

