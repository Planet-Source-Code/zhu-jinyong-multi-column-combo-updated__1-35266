VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmpopup 
   Appearance      =   0  'Flat
   BackColor       =   &H00DAE0E4&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   ScaleHeight     =   127
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   191
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView lsw 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "frmpopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Email Me: zhujy@samling.com.my
'Welcome to visit our company WebSite at:
  'www.Samling.com.my
  'www.Samling.com.cn
'Samling Group---One of a leading and largest WoodBased Industry company in Asia
'Copyright only credit to Author

Option Explicit

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long

Public Selectedtext    As String
Public isclick         As Boolean
Dim Text               As String
Dim crlClick           As Boolean

Private Sub Form_Activate()
     'isHScrollbarvisible = IsScrollbarVisible(frmpopup.lsw.hWnd, 2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           '## Escape
        Case 27
        ReleaseCapture
        Unload Me
       Set frmpopup = Nothing
          
    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (GetCapture() <> Me.hwnd) Then
        Call SetCapture(Me.hwnd)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim IsMouseOver1 As Boolean
    IsMouseOver1 = False
    IsMouseOver1 = X >= 0 And Y >= 0 And X <= ScaleWidth And Y <= ScaleHeight
    
    If IsMouseOver1 Then
        SetCapture Me.hwnd
      Else
        ReleaseCapture
        lsw.Visible = False
        Call Form_KeyDown(vbKeyEscape, 0)
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
 
    If HideColumnHeaders = True Then
       lsw.HideColumnHeaders = True
    Else
       lsw.HideColumnHeaders = False
    End If
  
    Call LSWNewStyle(Me.lsw.hwnd)
  
    lsw.Visible = True
    crlClick = False
End Sub

Private Sub Form_Resize()

  lsw.Move 0, 0, ScaleWidth, ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Call ReleaseCapture
        Unload Me
   Set frmpopup = Nothing

End Sub

Private Sub lsw_DblClick()
 
  Selectedtext = Text
  Call ReleaseCapture
  Unload Me
  Set frmpopup = Nothing
 
End Sub

Private Sub lsw_ItemClick(ByVal Item As ComctlLib.ListItem)
    crlClick = True
    isclick = True
    
    Text = Item.Text
    'Check whether Record is not set or "Null"
    If Text = "-" Then Text = ""
End Sub


Private Sub lsw_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then Selectedtext = Text: lsw.Visible = False: Unload Me
   If (KeyCode = vbKeyEscape) Then
        ReleaseCapture
        Unload Me
   End If
End Sub


