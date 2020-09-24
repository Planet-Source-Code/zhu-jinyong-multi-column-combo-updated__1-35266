Attribute VB_Name = "mlistview"
Option Explicit

'## Message functions:

Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---

'#Old File Name      New File Name
'ComCtl32.ocx        MsComctl.ocx
'ComCt232.ocx        MsComct2.ocx
'ComCtl32.dll        -- This file is not needed

Const LVM_FIRST = &H1000                   '// ListView messages
Const LVM_GETHEADER = (LVM_FIRST + 31)
Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54) '// optional wParam == mask
Const LVM_SCROLL = (LVM_FIRST + 20)

Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Const LVS_EX_GRIDLINES = &H1
Const LVS_EX_FULLROWSELECT = &H20         '// applies to report mode only
Const LVS_EX_ONECLICKACTIVATE = &H40
Const LVS_EX_TWOCLICKACTIVATE = &H80
'## Header control styles

Public Const HDS_BUTTONS = &H2
Const LVS_NOSCROLL = 8192

'## Window Long indexes:
Public Const GWL_STYLE = (-16)
Const WS_HSCROLL = &H100000
Const WS_VSCROLL = &H200000
'#if (_WIN32_IE >= =&H0400)
Const LVS_EX_FLATSB = &H100
'#endif


Public Enum eIsScrollbarVisible '## Scroll visible
    [Horizontal] = 1
    [Vertical] = 2
End Enum

Public HideColumnHeaders        As Boolean
Public IniLat                   As Long
Public IniLung                  As Long
Public m_NrColVisible           As Integer ' Numbers of visible columns
Public m_ListHeight             As Long
Public m_ListWidth              As String 'example : 100;500;200 The first value will be ignored and she be considered the width of control
Public m_ColumnHeads            As Boolean
Public NumBounds                As Integer
Public m_BoundColumns           As String
Public isHScrollbarvisible      As Boolean
Public lCol, iSubCol, iTotalCol As Long

Public Function LSWNewStyle(ByVal hwnd As Long)
  Dim lStyle1 As Long
  Dim lStyle2 As Long
  Dim lStyle3 As Long
  Dim lS1 As Long
  Dim ls2 As Long
  Dim ls3 As Long
  Dim lS4 As Long
  Dim lhWnd As Long
     '# Flat ScrollBar
     lStyle3 = SendMessageByLong(hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
     ls3 = LVS_EX_FLATSB
     lStyle3 = lStyle3 Or ls3
     SendMessageByLong hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle3
     '# Full Row Select
    lStyle1 = SendMessageByLong(hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
     lS1 = LVS_EX_FULLROWSELECT
     lStyle1 = lStyle1 Or LVS_EX_TWOCLICKACTIVATE
     lStyle1 = lStyle1 And Not LVS_EX_ONECLICKACTIVATE
     lStyle1 = lStyle1 Or lS1
     SendMessageByLong hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle1
      '# Add Gridline
     lStyle2 = SendMessageByLong(hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
     ls2 = LVS_EX_GRIDLINES
     lStyle2 = lStyle2 Or LVS_EX_TWOCLICKACTIVATE
     lStyle2 = lStyle2 And Not LVS_EX_ONECLICKACTIVATE
     lStyle2 = lStyle2 Or ls2
     SendMessageByLong hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle2
     
   ' Set the Buttons mode of the ListView's header control:
   lhWnd = SendMessageByLong(hwnd, LVM_GETHEADER, 0, 0)
   If (lhWnd <> 0) Then
      lS4 = GetWindowLong(lhWnd, GWL_STYLE)
      
         lS4 = lS4 And Not HDS_BUTTONS
      
      SetWindowLong lhWnd, GWL_STYLE, lS4
   End If
   
End Function

Public Function IsScrollbarVisible(ByVal hwnd As Long, WhichScroll As eIsScrollbarVisible) As Boolean
    'This Function routine come from Mr.Slider
    
    Dim lTest As Long
    lTest = WhichScroll * WS_HSCROLL
    IsScrollbarVisible = ((GetWindowLong(hwnd, GWL_STYLE) And lTest) = lTest)
End Function

Public Function ListView_GetColumnWidth(hwnd As Long, iCol As Long) As Long
  ListView_GetColumnWidth = SendMessage(hwnd, LVM_GETCOLUMNWIDTH, ByVal iCol, 0)
End Function
