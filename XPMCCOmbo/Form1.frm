VERSION 5.00
Object = "*\APJDMMCCombo.vbp"
Begin VB.Form Form1 
   Caption         =   "Dropable Multi-Columns Combo Test Demo"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   568
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check5 
      Caption         =   "Show Icon"
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   600
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Text            =   "2"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Text            =   """200;80;200;200"""
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Text            =   """0;2;1;3"""
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CheckBox Check4 
      Caption         =   "DPMCCombo1.Text_Locked"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CheckBox Check3 
      Caption         =   "DPMCCombo1.Text_Enabled"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   1560
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      Caption         =   "DPMCCombo1.Enabled"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1200
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "DPMCCombo1.DropListEnabled"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   840
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin MCCombo.DPMCCombo DPMCCombo2 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      Icon            =   "Form1.frx":0000
      NrColVisible    =   2
      ListWidth       =   "200;80;200;200"
   End
   Begin MCCombo.DPMCCombo DPMCCombo1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      Icon            =   "Form1.frx":015A
      NrColVisible    =   2
      ListWidth       =   "200;80;200;200"
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1.Text = DPMCCombo1.Text"
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "DPMCCombo2:"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "DPMCCombo1.NrColVisible"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "DPMCCombo1.ListWidth (Seperated by "";"")"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "DPMCCombo1.BoundColumns (Seperated by "";"")"
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "DPMCCombo1:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "DPMCCombo1_Change()"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
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
Dim cnCountry As ADODB.Connection
Dim rsCountry As ADODB.Recordset

Private Sub Check1_Click()
 DPMCCombo1.DropListEnabled = Check1.Value
End Sub

Private Sub Check2_Click()
DPMCCombo1.Enabled = Check2.Value
End Sub

Private Sub Check3_Click()
DPMCCombo1.Text_Enabled = Check3.Value
    
End Sub

Private Sub Check4_Click()
 DPMCCombo1.Text_Locked = Check4.Value
End Sub
Private Sub Check5_Click()
 DPMCCombo1.ShowIcon = Check5.Value
End Sub
Private Sub DPMCCombo1_Change()
    Text1.Text = DPMCCombo1.Text
End Sub

Private Sub Form_Load()

  Dim strFilespec As String
  Dim strConn As String
  Dim strDbName As String
    strDbName = "country.mdb"                      'database name to use
    strFilespec = App.Path & "\" & strDbName
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilespec & ";"
    Set cnCountry = New ADODB.Connection       'use ADO connection
    Set rsCountry = New ADODB.Recordset        'and ADO recordset
    If cnCountry.State = adStateClosed Then    'newly-created db may be open so
        cnCountry.CursorLocation = adUseClient 'create client side
        cnCountry.Open strConn                 'check first before opening
    End If
    '~~ First Field of Recordset is Bounding Column
    rsCountry.Open "SELECT [Full Name],[Currency],[Currency Code],[Capital] FROM country ORDER BY [Full Name]", cnCountry, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    '~~ The ADODB connection is still opening.You can do other job here...
    Set rsCountry.ActiveConnection = Nothing
    cnCountry.Close
    
    '## Example
    Text2.Text = "0;2;1;3"
    Text3.Text = "200;50;120;80"
    Text4.Text = 2
        
End Sub

Private Sub DPMCCombo1_DropList()
    DPMCCombo1.BoundColumns = CStr(Text2.Text)
    '~~ the width of First Column is determined by XPCalendar.width.You can put
       'Whatever value you like.
   
    DPMCCombo1.ListWidth = CStr(Text3.Text)
    DPMCCombo1.NrColVisible = Val(Text4.Text)
    DPMCCombo1.load_rs_to_lsw rsCountry
    
   '~~ Check if recordset is EOF
   If Not rsCountry.EOF Then
       '~~ Load the recordset in to the XPMCCombo
       Me.DPMCCombo1.load_rs_to_lsw rsCountry
    
   Else
    '~~ Return to combo to normal (usable) state
       SendKeys "{ENTER}"
   End If
   
End Sub
Private Sub DPMCCombo2_DropList()
    DPMCCombo2.BoundColumns = "0;1;2;3"
    '~~ the width of First Column is determined by XPCalendar.width.You can put
        'Whatever value you like.
   
    DPMCCombo2.ListWidth = "200;120;80;200"
    DPMCCombo2.NrColVisible = 3
    DPMCCombo2.load_rs_to_lsw rsCountry
    
   '~~ Check if recordset is EOF
   If Not rsCountry.EOF Then
       '~~ Load the recordset in to the XPMCCombo
       Me.DPMCCombo2.load_rs_to_lsw rsCountry
    
   Else
    '~~ Return to combo to normal (usable) state
       SendKeys "{ENTER}"
   End If
End Sub
