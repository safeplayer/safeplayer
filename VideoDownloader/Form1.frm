VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "הורדת וידאו"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13395
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFolder 
      Caption         =   "פתח תיקייה"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4118
      Width           =   1455
   End
   Begin VB.ComboBox cmbApps 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   5820
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   4200
   End
   Begin VB.CommandButton cmdLink 
      Caption         =   "העתק קישור"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   4200
   End
   Begin VB.CommandButton cmdPlayer 
      Caption         =   "הוסף..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   375
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   5805
      Width           =   855
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00FFFFC0&
      Caption         =   "נגן וידאו"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5880
      Width           =   2055
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      ItemData        =   "Form1.frx":0442
      Left            =   240
      List            =   "Form1.frx":0452
      TabIndex        =   16
      Top             =   4200
      Width           =   4095
   End
   Begin VB.CommandButton cmdDownload 
      BackColor       =   &H00FFFFC0&
      Caption         =   "הורד וידאו"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5880
      Width           =   2055
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   11
      Top             =   5160
      Width           =   3975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "חפש..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4515
      Width           =   855
   End
   Begin VB.TextBox txtDest 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   4560
      Width           =   7215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   12735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2415
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbType 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Form1.frx":0469
      Left            =   10200
      List            =   "Form1.frx":0476
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1020
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "טען"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   315
      Width           =   615
   End
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "תוכנת נגן:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   5880
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "מיזוג וידאו+אודיו:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11340
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "<אורך>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   660
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   375
      Left            =   255
      TabIndex        =   13
      Top             =   315
      Width           =   4575
      VariousPropertyBits=   746604571
      Size            =   "8070;661"
      Value           =   "<שם וידאו>"
      SpecialEffect   =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   177
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "שם קובץ:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   12075
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   5160
      Width           =   840
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "הורד לתיקייה:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11655
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4200
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "סוג:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   12555
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "כתובת וידאו:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11790
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   1170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbApps_Click()
gPlayerIndex = cmbApps.ListIndex
gPlayer = gaPlayers(gPlayerIndex)
SaveSetting "VideoDownloader", "Settings", "PlayerIndex", CStr(gPlayerIndex)
End Sub

Private Sub cmbType_Click()
FillFormats
End Sub

Private Sub cmdBrowse_Click()
Dim s As String
s = GetFolder(hWnd, "תיקיית הורדה", txtDest)
If s <> "" Then txtDest = s
End Sub

Private Sub cmdDownload_Click()
If List1.ListIndex < 1 Then MsgBox "נא לבחור פורמט להורדה", vbExclamation: Exit Sub
Dim sCmd As String
Dim sOut As String
'Dim s As String
Dim sFormat As String
Dim p As Long
sOut = txtDest
's = List1.List(List1.ListIndex)
'p = InStr(1, s, " ")
'sFormat = Left(s, p - 1)
p = List1.ItemData(List1.ListIndex)
sFormat = jData("formats")(p)("format_id")

If sOut <> "" Then
    If IsExist(sOut) = False Then MakeDirTree sOut
    If IsFolder(sOut) Then SetCurrentDirectory StrPtr(sOut)
End If
'If List2.ListCount > 3 Then
'    s = List2.List(1)
'    If s <> "" And List2.List(3) <> "" Then
'        p = InStr(1, s, " ")
'        sFormat = Left(s, p - 1)
'        s = List2.List(3)
'        p = InStr(1, s, " ")
'        sFormat = sFormat & "+" & Left(s, p - 1)
'        sCmd = sCmd & "--merge-output-format mp4 "
'    End If
'End If
sCmd = gDownloader & " -f " & sFormat & " "
sCmd = sCmd & "--no-part "
sOut = txtFile
If sOut <> "" Then sCmd = sCmd & "-o " & Quote(sOut) & " "
sCmd = sCmd & Quote(txtURL)
ShellEx "cmd", 1, "/c " & Quote(sCmd), 0
End Sub

Function GetLink(ByVal ix As Long) As String
Dim s As String
Dim p As Long
Dim sFormat As String
Dim aContent

GetLink = jData("formats")(ix)("url")
Exit Function

If ix < 1 Then Exit Function
If gaFormats(ix).url <> "" Then GetLink = gaFormats(ix).url: Exit Function
s = gaFormats(ix).txt
p = InStr(1, s, " ")
sFormat = Left(s, p - 1)
aContent = GetCmdOutput(gDownloader, "-g", "-f", sFormat, Quote(txtURL))
If IsEmpty(aContent) Then Exit Function
gaFormats(ix).url = aContent(0)
GetLink = gaFormats(ix).url
End Function

Private Sub cmdFolder_Click()
If txtDest.Text <> "" Then Shell "explorer """ & txtDest.Text & """", vbNormalFocus
End Sub

Private Sub cmdLink_Click()
Dim ix As Long
Dim s As String

ix = List1.ListIndex
If ix < 1 Then MsgBox "נא לבחור פורמט", vbExclamation: Exit Sub
ix = List1.ItemData(ix)
If ix < 1 Then Exit Sub
cmdLink.Caption = "מעתיק..."
s = GetLink(ix)
If s = "" Then Exit Sub
CopyStringToClipboard s
cmdLink.Caption = "הועתק!"
cmdLink.BackColor = &HC0FFC0
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
cmdLink.Caption = "העתק קישור"
cmdLink.BackColor = vbButtonFace
End Sub

Private Sub cmdLoad_Click()
Dim aContent
List1.Clear
List2.List(1) = ""
gFmtCount = 0
If List2.ListCount > 3 Then List2.List(3) = ""
TextBox1.Text = ""
lblSize = ""
'GetFormatList txtURL
aContent = GetCmdOutput(gDownloader, "-j", Quote(txtURL.Text))
If IsEmpty(aContent) Then MsgBox "טעינת נתוני וידאו נכשלה", vbExclamation: Exit Sub
Set jData = New JSon
jData.Load aContent(0)
If jData.KeyExists("formats") Then gFmtCount = jData("formats").Length
TextBox1.Text = NZ(jData("title"))
lblSize = MS2Dur(Val(NZ(jData("duration"), 0)) * 1000&)
FillFormats
End Sub

Private Sub Timer1_Timer()
Dim a
If WaitForSingleObject(gProc, 0) = 0 Then
    Timer1.Enabled = False
    a = GetCmdOutput
    If IsEmpty(a) = False Then
        TextBox1.Text = a(0)
        lblSize = a(1)
    End If
End If
End Sub

Private Sub cmdPlay_Click()
Dim aContent
Dim sFormat As String
Dim s As String, s2 As String
Dim p As Long
Dim ix As Long, ix2 As Long

ix = List2.ItemData(1)
ix2 = List2.ItemData(3)
If ix <> 0 And ix2 <> 0 Then
    If IsFile(gPlayer) = False Then cmdPlay_Click
    If IsFile(gPlayer) = False Then Exit Sub
    
    s = GetLink(ix)
    s2 = GetLink(ix2)
    If s <> "" And s2 <> "" Then Shell Quote(gPlayer) & " " & Quote(s) & " " & Quote(s2), vbNormalFocus
    Exit Sub
End If

ix = List1.ListIndex
If ix < 1 Then MsgBox "נא לבחור פורמט להורדה", vbExclamation: Exit Sub
If IsFile(gPlayer) = False Then cmdPlay_Click
If IsFile(gPlayer) = False Then Exit Sub

ix = List1.ItemData(ix)
If ix < 1 Then Exit Sub
s = GetLink(ix)
If s <> "" Then Shell Quote(gPlayer) & " " & Quote(s), vbNormalFocus
End Sub

Private Sub cmdPlayer_Click()
Static once As Boolean
If once = False Then
    once = True
    gdlg.InitDir = GetSpecialFolder(CSIDL_PROGRAM_FILES)
End If
gdlg.Filter = "*.exe|*.exe"
If gdlg.ShowOpen(hWnd) = False Then Exit Sub
If IsEmpty(gaPlayers) Then
    gaPlayers = Array(gdlg.FileName)
Else
    ReDim Preserve gaPlayers(0 To UBound(gaPlayers) + 1)
    gaPlayers(UBound(gaPlayers)) = gdlg.FileName
End If
gPlayer = gdlg.FileName
cmbApps.AddItem ExtractFileName(gPlayer, False)
gPlayerIndex = cmbApps.ListCount - 1
cmbApps.ListIndex = gPlayerIndex
SaveSetting "VideoDownloader", "Settings", "Player", Join(gaPlayers, "|")
SaveSetting "VideoDownloader", "Settings", "PlayerIndex", CStr(gPlayerIndex)
End Sub

Private Sub Form_Load()
cmbType.ListIndex = 0
Dim a
Dim i As Long
Dim cw As Long
txtDest = GetSetting("VideoDownloader", "Settings", "DestDir", "")
If gPlayer <> "" Then
    For i = 0 To UBound(gaPlayers)
        cmbApps.AddItem ExtractFileName(CStr(gaPlayers(i)), False)
    Next
    cmbApps.ListIndex = gPlayerIndex
End If

List2.Clear
List2.AddItem "וידאו"
List2.AddItem ""
List2.AddItem "אודיו"
List2.AddItem ""

a = Array("קובץ", "רזולוציה", "קצב סיביות", "גודל")
MSFlexGrid1.Cols = UBound(a) + 1
MSFlexGrid1.Rows = 1
cw = ScaleX(80, vbPixels, ScaleMode)
For i = 0 To UBound(a)
    MSFlexGrid1.ColWidth(i) = cw
    MSFlexGrid1.Col = i
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, i) = a(i)
Next
MSFlexGrid1.Width = cw * (UBound(a) + 1) + ScaleX(30, vbPixels, ScaleMode)
MSFlexGrid1.Left = Label2.Left + Label2.Width - MSFlexGrid1.Width
End Sub

Function GetType(ByVal ix As Long) As Long
If NZ(jData("formats")(ix)("acodec"), "none") = "none" Then
    GetType = 2
ElseIf NZ(jData("formats")(ix)("vcodec"), "none") = "none" Then
    GetType = 1
Else
    GetType = 0
End If
End Function

Sub FillFormats()
Dim i As Long
Dim t As Long
Dim r As Long
Dim vi
t = cmbType.ListIndex
Dim s As String
Dim filesize As Double
Dim dur As Double
Dim fps As Double
'MSFlexGrid1.Rows = 1
List1.Clear
'List1.AddItem "סוג-קובץ רזולוציה גודל-קובץ פרטים"
List1.AddItem "פרטים   גודל-קובץ  רזולוציה סוג-קובץ"
For i = 1 To gFmtCount
    If GetType(i - 1) = t Then
        Set vi = jData("formats")(i - 1)
        s = Pad(vi("ext"), 9)
        If t = 1 Then
            s = s & Pad("   -", 10)
        Else
            s = s & Pad(vi("width") & "x" & vi("height"), 10)
        End If
        filesize = Val(NZ(vi("filesize"), 0))
        If filesize = 0 And t = 0 Then
            On Error Resume Next
            filesize = jData("requested_formats")(0)("filesize") + jData("requested_formats")(1)("filesize")
            On Error GoTo 0
        End If
        s = s & Pad(FormatSize(filesize), 12)
        If t <> 1 Then s = s & vi("height") & "p "
        dur = Val(NZ(jData("duration"), 0))
        If filesize > 0 And dur > 0 Then s = s & Int(filesize * 8 / 1000 / dur) & "k "
        If t <> 1 Then
            fps = Val(NZ(vi("fps"), 0))
            If fps > 0 Then s = s & fps & "fps "
            s = s & vi("vcodec") & " "
        End If
        If t <> 2 Then s = s & vi("acodec") & " "
        s = s & "#" & vi("format_id") & " "
        
'        r = r + 1
'        List1.AddItem gaFormats(i).txt
        List1.AddItem s
        List1.ItemData(List1.ListCount - 1) = i - 1
'        MSFlexGrid1.Rows = r + 1
'        MSFlexGrid1.Row = r
'        MSFlexGrid1.TextMatrix(r, 0) = gaFormats(i).ext
'        MSFlexGrid1.TextMatrix(r, 1) = gaFormats(i).resolution
'        MSFlexGrid1.TextMatrix(r, 2) = gaFormats(i).bitrate
'        MSFlexGrid1.TextMatrix(r, 3) = gaFormats(i).Size
'        MSFlexGrid1.RowData(r) = gaFormats(i).formatCode
'        MSFlexGrid1.Col = 0
'        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'        MSFlexGrid1.Col = 1
'        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'        MSFlexGrid1.Col = 2
'        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'        MSFlexGrid1.Col = 3
'        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    End If
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "VideoDownloader", "Settings", "DestDir", txtDest
End Sub

Private Sub List1_DblClick()
Dim s As String
Dim n As Long
n = List1.ListIndex
If n < 1 Then Exit Sub
s = List1.List(n)
If cmbType.ListIndex = 2 Then List2.List(1) = s: List2.ItemData(1) = List1.ItemData(n)
If cmbType.ListIndex = 1 Then List2.List(3) = s: List2.ItemData(3) = List1.ItemData(n)
End Sub

Private Sub List2_DblClick()
Dim n As Long
n = List2.ListIndex
If n = 1 Or n = 3 Then List2.List(n) = "": List2.ItemData(n) = 0
End Sub

Private Sub txtDest_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Files As Collection
Set Files = GetDragFiles(Data)
If Files.Count Then txtDest = Files(1)
End Sub

Private Sub txtURL_GotFocus()
txtURL.SelStart = 0
txtURL.SelLength = Len(txtURL)
End Sub

Private Sub txtURL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0: cmdLoad_Click
End Sub

Private Sub txtURL_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim s As String
Dim p As Long, p2 As Long
s = GetHtmlDrag(Data)
'txtURL = s
p = InStr(1, s, "href=""")
If p < 1 Then Exit Sub
p = p + 6
p2 = InStr(p, s, """")
If p2 < p Then Exit Sub
s = Mid(s, p, p2 - p)
txtURL = s
End Sub
