Attribute VB_Name = "Globals"
Option Explicit

Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryW" (ByVal lpPathName As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)

Type tMediaFormat
    formatCode As String
    ext As String
    resolution As String
    iType As Long
    quality As String
    bitrate As String
    Size As String
    txt As String
    url As String
End Type

Public gDownloader As String
Public gPlayer As String
Public gaPlayers
Public gPlayerIndex As Long
Public gaFormats(100) As tMediaFormat
Public gFmtCount As Long
Public gdlg As CmnDialogEx
Public gProc As Long
Public gTemp As String
Public jData As JSon

Sub Main()
Dim s As String
InitCommonControls
Set gdlg = New CmnDialogEx
gdlg.CancelError = False
s = GetAppPath(0)
s = ExtractPath(s)
gDownloader = AttachPath("youtube-dl.exe", s)
If Not IsFile(gDownloader) Then
    gDownloader = "D:\youtube-dl\youtube-dl.exe"
    If IsFile(gDownloader) = False Then MsgBox "תוכנת ההורדה לא נמצאה בתיקיית התוכנה", vbCritical: Exit Sub
End If
gDownloader = Quote(gDownloader)
gPlayer = GetSetting("VideoDownloader", "Settings", "Player", "")
If gPlayer <> "" Then
    gaPlayers = Split(gPlayer, "|")
    gPlayerIndex = Val(GetSetting("VideoDownloader", "Settings", "PlayerIndex", "0"))
    gPlayer = gaPlayers(gPlayerIndex)
End If
Form1.Show
End Sub

Function Quote(s) As String
Quote = """" & s & """"
End Function

' This function runs a cmd command, consisting of the elements of params(), and redirects the output to a temp file.
' It returns the content of the temp file as an array of lines, or Empty for an empty file
' It also sets LastDllError to the process's exit code
' The file assumes to be in UTF-8 format, or ANSI if no UTF-8 compliance. BOM is also supported for UTF-8 or UTF-16LE
' The file lines can be in either unix or windows format
' It has a "async mode" in which the command is running without waiting
' To enter async mode, set the first param to "|"
' In async mode, the temp file name is stored in gTemp and the return value is the cmd process.
' When the process has finished, call this function with no params to get the resulting output and close the process
Function GetCmdOutput(ParamArray params())
Dim sCmd As String
Dim sTemp As String
Dim s As String
Dim fh As Integer
Dim a() As Byte
Dim i As Long
Dim lf As Long
Dim bAsync As Boolean
Dim ExitCode
Dim l As Long, u As Long

l = LBound(params)
u = UBound(params)
For i = l To u
    If i = l And params(i) = "|" Then
        bAsync = True
    ElseIf LenB(sCmd) Then
        sCmd = sCmd & (" " & params(i))
    Else
        sCmd = params(i)
    End If
Next
If sCmd <> "" Then
    sTemp = CreateTempFile
    sCmd = sCmd & ">" & Quote(sTemp)
    If bAsync Then gTemp = sTemp: GetCmdOutput = ShellEx("cmd", 0, "/c " & Quote(sCmd), 0, False): Exit Function
    ShellEx "cmd", 0, "/c " & Quote(sCmd), ExitCode:=ExitCode
Else
    sTemp = gTemp
    If gProc Then
        If GetExitCodeProcess(gProc, lf) Then ExitCode = lf
        CloseHandle gProc
        gProc = 0
    End If
End If
GetCmdOutput = LoadTextFile(sTemp)
Kill sTemp
If IsEmpty(ExitCode) = False Then SetLastError ExitCode
End Function

Sub GetFormatListX(sUrl As String)
Dim aContent
Dim i As Long, n As Long
Dim s As String
Dim p As Long
Dim s2 As String
Dim z As tMediaFormat

If gProc Then WaitForSingleObject gProc, -1: GetCmdOutput
gFmtCount = 0
gaFormats(0) = z
aContent = GetCmdOutput("|", gDownloader, "--get-title", "--get-duration", Quote(sUrl))
If IsEmpty(aContent) = False Then gProc = aContent
'If IsEmpty(aContent) = False Then
'    gaFormats(0).ext = aContent(0)
'    gaFormats(0).Size = aContent(1)
'End If

aContent = GetCmdOutput(gDownloader, "-F", Quote(sUrl))
If IsEmpty(aContent) Then Exit Sub
n = UBound(aContent)
For i = 0 To n
    s = aContent(i)
    If Len(s) = 0 Then Exit For
    If Left(s, 1) <> "[" Then gaFormats(0).txt = s: Exit For
Next
For i = i + 1 To n
    s = aContent(i)
    If Len(s) = 0 Then Exit For
    gFmtCount = gFmtCount + 1
    gaFormats(gFmtCount) = z
    gaFormats(gFmtCount).txt = s
    If InStr(1, s, "audio only") Then
        gaFormats(gFmtCount).iType = 1
    ElseIf InStr(1, s, "video only") Then
        gaFormats(gFmtCount).iType = 2
    Else
        gaFormats(gFmtCount).iType = 0
    End If
    p = InStr(1, s, " ")
    gaFormats(gFmtCount).formatCode = Left(s, p - 1)
    'gaFormats(gFmtCount).ext = Trim(Mid(s, 14, 8))
    'If gaFormats(gFmtCount).iType <> 1 Then gaFormats(gFmtCount).resolution = Trim(Mid(s, 25, 10)) Else gaFormats(gFmtCount).resolution = "-"
    'gaFormats(gFmtCount).quality = Trim(Mid(s, 36, 4))
    'gaFormats(gFmtCount).bitrate = Trim(Mid(s, 40, 7))
    'p = InStrRev(s, ",")
    's2 = Mid(s, p + 1)
    'gaFormats(gFmtCount).Size = Trim(s2)
Next
End Sub

Function MS2Dur(ByVal ms As Long) As String
Dim h As Long, m As Long, s As Long
Dim d As String
h = ms \ 3600000
If h Then d = CStr(h) & ":": ms = ms - h * 3600000
m = ms \ 60000
If h Then d = d & Format(m, "00") & ":" Else d = d & CStr(m) & ":"
If m Then ms = ms - m * 60000
s = ms \ 1000
d = d & Format(s, "00")
MS2Dur = d
End Function

Function FormatSize(ByVal iSize As Double) As String
Dim s As String
Dim dig As Integer
If iSize = 0 Then FormatSize = "0 Bytes": Exit Function
dig = Int(Log(iSize) / Log(2))
If dig < 10 Then
    s = iSize & " Bytes"
ElseIf dig < 20 Then
    s = NumFormat(iSize / 2 ^ 10, , , 2) & "KiB"
ElseIf dig < 30 Then
    s = NumFormat(iSize / 2 ^ 20, , , 2) & "MiB"
Else
    s = NumFormat(iSize / 2 ^ 30, , , 2) & "GiB"
End If
FormatSize = s
End Function

Function NZ(v, Optional nv)
If IsNull(v) = False Then NZ = v Else If Not IsMissing(nv) Then NZ = nv
End Function

Function Pad(s, ByVal num As Long) As String
If Len(s) >= Abs(num) Then Pad = s: Exit Function
If num > 0 Then Pad = s & String(num - Len(s), 32) Else Pad = String(-num - Len(s), 32) & s
End Function

'Sub jasona()
'Dim j As New JSon
'Dim s As String
'Dim s2 As String
'Dim p As Long
'j.test
'End Sub
's = String(1000000, 65)
'p = 2
'Mid(s, 1, 1) = """"
'Mid(s, 1000000, 1) = """"
'Lapse
''s2 = j.ReadJsonStr(s, p)
'j.Load s
'Debug.Print Lapse(1)
's = "abcdefg\n012\""345\u05d0end"",qwert"
'p = 1
's2 = j.ReadJasonStr(s, p)
''j.test
'End Sub
''
'Sub jason()
'Dim jsn As New JSon
'Dim a
'gTemp = "D:\youtube-dl\x.json"
'gTemp = "F:\Users\Chezi\AppData\Local\Temp\5E0A.tmp"
'FileCopy gTemp, gTemp & ".tmp"
'gTemp = gTemp & ".tmp"
'a = GetCmdOutput
'jsn.Load (a(0))
'End Sub

