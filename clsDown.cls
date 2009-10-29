VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'表示正在下载一个任务
Public bBusy As Boolean
'要下载的url地址
Public DownUrl As String
'使用的Winsock的索引
Public WhichSocket As Integer
'下载任务的索引
Public WhichDown As Integer
'接收到的下载数据（字符串类型）
Public ReceiveData As String
'下载的开始连接时间
Public StartTime As Date
'已下载的文件数据字节数
Public ReceiveBytes As Long
'下载文件长度
Public mFlen As Long
'用户是否取消下载
Public bCancel As Boolean
'代理服务器地址和端口
Public mProxy As String
Public mProxyPort As Integer
'代理服务器的认证帐号及密码
Public mProxyId As String
Public mProxyPass As String
'保存的文件路径
Public mFile As String
'连接的主机名和端口
Private mHost As String
Private mPort As Integer
'下载的相对URL
Private mRelativeUrl As String

'分析下载的URL
Public Function AnalyzeUrl() As Boolean
Dim pos1, pos2 As Integer
Dim mUrl As String
mUrl = DownUrl
If InStr(1, mUrl, "http://") > 0 Then
    '得到端口号
    mPort = 80
Else
    AnalyzeUrl = False
    Exit Function
End If
pos1 = InStr(1, mUrl, "http://")
pos2 = InStr(8, mUrl, "/")
If pos2 = 0 Then
    AnalyzeUrl = False
    Exit Function
Else
    '得到主机地址
    mHost = Mid(mUrl, 8, pos2 - 8)
    '得到相对路径
    mRelativeUrl = Mid(mUrl, pos2)
End If
pos2 = InStrRev(mUrl, "/")
If pos2 > 8 Then
    '得到文件名
    mFile = Mid(mUrl, pos2 + 1)
Else
    AnalyzeUrl = False
    Exit Function
End If
AnalyzeUrl = True
End Function

'根据代理的设置使用不同的函数连接服务器
'连接成功返回true，否则返回false
Public Function StartDown() As Boolean
bBusy = True
If mProxy <> "" And mProxyPort > 0 Then
    '使用代理服务器下载
    StartDown = StartDownProxy()
Else
    '直接下载
    StartDown = StartDownNoProxy()
End If
End Function

'直接连接Url指定的服务器下载
Public Function StartDownNoProxy() As Boolean
StartTime = Time()
'设置Winsock属性并连接服务器
frmDown.Wsock(WhichSocket).RemoteHost = mHost
frmDown.Wsock(WhichSocket).RemotePort = mPort
frmDown.Wsock(WhichSocket).Connect
'使用循环等待连接服务器成功
Do While frmDown.Wsock(WhichSocket).State <> sckConnected
    DoEvents: DoEvents: DoEvents: DoEvents
    '连接时间超过20秒或取消下载，退出该过程并返回false
    If DateDiff("s", StartTime, Time()) > 20 Or bCancel = True Or bBusy = False Then
        frmDown.CloseSocket WhichSocket, "连接服务器时间过长"
        StartDownNoProxy = False
        Exit Function
    End If
Loop
'向服务器发送下载文件请求
Dim Getstr As String
Getstr = Getstr & "GET " & mRelativeUrl & " HTTP/1.1" & vbCrLf
Getstr = Getstr & "Accept: */*" & vbCrLf
Getstr = Getstr & "Accept -Language: zh -cn" & vbCrLf
Getstr = Getstr & "Accept -Encoding: gzip , deflate" & vbCrLf
Getstr = Getstr & "User-Agent: DownJet1.0" & vbCrLf
Getstr = Getstr & "Host: " & mHost & vbCrLf
If mFlen > 0 Then
    '如果以前已经下载了一部分数据，发送断点续传请求
    Getstr = Getstr & "Range: bytes=" & ReceiveBytes & "-" & vbCrLf
End If
Getstr = Getstr & "Connection: close" & vbCrLf
Getstr = Getstr & vbCrLf
frmDown.Wsock(WhichSocket).SendData Getstr
StartDownNoProxy = True
End Function

'通过代理服务器连接下载
Public Function StartDownProxy() As Boolean
StartTime = Time()
'设置winsock属性并连接代理服务器
frmDown.Wsock(WhichSocket).RemoteHost = mProxy
frmDown.Wsock(WhichSocket).RemotePort = mProxyPort
frmDown.Wsock(WhichSocket).Connect
Do While frmDown.Wsock(WhichSocket).State <> sckConnected
    DoEvents: DoEvents: DoEvents: DoEvents
    If DateDiff("s", StartTime, Time()) > 10 Or bCancel = True Or bBusy = False Then
        frmDown.CloseSocket WhichSocket, "连接代理服务器时间过长"
        StartDownProxy = False
        Exit Function
    End If
Loop
'Mozilla/4.0 (compatible; MSIE 4.01; Windows 98)
'向代理服务器发送下载文件请求
Dim Getstr As String
Getstr = Getstr & "GET " & DownUrl & " HTTP/1.1" & vbCrLf
Getstr = Getstr & "Accept: */*" & vbCrLf
Getstr = Getstr & "Accept -Language: zh -cn" & vbCrLf
Getstr = Getstr & "Accept -Encoding: gzip , deflate" & vbCrLf
Getstr = Getstr & "User-Agent: DownJet1.0" & vbCrLf
Getstr = Getstr & "Host: " & mHost & vbCrLf
If mProxyId <> "" Then
    '如果使用身份验证，编码后加入到请求字符串中
    Getstr = Getstr & "Proxy-Authorization: Basic " & EncodeStr(mProxyId & ":" & mProxyPass) & vbCrLf
End If
If mFlen > 0 Then
    '如果以前已经下载了一部分数据，发送断点续传请求
    Getstr = Getstr & "Range: bytes=" & ReceiveBytes & "-" & vbCrLf
End If
Getstr = Getstr & "Connection: close" & vbCrLf
Getstr = Getstr & vbCrLf
frmDown.Wsock(WhichSocket).SendData Getstr
StartDownProxy = True
End Function

'分析并保存Winsock得到服务器响应的数据
'入口变量
'ByteNum：  接收到数据的字节数
'ByteData： 接收数据的Byte类型的数组
'出口变量：
'Flen：     文件长度
'函数返回值：表示一定意思的字符串
Public Function SaveData(ByteNum As Long, ByteData() As Byte, Flen As Long) As String
Dim Tfile As String
Dim Fnum As Integer
Static m3Byte(3) As Byte
Static bAppend As Boolean
Dim StartPos As Long
Dim i As Long
If bAppend = False Then
    ReceiveData = ReceiveData & StrConv(ByteData(), vbUnicode)
    Clipboard.SetText ReceiveData
    If (InStr(1, ReceiveData, "HTTP/1.0 200 OK") Or InStr(1, ReceiveData, "HTTP/1.1 200 OK")) Then
        '表示请求下载文件成功
        SaveData = "200"
    ElseIf (InStr(1, ReceiveData, "HTTP/1.0 206 ") Or InStr(1, ReceiveData, "HTTP/1.1 206")) Then
        '表示请求断点续传成功
        SaveData = "206"
    ElseIf (InStr(1, ReceiveData, "HTTP/1.0 404 ") Or InStr(1, ReceiveData, "HTTP/1.1 404")) Then
        '表示服务器未找到请求的资源
        SaveData = "404"
    Else
        '请求错误
        SaveData = "error"
        Exit Function
    End If
    '如果服务器响应的字符串有指定文件大小的标题字段，取得文件大小
    If InStr(1, ReceiveData, "Content-Length:") > 0 And mFlen = 0 Then
    Dim pos1 As Long, pos2 As Long
    pos1 = InStr(1, ReceiveData, "Content-Length:")
    pos2 = InStr(pos1 + 16, ReceiveData, vbCrLf)
        If pos2 > pos1 Then
            mFlen = Mid(ReceiveData, pos1 + 16, pos2 - pos1 - 16)
            Flen = mFlen
        End If
    End If
    '从服务器响应返回的数据中查找下载文件的起始位置
    For i = 0 To UBound(ByteData()) - 3
        If ByteData(i) = 13 And ByteData(i + 1) = 10 And ByteData(i + 2) = 13 And ByteData(i + 3) = 10 Then
            StartPos = i + 4
            bAppend = True
            Exit For
        End If
    Next i
End If
'如果取消，则退出该过程，并返回字符串“cancel”
If bAppend = False Then
    If bCancel = True Then
        SaveData = "cancel"
    End If
    Exit Function
End If
'在调用frmDown的Public函数DraoDownPic反映下载情况
frmDown.DrawDownPic WhichSocket, ByteNum - StartPos, mFlen, ReceiveBytes
ReceiveBytes = ReceiveBytes + ByteNum - StartPos
Tfile = mFile
Fnum = FreeFile()
'向二进制文件中加入下载文件的数据
Open Tfile For Binary Lock Write As #Fnum
If LOF(Fnum) > 0 Then
    Seek #Fnum, LOF(Fnum) + 1
End If
If StartPos > 0 Then
    For i = StartPos To UBound(ByteData())
        Put #Fnum, , ByteData(i)
    Next i
Else
    Put #Fnum, , ByteData()
End If
Close #Fnum
'If bCancel = True Then
'    SaveData = "cancel"
'End If
End Function

