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
'��ʾ��������һ������
Public bBusy As Boolean
'Ҫ���ص�url��ַ
Public DownUrl As String
'ʹ�õ�Winsock������
Public WhichSocket As Integer
'�������������
Public WhichDown As Integer
'���յ����������ݣ��ַ������ͣ�
Public ReceiveData As String
'���صĿ�ʼ����ʱ��
Public StartTime As Date
'�����ص��ļ������ֽ���
Public ReceiveBytes As Long
'�����ļ�����
Public mFlen As Long
'�û��Ƿ�ȡ������
Public bCancel As Boolean
'������������ַ�Ͷ˿�
Public mProxy As String
Public mProxyPort As Integer
'��������������֤�ʺż�����
Public mProxyId As String
Public mProxyPass As String
'������ļ�·��
Public mFile As String
'���ӵ��������Ͷ˿�
Private mHost As String
Private mPort As Integer
'���ص����URL
Private mRelativeUrl As String

'�������ص�URL
Public Function AnalyzeUrl() As Boolean
Dim pos1, pos2 As Integer
Dim mUrl As String
mUrl = DownUrl
If InStr(1, mUrl, "http://") > 0 Then
    '�õ��˿ں�
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
    '�õ�������ַ
    mHost = Mid(mUrl, 8, pos2 - 8)
    '�õ����·��
    mRelativeUrl = Mid(mUrl, pos2)
End If
pos2 = InStrRev(mUrl, "/")
If pos2 > 8 Then
    '�õ��ļ���
    mFile = Mid(mUrl, pos2 + 1)
Else
    AnalyzeUrl = False
    Exit Function
End If
AnalyzeUrl = True
End Function

'���ݴ���������ʹ�ò�ͬ�ĺ������ӷ�����
'���ӳɹ�����true�����򷵻�false
Public Function StartDown() As Boolean
bBusy = True
If mProxy <> "" And mProxyPort > 0 Then
    'ʹ�ô�������������
    StartDown = StartDownProxy()
Else
    'ֱ������
    StartDown = StartDownNoProxy()
End If
End Function

'ֱ������Urlָ���ķ���������
Public Function StartDownNoProxy() As Boolean
StartTime = Time()
'����Winsock���Բ����ӷ�����
frmDown.Wsock(WhichSocket).RemoteHost = mHost
frmDown.Wsock(WhichSocket).RemotePort = mPort
frmDown.Wsock(WhichSocket).Connect
'ʹ��ѭ���ȴ����ӷ������ɹ�
Do While frmDown.Wsock(WhichSocket).State <> sckConnected
    DoEvents: DoEvents: DoEvents: DoEvents
    '����ʱ�䳬��20���ȡ�����أ��˳��ù��̲�����false
    If DateDiff("s", StartTime, Time()) > 20 Or bCancel = True Or bBusy = False Then
        frmDown.CloseSocket WhichSocket, "���ӷ�����ʱ�����"
        StartDownNoProxy = False
        Exit Function
    End If
Loop
'����������������ļ�����
Dim Getstr As String
Getstr = Getstr & "GET " & mRelativeUrl & " HTTP/1.1" & vbCrLf
Getstr = Getstr & "Accept: */*" & vbCrLf
Getstr = Getstr & "Accept -Language: zh -cn" & vbCrLf
Getstr = Getstr & "Accept -Encoding: gzip , deflate" & vbCrLf
Getstr = Getstr & "User-Agent: DownJet1.0" & vbCrLf
Getstr = Getstr & "Host: " & mHost & vbCrLf
If mFlen > 0 Then
    '�����ǰ�Ѿ�������һ�������ݣ����Ͷϵ���������
    Getstr = Getstr & "Range: bytes=" & ReceiveBytes & "-" & vbCrLf
End If
Getstr = Getstr & "Connection: close" & vbCrLf
Getstr = Getstr & vbCrLf
frmDown.Wsock(WhichSocket).SendData Getstr
StartDownNoProxy = True
End Function

'ͨ��������������������
Public Function StartDownProxy() As Boolean
StartTime = Time()
'����winsock���Բ����Ӵ���������
frmDown.Wsock(WhichSocket).RemoteHost = mProxy
frmDown.Wsock(WhichSocket).RemotePort = mProxyPort
frmDown.Wsock(WhichSocket).Connect
Do While frmDown.Wsock(WhichSocket).State <> sckConnected
    DoEvents: DoEvents: DoEvents: DoEvents
    If DateDiff("s", StartTime, Time()) > 10 Or bCancel = True Or bBusy = False Then
        frmDown.CloseSocket WhichSocket, "���Ӵ���������ʱ�����"
        StartDownProxy = False
        Exit Function
    End If
Loop
'Mozilla/4.0 (compatible; MSIE 4.01; Windows 98)
'��������������������ļ�����
Dim Getstr As String
Getstr = Getstr & "GET " & DownUrl & " HTTP/1.1" & vbCrLf
Getstr = Getstr & "Accept: */*" & vbCrLf
Getstr = Getstr & "Accept -Language: zh -cn" & vbCrLf
Getstr = Getstr & "Accept -Encoding: gzip , deflate" & vbCrLf
Getstr = Getstr & "User-Agent: DownJet1.0" & vbCrLf
Getstr = Getstr & "Host: " & mHost & vbCrLf
If mProxyId <> "" Then
    '���ʹ��������֤���������뵽�����ַ�����
    Getstr = Getstr & "Proxy-Authorization: Basic " & EncodeStr(mProxyId & ":" & mProxyPass) & vbCrLf
End If
If mFlen > 0 Then
    '�����ǰ�Ѿ�������һ�������ݣ����Ͷϵ���������
    Getstr = Getstr & "Range: bytes=" & ReceiveBytes & "-" & vbCrLf
End If
Getstr = Getstr & "Connection: close" & vbCrLf
Getstr = Getstr & vbCrLf
frmDown.Wsock(WhichSocket).SendData Getstr
StartDownProxy = True
End Function

'����������Winsock�õ���������Ӧ������
'��ڱ���
'ByteNum��  ���յ����ݵ��ֽ���
'ByteData�� �������ݵ�Byte���͵�����
'���ڱ�����
'Flen��     �ļ�����
'��������ֵ����ʾһ����˼���ַ���
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
        '��ʾ���������ļ��ɹ�
        SaveData = "200"
    ElseIf (InStr(1, ReceiveData, "HTTP/1.0 206 ") Or InStr(1, ReceiveData, "HTTP/1.1 206")) Then
        '��ʾ����ϵ������ɹ�
        SaveData = "206"
    ElseIf (InStr(1, ReceiveData, "HTTP/1.0 404 ") Or InStr(1, ReceiveData, "HTTP/1.1 404")) Then
        '��ʾ������δ�ҵ��������Դ
        SaveData = "404"
    Else
        '�������
        SaveData = "error"
        Exit Function
    End If
    '�����������Ӧ���ַ�����ָ���ļ���С�ı����ֶΣ�ȡ���ļ���С
    If InStr(1, ReceiveData, "Content-Length:") > 0 And mFlen = 0 Then
    Dim pos1 As Long, pos2 As Long
    pos1 = InStr(1, ReceiveData, "Content-Length:")
    pos2 = InStr(pos1 + 16, ReceiveData, vbCrLf)
        If pos2 > pos1 Then
            mFlen = Mid(ReceiveData, pos1 + 16, pos2 - pos1 - 16)
            Flen = mFlen
        End If
    End If
    '�ӷ�������Ӧ���ص������в��������ļ�����ʼλ��
    For i = 0 To UBound(ByteData()) - 3
        If ByteData(i) = 13 And ByteData(i + 1) = 10 And ByteData(i + 2) = 13 And ByteData(i + 3) = 10 Then
            StartPos = i + 4
            bAppend = True
            Exit For
        End If
    Next i
End If
'���ȡ�������˳��ù��̣��������ַ�����cancel��
If bAppend = False Then
    If bCancel = True Then
        SaveData = "cancel"
    End If
    Exit Function
End If
'�ڵ���frmDown��Public����DraoDownPic��ӳ�������
frmDown.DrawDownPic WhichSocket, ByteNum - StartPos, mFlen, ReceiveBytes
ReceiveBytes = ReceiveBytes + ByteNum - StartPos
Tfile = mFile
Fnum = FreeFile()
'��������ļ��м��������ļ�������
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
