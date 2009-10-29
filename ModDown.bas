Attribute VB_Name = "ModDown"
Option Explicit
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'定义保存下载任务信息的类型
Public Type DownInfo
    mIndex As Integer       '位置索引
    mUrl As String          '地址
    mFile As String         '文件名
    mSize As Long           '文件的大小
    mGetSize As Long        '已获得的字节数
    mUseProxy As Boolean    '是否使用proxy
    mProxy As String        '使用的代理服务器地址
    mProxyPort As Integer   '代理服务器端口
    mProxyId As String      '代理服务器的用户名
    mProxyPass As String    '代理服务器的密码
End Type

'用于把下载任务信息读写到随机文件的类型
Public Type DownInfoSave
    mIndex As Integer       '位置索引
    mUrl As String * 180    '地址
    mFile As String * 50    '文件名
    mSize As Long           '文件的大小
    mGetSize As Long        '已获得的字节数
    mUseProxy As Boolean    '是否使用proxy
    mProxy As String * 50   '使用的代理服务器地址
    mProxyPort As Integer   '代理服务器端口
    mProxyId As String * 20 '代理服务器的用户名
    mProxyPass As String * 20 '代理服务器的密码
End Type

'记录每个下载任务信息的变量数组
Public mDownInfo() As DownInfo

'当前正在下载的任务的索引
Public CurrentDown(3) As Integer

'当前在ListView中选择的任务的索引
Public SelectDown As Integer

'Base64编码函数：Base64Encode
'Instr1     编码前字符串
'Outstr1    编码后字符串
Public Function Base64Encode(InStr1 As String, OutStr1 As String)
Dim mInByte(3) As Byte, mOutByte(4) As Byte
Dim myByte As Byte
Dim i As Integer, LenArray As Integer, j As Integer
Dim myBArray() As Byte
myBArray() = StrConv(InStr1, vbFromUnicode)
LenArray = UBound(myBArray) + 1
For i = 0 To LenArray Step 3
    If LenArray - i = 0 Then
        Exit For
    End If
    If LenArray - i = 2 Then
        mInByte(0) = myBArray(i)
        mInByte(1) = myBArray(i + 1)
        Base64EncodeByte mInByte, mOutByte, 2
    ElseIf LenArray - i = 1 Then
        mInByte(0) = myBArray(i)
        Base64EncodeByte mInByte, mOutByte, 1
    Else
        mInByte(0) = myBArray(i)
        mInByte(1) = myBArray(i + 1)
        mInByte(2) = myBArray(i + 2)
        Base64EncodeByte mInByte, mOutByte, 3
    End If
    For j = 0 To 3
        OutStr1 = OutStr1 & Chr(mOutByte(j))
    Next j
Next i
End Function

Private Sub Base64EncodeByte(mInByte() As Byte, mOutByte() As Byte, Num As Integer)
Dim tByte As Byte
Dim i As Integer
If Num = 1 Then
    mInByte(1) = 0
    mInByte(2) = 0
ElseIf Num = 2 Then
    mInByte(2) = 0
End If
tByte = mInByte(0) And &HFC
mOutByte(0) = tByte / 4
tByte = ((mInByte(0) And &H3) * 16) + (mInByte(1) And &HF0) / 16
mOutByte(1) = tByte
tByte = ((mInByte(1) And &HF) * 4) + ((mInByte(2) And &HC0) / 64)
mOutByte(2) = tByte
tByte = (mInByte(2) And &H3F)
mOutByte(3) = tByte
For i = 0 To 3
    If mOutByte(i) >= 0 And mOutByte(i) <= 25 Then
        mOutByte(i) = mOutByte(i) + Asc("A")
    ElseIf mOutByte(i) >= 26 And mOutByte(i) <= 51 Then
        mOutByte(i) = mOutByte(i) - 26 + Asc("a")
    ElseIf mOutByte(i) >= 52 And mOutByte(i) <= 61 Then
        mOutByte(i) = mOutByte(i) - 52 + Asc("0")
    ElseIf mOutByte(i) = 62 Then
        mOutByte(i) = Asc("+")
    Else
        mOutByte(i) = Asc("/")
    End If
Next i
If Num = 1 Then
    mOutByte(2) = Asc("=")
    mOutByte(3) = Asc("=")
ElseIf Num = 2 Then
    mOutByte(3) = Asc("=")
End If
End Sub

'编码函数：EncodeStr
'Str1   编码前字符串
'返回值 编码后字符串
Public Function EncodeStr(Str1 As String) As String
    Dim OutStr1 As String
    Call Base64Encode(Str1, OutStr1)
    EncodeStr = OutStr1
End Function
