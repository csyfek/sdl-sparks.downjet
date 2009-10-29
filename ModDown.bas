Attribute VB_Name = "ModDown"
Option Explicit
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'���屣������������Ϣ������
Public Type DownInfo
    mIndex As Integer       'λ������
    mUrl As String          '��ַ
    mFile As String         '�ļ���
    mSize As Long           '�ļ��Ĵ�С
    mGetSize As Long        '�ѻ�õ��ֽ���
    mUseProxy As Boolean    '�Ƿ�ʹ��proxy
    mProxy As String        'ʹ�õĴ����������ַ
    mProxyPort As Integer   '����������˿�
    mProxyId As String      '������������û���
    mProxyPass As String    '���������������
End Type

'���ڰ�����������Ϣ��д������ļ�������
Public Type DownInfoSave
    mIndex As Integer       'λ������
    mUrl As String * 180    '��ַ
    mFile As String * 50    '�ļ���
    mSize As Long           '�ļ��Ĵ�С
    mGetSize As Long        '�ѻ�õ��ֽ���
    mUseProxy As Boolean    '�Ƿ�ʹ��proxy
    mProxy As String * 50   'ʹ�õĴ����������ַ
    mProxyPort As Integer   '����������˿�
    mProxyId As String * 20 '������������û���
    mProxyPass As String * 20 '���������������
End Type

'��¼ÿ������������Ϣ�ı�������
Public mDownInfo() As DownInfo

'��ǰ�������ص����������
Public CurrentDown(3) As Integer

'��ǰ��ListView��ѡ������������
Public SelectDown As Integer

'Base64���뺯����Base64Encode
'Instr1     ����ǰ�ַ���
'Outstr1    ������ַ���
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

'���뺯����EncodeStr
'Str1   ����ǰ�ַ���
'����ֵ ������ַ���
Public Function EncodeStr(Str1 As String) As String
    Dim OutStr1 As String
    Call Base64Encode(Str1, OutStr1)
    EncodeStr = OutStr1
End Function
