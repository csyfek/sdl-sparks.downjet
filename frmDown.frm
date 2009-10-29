VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDown 
   BackColor       =   &H00C0C0C0&
   Caption         =   "下载引擎 - DwonJet"
   ClientHeight    =   7080
   ClientLeft      =   132
   ClientTop       =   708
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "BigFox's DownJet"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.8
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3552
      Left            =   6420
      TabIndex        =   9
      Top             =   3300
      Width           =   2232
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   2652
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmDown.frx":0000
         Top             =   780
         Width           =   2112
      End
      Begin VB.Label Label1 
         Caption         =   "=2K Bytes"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   10.8
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   540
         TabIndex        =   10
         Top             =   420
         Width           =   1332
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   192
         Left            =   180
         Shape           =   2  'Oval
         Top             =   420
         Width           =   192
      End
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   0
      Top             =   2940
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "下载任务列表"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3132
      Left            =   0
      TabIndex        =   3
      Top             =   60
      Width           =   8652
      Begin MSComctlLib.ListView LView 
         DragIcon        =   "frmDown.frx":008E
         Height          =   2472
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   8532
         _ExtentX        =   15050
         _ExtentY        =   4360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   33023
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   0
      End
      Begin VB.Label lblTishi 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   8292
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   1680
   End
   Begin MSWinsockLib.Winsock Wsock 
      Index           =   0
      Left            =   0
      Top             =   3240
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3792
      Left            =   60
      TabIndex        =   0
      Top             =   3240
      Width           =   6312
      _ExtentX        =   11134
      _ExtentY        =   6689
      _Version        =   393216
      TabHeight       =   420
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "幼圆"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "信息提示"
      TabPicture(0)   =   "frmDown.frx":04D0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtInfo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "下载情况"
      TabPicture(1)   =   "frmDown.frx":04EC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Pic(0)"
      Tab(1).Control(1)=   "Pic(1)"
      Tab(1).Control(2)=   "Pic(2)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "文件信息"
      TabPicture(2)   =   "frmDown.frx":0508
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtFile"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox Pic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   3372
         Index           =   2
         Left            =   -74940
         ScaleHeight     =   3324
         ScaleWidth      =   6144
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   6192
      End
      Begin VB.PictureBox Pic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   3372
         Index           =   1
         Left            =   -74940
         ScaleHeight     =   3324
         ScaleWidth      =   6144
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   6192
      End
      Begin VB.TextBox txtFile 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   7.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3432
         Left            =   -74940
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   300
         Width           =   6192
      End
      Begin VB.PictureBox Pic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   3372
         Index           =   0
         Left            =   -74940
         ScaleHeight     =   3324
         ScaleWidth      =   6144
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   6192
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3432
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   300
         Width           =   6192
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -120
      Top             =   2520
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDown.frx":0524
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDown.frx":0978
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDown.frx":0DCC
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDown.frx":1220
            Key             =   "error"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDown.frx":1674
            Key             =   "start"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDown.frx":1AC8
            Key             =   "file"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDown.frx":1F1C
            Key             =   "root"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDown.frx":2370
            Key             =   "open"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDown.frx":27C4
            Key             =   "stop"
         EndProperty
      EndProperty
   End
   Begin VB.Menu menufile 
      Caption         =   "文件"
      Begin VB.Menu menuadd 
         Caption         =   "加入新任务"
      End
      Begin VB.Menu menuquit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu menusetup 
      Caption         =   "设置"
      Begin VB.Menu menuDel 
         Caption         =   "删除"
      End
   End
End
Attribute VB_Name = "frmDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Xitem As ListItem
Dim mDownInfoSave As DownInfoSave

'声明正在下载的任务的类变量
Dim DownJet(1 To 2) As clsDown

Private Sub add_Click()
frmAdd.Show vbModal
End Sub

Private Sub Form_Load()

MinForm.Show
CDlg.FileName = App.Path & "\TEMP"
'加载用于下载的Winsock控件，可同时执行两个任务
Load Wsock(1)
Load Wsock(2)
'创建两个类对象，每个对象负责一个下载任务
Set DownJet(1) = New clsDown
Set DownJet(2) = New clsDown
'初始化ListView控件
LView.ColumnHeaders.Clear
LView.ColumnHeaders.Add , , "URL地址", LView.Width - 240 * Screen.TwipsPerPixelX
LView.ColumnHeaders.Add , , "大小", 80 * Screen.TwipsPerPixelX
LView.ColumnHeaders.Add , , "已下载大小", 80 * Screen.TwipsPerPixelX
LView.ColumnHeaders.Add , , "时间", 80 * Screen.TwipsPerPixelX
'从数据文件中读取下载任务的信息，加入到ListView中
Dim i
Dim Fnum As Integer
Dim mFname As String
'保存下载任务信息的文件
mFname = App.Path & "\Downjet.djt"
Fnum = FreeFile
Open mFname For Random As #Fnum Len = Len(mDownInfoSave)
i = 1
While Not EOF(Fnum)
    ReDim Preserve mDownInfo(i)
    Get #Fnum, i, mDownInfoSave
    If Not EOF(Fnum) Then
        mDownInfo(i).mFile = Trim(mDownInfoSave.mFile)
        mDownInfo(i).mGetSize = mDownInfoSave.mGetSize
        mDownInfo(i).mIndex = mDownInfoSave.mIndex
        mDownInfo(i).mProxy = Trim(mDownInfoSave.mProxy)
        mDownInfo(i).mProxyId = Trim(mDownInfoSave.mProxyId)
        mDownInfo(i).mProxyPass = Trim(mDownInfoSave.mProxyPass)
        mDownInfo(i).mProxyPort = mDownInfoSave.mProxyPort
        mDownInfo(i).mSize = mDownInfoSave.mSize
        mDownInfo(i).mUrl = Trim(mDownInfoSave.mUrl)
        mDownInfo(i).mUseProxy = mDownInfoSave.mUseProxy
        If mDownInfo(i).mGetSize + 1 < mDownInfo(i).mSize Then
            If Dir(mDownInfo(i).mFile) <> "" And mDownInfo(i).mFile <> "" Then
                mDownInfo(i).mGetSize = FileLen(mDownInfo(i).mFile)
            Else
                mDownInfo(i).mGetSize = 0
            End If
        End If
        AddUrl mDownInfo(i).mUrl, mDownInfo(i).mSize, mDownInfo(i).mGetSize, mDownInfo(i).mFile
        i = i + 1
    End If
Wend
Close Fnum
SelectDown = -1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer, j As Integer
Dim Fnum As Integer
Dim mFname As String
'往文件DownJet.djt中保存在ListView中的下载任务
mFname = App.Path & "\Downjet.djt"
If Dir(mFname) <> "" Then
    Kill mFname
End If
Fnum = FreeFile
j = 1
Open mFname For Random As #Fnum Len = Len(mDownInfoSave)
For i = 1 To UBound(mDownInfo) - 1
    If mDownInfo(i).mUrl <> "" And LView.ListItems(i).SmallIcon <> "delete" Then
        mDownInfoSave.mFile = mDownInfo(i).mFile
        mDownInfoSave.mGetSize = mDownInfo(i).mGetSize
        mDownInfoSave.mIndex = mDownInfo(i).mIndex
        mDownInfoSave.mProxy = mDownInfo(i).mProxy
        mDownInfoSave.mProxyId = mDownInfo(i).mProxyId
        mDownInfoSave.mProxyPass = mDownInfo(i).mProxyPass
        mDownInfoSave.mProxyPort = mDownInfo(i).mProxyPort
        mDownInfoSave.mSize = mDownInfo(i).mSize
        mDownInfoSave.mUrl = mDownInfo(i).mUrl
        mDownInfoSave.mUseProxy = mDownInfo(i).mUseProxy
        Put #Fnum, j, mDownInfoSave
        j = j + 1
    End If
Next i
Close Fnum
End
End Sub

Private Sub Form_Resize()
'如果主窗体最小化，显示浮动小窗体
If Me.WindowState = vbMinimized Then
    MinForm.Show
End If
End Sub

Private Sub LView_DblClick()
'如果选中的任务处于下载状态则停止，否则开始下载
Dim i As Integer
Dim mSel As Integer
mSel = LView.SelectedItem.Index
'如果选中的任务正在下载，则停止该下载
For i = 1 To 2
    If Wsock(i).State <> sckClosed And CurrentDown(i) = mSel Then
        DownJet(i).bCancel = True
        CurrentDown(i) = 0
        Exit Sub
    End If
Next i
'如果选中的已经下载完毕，显示下载完毕提示
If mDownInfo(mSel).mGetSize + 1 >= mDownInfo(mSel).mSize And mDownInfo(mSel).mSize > 0 Then
    lblTishi.Caption = mDownInfo(mSel).mFile & "已经下载完毕！！！"
    txtInfo.Text = txtInfo.Text & lblTishi.Caption & vbCrLf
    Exit Sub
End If
'检查是否有空闲的winsock
For i = 1 To 2
    If Wsock(i).State = sckClosed And CurrentDown(i) = 0 Then
        'winsock已经关闭，处于空闲状态，或者处于连接请求状态
        Dim mSel2 As Integer
        mSel2 = mSel
        Set DownJet(i) = Nothing
        Set DownJet(i) = New clsDown
        DownJet(i).DownUrl = LView.SelectedItem.Text
        '分析下载的Url是否合法
        If DownJet(i).AnalyzeUrl = False Then
            Exit For
        End If
        '如果第一次下载，选择路径保存下载文件
        If mDownInfo(mSel).mFile = "" Then
            '如果取消保存则退出该过程，取消下载
            On Error GoTo err1
            CDlg.CancelError = True
            CDlg.Flags = cdlOFNOverwritePrompt
            CDlg.FileName = DownJet(i).mFile
            CDlg.ShowSave
            DownJet(i).mFile = CDlg.FileName
        Else
            DownJet(i).mFile = mDownInfo(mSel).mFile
        End If
        Pic(i).Cls
        '下载任务索引
        DownJet(i).WhichDown = mSel
        '下载使用的Winsock索引
        DownJet(i).WhichSocket = i
        '下载的文件总长度
        DownJet(i).mFlen = mDownInfo(mSel2).mSize
        '已经下载的文件的大小
        DownJet(i).ReceiveBytes = mDownInfo(mSel2).mGetSize
        '****设置代理服务器选项
        DownJet(i).mProxy = mDownInfo(mSel2).mProxy
        DownJet(i).mProxyPort = mDownInfo(mSel2).mProxyPort
        DownJet(i).mProxyId = mDownInfo(mSel2).mProxyId
        DownJet(i).mProxyPass = mDownInfo(mSel2).mProxyPass
        '****
        LView.SelectedItem.SmallIcon = "start"
        '表明当前下载的任务索引
        CurrentDown(i) = LView.SelectedItem.Index
        '根据文件长度和已下载的文件长度在图片框画表示下载情况的圆点
        DrawDownPic i, 0, mDownInfo(CurrentDown(i)).mSize, mDownInfo(CurrentDown(i)).mGetSize
        txtInfo.Text = txtInfo.Text & "开始下载：" & mDownInfo(mSel2).mUrl & vbCrLf
        '开始下载，如果StartDown返回True表示连接服务器成功，发送请求
        If DownJet(i).StartDown() = False Then
            LView.ListItems(DownJet(i).WhichDown).SmallIcon = "error"
            lblTishi.Caption = "Wisock" & i & "连接服务器失败！！"
            txtInfo.Text = txtInfo.Text & "Wisock" & i & "连接服务器失败！！" & vbCrLf
            CurrentDown(i) = 0
        Else
            '开始下载成功，下载文件的路径保存到任务变量中
            mDownInfo(mSel2).mFile = DownJet(i).mFile
        End If
        Exit For
    Else
        lblTishi.Caption = "Wisock" & i & "已经有文件在下载了!!!"
        txtInfo.Text = txtInfo.Text & "Wisock" & i & "已经有文件在下载了!!!" & vbCrLf
    End If
Next i
err1:
End Sub

'向ListView添加Item的过程
Public Sub AddUrl(myUrl As String, Optional ByVal mSize As String, Optional ByVal mGetSize As String, Optional ByVal mFile As String)
If myUrl <> "" Then
    If Val(mGetSize) + 1 < Val(mSize) Or Val(mSize) = 0 Then
        Set Xitem = LView.ListItems.Add(, "", myUrl, "stop", "stop")
    Else
        Set Xitem = LView.ListItems.Add(, "", myUrl, "ok", "ok")
    End If
    Xitem.Tag = mFile
    Xitem.ListSubItems.Add , "size", mSize
    Xitem.ListSubItems.Add , "getsize", mGetSize
End If
End Sub

Private Sub LView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Static mIndex As Integer
SelectDown = LView.SelectedItem.Index
If SelectDown < 1 Then Exit Sub
'在txtfile中显示选中的下载任务信息
txtFile.Text = "URL:         " & mDownInfo(SelectDown).mUrl & vbCrLf
txtFile.Text = txtFile.Text & "文件名:      " & mDownInfo(SelectDown).mFile & vbCrLf
txtFile.Text = txtFile.Text & "大小:        " & mDownInfo(SelectDown).mSize & "字节" & vbCrLf
txtFile.Text = txtFile.Text & "已下载大小:  " & mDownInfo(SelectDown).mGetSize & "字节" & vbCrLf
txtFile.Text = txtFile.Text & "代理服务器:  " & mDownInfo(SelectDown).mProxy & vbCrLf
'在picturebox控件中显示当前选中的下载任务的block信息
'其中pic(0)显示选中的没有在下载的任务信息
'pic(1)和pic(2)显示第一个和第二个Winsock的下载信息
If SelectDown = CurrentDown(1) Then
    PicVisible (1)
ElseIf SelectDown = CurrentDown(2) Then
    PicVisible (2)
Else
    Pic(0).Cls
    PicVisible (0)
    DrawDownPic 0, 0, mDownInfo(SelectDown).mSize, mDownInfo(SelectDown).mGetSize
End If
mIndex = SelectDown
'如果按了鼠标右键弹出删除菜单
If Button = 2 Then
    PopupMenu menusetup
End If
End Sub

'接收拖放的信息，如果是可下载的Url，加到ListView中
Private Sub LView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Data.GetFormat(vbCFText) Then
    If vbOK = MsgBox("确定要下载" & Data.GetData(vbCFText), vbOKCancel) Then
        AddNewUrl Data.GetData(vbCFText)
    End If
End If
End Sub

Private Sub menuadd_Click()
'加入下载任务
frmAdd.Show vbModal
End Sub

Private Sub menuDel_Click()
'删除下载任务
LView.ListItems(LView.SelectedItem.Index).SmallIcon = "delete"
End Sub

Private Sub menuquit_Click()
'退出程序
Unload Me
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Static Count(1 To 2) As Integer
Static mColor(1 To 2) As Long
'定时显示下载时的状态
For i = 1 To 2
    If CurrentDown(i) > 0 Then
        LView.ListItems(CurrentDown(i)).SubItems(3) = Format(DateAdd("s", DateDiff("s", DownJet(i).StartTime, Time()), #12:00:00 AM#), "hh:mm:ss")
    End If
    If DownJet(i).bBusy = True Then
        Count(i) = Count(i) + 1
    Else
        Count(i) = 0
    End If
    If Count(i) > 6 Then
        Count(i) = 0
        If mColor(i) = vbRed Then
            mColor(i) = vbGreen
        Else
            mColor(i) = vbRed
        End If
    End If
    MinForm.FillColor = mColor(i)
    MinForm.Circle ((Count(i) + 1) * 120, i * 120 + 60), 50
Next i
End Sub

Private Sub Wsock_Close(Index As Integer)
CloseSocket Index, "winsock关闭"
End Sub

Private Sub Wsock_Connect(Index As Integer)
txtInfo.Text = txtInfo.Text & "Winsock" & Index & "与" & Wsock(Index).RemoteHostIP & "连接成功！" & vbCrLf
End Sub

Private Sub Wsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim ByteData1() As Byte
Dim ByteData2() As Byte
'根据Winsock的索引接收和保存数据
If Index = 1 Then
    '文件总长度的变量
    Dim Flen1 As Long
    '请求服务器返回的响应码
    Dim ReCode1 As String
    Wsock(Index).GetData ByteData1, vbByte
    '下载数据保存数据，如果是连接后第一次返回的数据，返回服务器的响应码
    ReCode1 = DownJet(Index).SaveData(bytesTotal, ByteData1(), Flen1)
    Select Case ReCode1
    Case "200"
        '响应码为200表示成功
        txtInfo.Text = txtInfo.Text & DownJet(Index).DownUrl & "开始下载" & vbCrLf
    Case "206"
        '响应码206表示断点续传成功
        txtInfo.Text = txtInfo.Text & DownJet(Index).DownUrl & "开始从" & DownJet(Index).mFlen & "断点续传下载" & vbCrLf
    Case "404"
        '响应码为404表示请求的下载的文件未找到
        txtInfo.Text = txtInfo.Text & DownJet(Index).DownUrl & "文件不存在" & vbCrLf
        CloseSocket Index, "文件未找到！"
    Case "error"
        '其他响应码视为错误
        CloseSocket Index, "请求时出错"
        txtInfo.Text = txtInfo.Text & DownJet(Index).DownUrl & "出错了" & vbCrLf
    Case "cancel"
        '用户取消
        CloseSocket Index, "用户取消"
        Exit Sub
    End Select
    If Flen1 > 0 Then
        '如果任务第一次下载，则保存后得到文件长度
        mDownInfo(DownJet(Index).WhichDown).mSize = Flen1
        LView.ListItems(DownJet(Index).WhichDown).SubItems(1) = Flen1
    End If
Else
    Dim Flen2 As Long
    Dim ReCode2 As String
    Wsock(Index).GetData ByteData2, vbByte
    ReCode2 = DownJet(Index).SaveData(bytesTotal, ByteData2(), Flen2)
    Select Case ReCode2
    Case "200"
        txtInfo.Text = txtInfo.Text & DownJet(Index).DownUrl & "开始下载" & vbCrLf
    Case "206"
        txtInfo.Text = txtInfo.Text & DownJet(Index).DownUrl & "开始从" & DownJet(Index).mFlen & "断点续传下载" & vbCrLf
    Case "404"
        txtInfo.Text = txtInfo.Text & DownJet(Index).DownUrl & "文件不存在" & vbCrLf
        CloseSocket Index, "文件未找到！"
    Case "error"
        CloseSocket Index, "请求时出错"
        txtInfo.Text = txtInfo.Text & DownJet(Index).DownUrl & "出错了" & vbCrLf
    Case "cancel"
        CloseSocket Index, "用户取消"
        Exit Sub
    End Select
    If Flen2 > 0 Then
        mDownInfo(DownJet(Index).WhichDown).mSize = Flen2
        LView.ListItems(DownJet(Index).WhichDown).SubItems(1) = Flen2
    End If
End If
End Sub

'控制描绘下载情况block的PictureBox的可见
Public Sub PicVisible(Index As Integer)
Dim i As Integer
For i = 0 To Pic.Count - 1
    Pic(i).Visible = False
Next i
Pic(Index).Visible = True
End Sub


'根据接收到的文件长度，已经下载长度的信息在Pic画Block图
'mflen：文件长度
'mNum：接收到的字节数
'ReceiveBytes：已经接收到的字节数
Public Sub DrawDownPic(Index As Integer, mNum As Long, Optional mFlen As Long, Optional ReceiveBytes As Long)
If mNum > 0 Then
    mDownInfo(DownJet(Index).WhichDown).mGetSize = mDownInfo(DownJet(Index).WhichDown).mGetSize + mNum
    LView.ListItems(DownJet(Index).WhichDown).SubItems(2) = ReceiveBytes + mNum
End If
Dim Getnum As Long
Getnum = ReceiveBytes
Dim TGetNum As Long
Dim i, j As Long
Dim kk1 As Long, kk2 As Long
If mNum = 0 Then
    Getnum = 0
End If

If Getnum = 0 Then
    Pic(Index).FillColor = vbWhite
    kk1 = mFlen / 4096
    j = 0
    For i = 1 To mFlen / 4096
        Pic(Index).Circle ((i - j * 50) * 120 + 0, j * 120 + 100), 50, vbBlack
        j = Fix(i / 50)
    Next i
    Pic(Index).FillColor = &HFF0000
End If

TGetNum = Getnum
If Getnum = 0 And ReceiveBytes > 0 Then
    '加上以前已经接收到的
    Getnum = ReceiveBytes
End If
Getnum = Getnum + mNum
kk1 = Fix(TGetNum / 4096)
kk2 = Fix(Getnum / 4096)
j = Fix(kk1 / 50) + 1
For i = kk1 To kk2
    Pic(Index).Circle ((i - j * 50) * 120 + 0, j * 120 + 100), 50, vbRed
    j = Fix(i / 50)
Next i
End Sub


Private Sub Wsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
CloseSocket Index, "Winsock出错"
End Sub

'关闭socket后做的一些处理
Public Sub CloseSocket(Index As Integer, ClsStr As String)
If mDownInfo(DownJet(Index).WhichDown).mGetSize + 1 >= mDownInfo(DownJet(Index).WhichDown).mSize And mDownInfo(DownJet(Index).WhichDown).mSize > 0 Then
    LView.ListItems(DownJet(Index).WhichDown).SmallIcon = "ok"
    txtInfo.Text = txtInfo.Text & mDownInfo(DownJet(Index).WhichDown).mUrl & "的下载完成了" & vbCrLf
Else
    LView.ListItems(DownJet(Index).WhichDown).SmallIcon = "stop"
    txtInfo.Text = txtInfo.Text & mDownInfo(DownJet(Index).WhichDown).mUrl & "的下载因为" & ClsStr & "被关闭了" & vbCrLf
End If
DownJet(Index).bBusy = False
Wsock(Index).Close
CurrentDown(Index) = 0
End Sub

'加入新的下载任务
Public Function AddNewUrl(myUrl As String)
    Dim i As Integer
    For i = 1 To LView.ListItems.Count
        If LView.ListItems(i).Text = myUrl Then
            MsgBox "该URL已经在下载队列中了！"
            Exit Function
        End If
    Next i
    AddUrl myUrl
    Dim kk As Integer
    kk = UBound(mDownInfo)
    ReDim Preserve mDownInfo(kk + 1)
    mDownInfo(kk).mUrl = myUrl
End Function

