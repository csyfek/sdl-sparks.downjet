VERSION 5.00
Begin VB.Form frmAdd 
   Caption         =   "加入新的下载任务"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   6465
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "加入下载的URL"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1392
      Left            =   60
      TabIndex        =   11
      Top             =   90
      Width           =   6312
      Begin VB.CommandButton cmdAdd 
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   1
         Left            =   4920
         TabIndex        =   14
         Top             =   900
         Width           =   792
      End
      Begin VB.TextBox txtUrl 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   960
         TabIndex        =   13
         Text            =   "http://zju.yi.org/~coolman/mp3/liu1.mp3"
         Top             =   360
         Width           =   4932
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "加入"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   0
         Left            =   3900
         TabIndex        =   12
         Top             =   900
         Width           =   792
      End
      Begin VB.Label Label5 
         Caption         =   "URL："
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   300
         TabIndex        =   15
         Top             =   420
         Width           =   552
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "选项"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3012
      Left            =   60
      TabIndex        =   0
      Top             =   1500
      Width           =   6312
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1800
         TabIndex        =   10
         Top             =   2160
         Width           =   1332
      End
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1800
         TabIndex        =   8
         Top             =   2580
         Width           =   1332
      End
      Begin VB.CheckBox chkAuth 
         Caption         =   "代理服务器身份验证"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   180
         TabIndex        =   6
         Top             =   1800
         Width           =   2172
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1800
         TabIndex        =   4
         Text            =   "6666"
         Top             =   1320
         Width           =   792
      End
      Begin VB.TextBox txtProxy 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1800
         TabIndex        =   2
         Text            =   "zjupry2.zju.edu.cn"
         Top             =   900
         Width           =   3912
      End
      Begin VB.CheckBox chkProxy 
         Caption         =   "使用代理服务器"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   180
         TabIndex        =   1
         Top             =   420
         Width           =   1632
      End
      Begin VB.Label Label4 
         Caption         =   "密码："
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   1632
      End
      Begin VB.Label Label3 
         Caption         =   "帐号："
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   7
         Top             =   2220
         Width           =   1632
      End
      Begin VB.Label Label2 
         Caption         =   "代理服务器端口："
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   5
         Top             =   1380
         Width           =   1632
      End
      Begin VB.Label Label1 
         Caption         =   "代理服务器地址："
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1632
      End
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAuth_Click()
If chkAuth.Value = 1 Then
    txtID.Enabled = True
    txtPass.Enabled = True
Else
    txtID.Enabled = False
    txtPass.Enabled = False
End If
End Sub

Private Sub chkProxy_Click()
If chkProxy.Value = 1 Then
        txtProxy.Enabled = True
        txtPort.Enabled = True
        chkAuth.Enabled = True
Else
        txtProxy.Enabled = False
        txtPort.Enabled = False
        chkAuth.Enabled = False
        txtID.Enabled = False
        txtPass.Enabled = False
End If
chkAuth.Value = 0
End Sub

Private Sub cmdAdd_Click(Index As Integer)
If Index = 0 Then
    Dim i As Integer
    For i = 1 To frmDown.LView.ListItems.Count
        If frmDown.LView.ListItems(i).Text = txtUrl.Text Then
            MsgBox "该URL已经在下载队列中了！"
            txtUrl.Text = ""
            Exit Sub
        End If
    Next i
    frmDown.AddUrl txtUrl
    Dim kk As Integer
    kk = UBound(mDownInfo)
    ReDim Preserve mDownInfo(kk + 1)
    mDownInfo(kk).mUrl = txtUrl.Text
    If chkProxy.Value = 1 Then
        mDownInfo(kk).mUseProxy = True
        mDownInfo(kk).mProxy = txtProxy.Text
        mDownInfo(kk).mProxyPort = Val(txtPort.Text)
        If chkAuth.Value = 1 Then
            mDownInfo(kk).mProxyId = txtID.Text
            mDownInfo(kk).mProxyPass = txtPass.Text
        End If
    End If
    
End If
    SelectDown = 0
    Unload Me
    
End Sub

Private Sub Form_Activate()
Dim mStr1 As String
mStr1 = Clipboard.GetText
If InStr(1, mStr1, "http://") = 1 And Len(mStr1) < 180 Then
    txtUrl.Text = mStr1
End If
End Sub

Private Sub Form_Load()

If SelectDown > 0 And SelectDown <= UBound(mDownInfo) Then

    txtUrl.Text = mDownInfo(SelectDown).mUrl
    If mDownInfo(SelectDown).mUseProxy = True Then
        chkProxy.Value = 1
        txtProxy.Text = mDownInfo(SelectDown).mProxy
        txtPort.Text = mDownInfo(SelectDown).mProxyPort
        If mDownInfo(SelectDown).mProxyId <> "" Then
            chkAuth.Value = 1
            txtID.Text = mDownInfo(SelectDown).mProxyId
            txtPass.Text = mDownInfo(SelectDown).mProxyPass
        End If
    End If
End If
End Sub
