VERSION 5.00
Begin VB.Form MinForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DownJet"
   ClientHeight    =   468
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   1248
   FillColor       =   &H000000C0&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   468
   ScaleWidth      =   1248
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MinForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_DblClick()
If frmDown.WindowState = vbNormal Then
    frmDown.WindowState = vbMinimized
Else
    frmDown.WindowState = vbNormal
End If
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Data.GetFormat(vbCFText) Then
    If vbOK = MsgBox("确定要下载" & Data.GetData(vbCFText), vbOKCancel) Then
        frmDown.AddNewUrl Data.GetData(vbCFText)
        frmDown.WindowState = vbNormal
    End If
End If
End Sub

