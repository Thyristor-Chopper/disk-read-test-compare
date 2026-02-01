VERSION 5.00
Begin VB.Form frmProcessing 
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComparing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소"
      Default         =   -1  'True
      Height          =   330
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin prjReadTest.ProgressBar pbProgress 
      Height          =   255
      Left            =   960
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      Step            =   10
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Picture         =   "frmComparing.frx":000C
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public AllowCancel As Boolean
Public Cancelled As Boolean

Private Sub DisableClose()
    Dim SystemMenu As Long
    SystemMenu = GetSystemMenu(Me.hWnd, 0&)
    DeleteMenu SystemMenu, SC_CLOSE, MF_BYCOMMAND
    DeleteMenu SystemMenu, 0&, MF_BYCOMMAND
    SetWindowPos Me.hWnd, 0&, 0&, 0&, 0&, 0&, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
End Sub

Private Sub cmdCancel_Click()
    If AllowCancel = False Then Exit Sub
    Cancelled = True
    cmdCancel.Enabled = False
    DisableClose
End Sub

Sub Init()
    If AllowCancel = False Then DisableClose
    cmdCancel.Enabled = AllowCancel
End Sub

Private Sub Form_Load()
    Cancelled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = 1
        If AllowCancel = False Then Exit Sub
        Cancelled = True
        cmdCancel.Enabled = False
        DisableClose
    End If
End Sub
