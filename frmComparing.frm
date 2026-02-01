VERSION 5.00
Begin VB.Form frmComparing 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "비교 중..."
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
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
   ScaleHeight     =   1380
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소"
      Default         =   -1  'True
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin prjReadTest.ProgressBar pbProgress 
      Height          =   255
      Left            =   960
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Step            =   10
   End
   Begin VB.Label Label1 
      Caption         =   "두 측정 기록을 비교하고 있습니다..."
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Picture         =   "frmComparing.frx":000C
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmComparing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmMain.FlagStopCompare = True
    cmdCancel.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = 1
        frmMain.FlagStopCompare = True
        cmdCancel.Enabled = False
    End If
End Sub
