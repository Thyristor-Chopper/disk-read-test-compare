VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "정보"
   ClientHeight    =   5265
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtLicense 
      Height          =   3255
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.TextBox txtLicense 
      Height          =   3255
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   345
      Left            =   5880
      TabIndex        =   4
      Top             =   4800
      Width           =   1335
   End
   Begin prjReadTest.ImageList imgItems 
      Left            =   360
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      InitListImages  =   "frmAbout.frx":000C
   End
   Begin VB.Frame FrameW1 
      BorderStyle     =   0  '없음
      Caption         =   "라이선스(&L)"
      Height          =   3255
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   6375
      Begin prjReadTest.ListView lvItems 
         Height          =   3255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   5741
         Icons           =   "imgItems"
         Arrange         =   2
         LabelEdit       =   2
         HideSelection   =   0   'False
         ShowInfoTips    =   -1  'True
         ShowLabelTips   =   -1  'True
         ShowColumnTips  =   -1  'True
         SnapToGrid      =   -1  'True
      End
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  '투명
      Caption         =   "버전"
      Height          =   225
      Left            =   1050
      TabIndex        =   1
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '투명
      Caption         =   "응용 프로그램 제목"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   0
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  '투명
      Caption         =   "This product includes software developed by vbAccelerator."
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1050
      TabIndex        =   2
      Top             =   960
      Width           =   6405
   End
   Begin VB.Image picIcon 
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ScrollBars(1 To 2) As Byte

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    lvItems.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    ScrollBars(1) = 0
    ScrollBars(2) = 0
    
    Me.Caption = App.Title & " 정보"
    Set picIcon.Picture = frmMain.Icon
    lblVersion.Caption = "버전 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    
    lvItems.ListItems.Add , , "Krool's Comctl", 1
    lvItems.ListItems.Add , , "vbAccelerator SSubTmr", 2
    lvItems.ListItems(1).Selected = True
End Sub

Private Sub lvItems_ItemSelect(Item As LvwListItem, ByVal Selected As Boolean)
    On Error Resume Next
    If Not Selected Then Exit Sub 'If Item Is lvItems.SelectedItem Then Item.Selected = True: Exit Sub
    txtLicense(-(Not -ScrollBars(Item.Index))).Visible = False
    txtLicense(ScrollBars(Item.Index)).Visible = True
    txtLicense(ScrollBars(Item.Index)).Text = LoadResText(200 + Item.Index, RCData)
End Sub
