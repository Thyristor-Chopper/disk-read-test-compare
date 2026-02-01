VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "디스크 섹터 검사 및 시간 경과 후 읽기 시간 비교"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9975
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows 기본값
   Begin prjReadTest.StatusBar sbStatusBar 
      Align           =   2  '아래 맞춤
      Height          =   330
      Left            =   0
      Top             =   9645
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   582
      InitPanels      =   "frmMain.frx":08CA
   End
   Begin VB.PictureBox pbPanel 
      Height          =   4815
      Index           =   2
      Left            =   120
      ScaleHeight     =   4755
      ScaleWidth      =   7995
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton cmdCompare 
         Caption         =   "비교(&C)"
         Height          =   330
         Left            =   5040
         TabIndex        =   23
         Top             =   3300
         Width           =   1455
      End
      Begin VB.CheckBox chkUseCurrent 
         Caption         =   "현재 측정 결과 사용(&U)"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   720
         Width           =   2895
      End
      Begin prjReadTest.SpinBox txtThreshold 
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   3330
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Min             =   1
         Max             =   32000
         Value           =   50
         Increment       =   5
         ThousandsSeparator=   0   'False
         AllowOnlyNumbers=   -1  'True
      End
      Begin VB.CommandButton cmdBrowseSecond 
         Caption         =   "..."
         Height          =   300
         Left            =   6120
         TabIndex        =   14
         Top             =   345
         Width           =   375
      End
      Begin VB.CommandButton cmdBrowseFirst 
         Caption         =   "..."
         Height          =   300
         Left            =   2760
         TabIndex        =   11
         Top             =   345
         Width           =   375
      End
      Begin VB.TextBox txtSecond 
         Height          =   270
         Left            =   3600
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtFirst 
         Height          =   270
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin prjReadTest.ListView lvCompare 
         Height          =   2175
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3836
         VisualTheme     =   1
         View            =   3
         FullRowSelect   =   -1  'True
         LabelEdit       =   2
         Sorted          =   -1  'True
         SortKey         =   3
         SortOrder       =   1
         SortType        =   2
      End
      Begin VB.Label lblUnit 
         Caption         =   "MB/초 이상"
         Height          =   255
         Left            =   2700
         TabIndex        =   24
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label lblThreshold 
         Caption         =   "오차 강조 범위(&T):"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lblSecond 
         Caption         =   "이번 측정 자료(&S):"
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblFirst 
         Caption         =   "이전 측정 자료(&F):"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.PictureBox pbPanel 
      Height          =   3735
      Index           =   1
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   600
      Width           =   6615
      Begin VB.CommandButton cmdStop 
         Caption         =   "측정 중지(&S)"
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "파일로 저장(&A)..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   1935
      End
      Begin VB.DriveListBox lvDrive 
         Height          =   300
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin prjReadTest.ProgressBar pbProgress 
         Height          =   270
         Left            =   2520
         Top             =   3300
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   476
         Step            =   10
      End
      Begin prjReadTest.ListView lvTestResult 
         Height          =   3015
         Left            =   2520
         TabIndex        =   7
         Top             =   120
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   5318
         VisualTheme     =   1
         View            =   3
         FullRowSelect   =   -1  'True
         LabelEdit       =   2
         Sorted          =   -1  'True
         SortKey         =   1
         SortType        =   2
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "측정 시작(&S)"
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblAllocUnit 
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblTotalSectors 
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "할당 단위 크기:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "총 섹터 수:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblDrive 
         Caption         =   "드라이브(&V):"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
   End
   Begin prjReadTest.TabStrip tsTabStrip 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      TabMinWidth     =   48
      InitTabs        =   "frmMain.frx":0A3E
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Running As Boolean
Dim FlagStop As Boolean
Public FlagStopCompare As Boolean
Dim TabBackgroundHint As Long

Implements IBSSubclass

Private Sub chkUseCurrent_Click()
    If chkUseCurrent = 1 Then
        txtSecond.Enabled = False
        txtSecond.BackColor = TabBackgroundHint
        cmdBrowseSecond.Enabled = False
    Else
        txtSecond.Enabled = True
        txtSecond.BackColor = &H80000005
        cmdBrowseSecond.Enabled = True
    End If
End Sub

Private Sub cmdBrowseFirst_Click()
    Dim Path As String
    Path = PromptOpen(Title:="첫째 측정 기록 열기")
    If LenB(Path) Then txtFirst = Path
End Sub

Private Sub cmdBrowseSecond_Click()
    Dim Path As String
    Path = PromptOpen(Title:="둘째 측정 기록 열기")
    If LenB(Path) Then txtSecond = Path
End Sub

Private Sub cmdCompare_Click()
    Dim iFileNo As Integer
    Dim FirstPath$, SecondPath$
    Dim RawFirst$, RawSecond$
    Dim First() As String, Second() As String
    Dim IndexMap As New Collection
    
    If chkUseCurrent <> 0 Then
        RawSecond = Stringify
    Else
        SecondPath = txtSecond
        If LenB(SecondPath) = 0 Then
            MsgBox "둘째 기록 파일을 선택하십시오.", 64
            Exit Sub
        End If
        If Not FileExists(SecondPath) Then
            MsgBox "둘째 기록 파일이 존재하지 않습니다.", 48
            Exit Sub
        End If
        
        iFileNo = FreeFile()
        Open SecondPath For Input As #iFileNo
            Line Input #iFileNo, RawSecond
        Close #iFileNo
    End If
    
    FirstPath = txtFirst
    If LenB(FirstPath) = 0 Then
        MsgBox "첫째 기록 파일을 선택하십시오.", 64
        Exit Sub
    End If
    If Not FileExists(FirstPath) Then
        MsgBox "첫째 기록 파일이 존재하지 않습니다.", 48
        Exit Sub
    End If
    
    iFileNo = FreeFile()
    Open FirstPath For Input As #iFileNo
        Line Input #iFileNo, RawFirst
    Close #iFileNo
    
    First = Split(RawFirst, ",")
    Second = Split(RawSecond, ",")
    
    cmdCompare.Enabled = False
    txtThreshold.Enabled = False
    cmdBrowseFirst.Enabled = False
    txtFirst.Enabled = False
    lblFirst.Enabled = False
    lblSecond.Enabled = False
    txtSecond.Enabled = False
    cmdBrowseSecond.Enabled = False
    chkUseCurrent.Enabled = False
    lblThreshold.Enabled = False
    
    FlagStopCompare = False
    
    Load frmComparing
    frmComparing.Top = pbPanel(2).Height / 2 - frmComparing.Height / 2
    frmComparing.Left = pbPanel(2).Width / 2 - frmComparing.Width / 2
    SetParent frmComparing.hWnd, pbPanel(2).hWnd
    ShowWindow frmComparing.hWnd, SW_SHOW
    SetWindowPos frmComparing.hWnd, hWnd_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
    
    frmComparing.pbProgress.Value = LBound(Second)
    frmComparing.pbProgress.Max = UBound(Second)
    frmComparing.pbProgress.Min = LBound(Second)
    
    Dim i&, k&
    Dim LVITEM As LvwListItem
    k = 0
    lvCompare.Sorted = False
    lvCompare.ListItems.Clear
    lvCompare.Redraw = False
    Dim Diff As Double
    For i = LBound(Second) To UBound(Second) Step 3
        If FlagStopCompare = True Then
            frmComparing.pbProgress.Value = i
            Exit For
        End If
        Set LVITEM = lvCompare.ListItems.Add(Text:=Second(i) & " - " & Second(i + 1))
        LVITEM.ListSubItems.Add Text:="-"
        LVITEM.ListSubItems.Add Text:=Second(i + 2)
        LVITEM.ListSubItems.Add Text:="-"
        IndexMap.Add LVITEM.Index, Second(i) & "-" & Second(i + 1)
        If UBound(First) >= i Then
            If First(i) = Second(i) And First(i + 1) = Second(i + 1) Then
                LVITEM.ListSubItems(1).Text = First(i + 2)
                If LVITEM.ListSubItems(1).Text = "오류" Or LVITEM.ListSubItems(2).Text = "오류" Then
                    LVITEM.ListSubItems(3).Text = "배드"
                ElseIf IsNumeric(LVITEM.ListSubItems(1).Text) And IsNumeric(LVITEM.ListSubItems(2).Text) Then
                    Diff = CDbl(LVITEM.ListSubItems(1).Text) - CDbl(LVITEM.ListSubItems(2).Text)
                    LVITEM.ListSubItems(3).Text = Format$(Diff, "0.00")
                    If Diff > txtThreshold.Value Then
                        LVITEM.Bold = True
                        LVITEM.ForeColor = vbRed
                        LVITEM.ListSubItems(1).Bold = True
                        LVITEM.ListSubItems(1).ForeColor = vbRed
                        LVITEM.ListSubItems(2).Bold = True
                        LVITEM.ListSubItems(2).ForeColor = vbRed
                        LVITEM.ListSubItems(3).Bold = True
                        LVITEM.ListSubItems(3).ForeColor = vbRed
                    End If
                End If
                LVITEM.Tag = "F"
            End If
        End If
        k = k + 1
        If k >= 50 Then
            k = 0
            lvCompare.Redraw = True
            DoEvents
            lvCompare.Redraw = False
            frmComparing.pbProgress.Value = i
        End If
    Next i
'    k = 0
'    For i = LBound(First) To UBound(First) Step 3
'        If Exists(IndexMap, First(i) & "-" & First(i + 1)) Then
'            Set lvItem = lvCompare.ListItems(IndexMap(First(i) & "-" & First(i + 1)))
'        Else
'            Set lvItem = lvCompare.ListItems.Add(Text:=Second(i) & " - " & Second(i + 1))
'            lvItem.ListSubItems.Add
'            lvItem.ListSubItems.Add Text:="-"
'            lvItem.ListSubItems.Add Text:="-"
'        End If
'        If lvItem.Tag <> "F" Then
'            lvItem.ListSubItems(1).Text = First(i + 2)
'            If lvItem.ListSubItems(1).Text = "오류" Or lvItem.ListSubItems(2).Text = "오류" Then
'                lvItem.ListSubItems(3).Text = "배드"
'            ElseIf IsNumeric(lvItem.ListSubItems(1).Text) And IsNumeric(lvItem.ListSubItems(2).Text) Then
'                lvItem.ListSubItems(3).Text = Format$(CDbl(lvItem.ListSubItems(2).Text) - CDbl(lvItem.ListSubItems(1).Text), "0.00")
'            End If
'        End If
'        If k >= 50 Then
'            k = 0
'            DoEvents
'        End If
'    Next i
    Unload frmComparing
    lvCompare.Redraw = False
    If lvCompare.ColumnHeaders(1).SortArrow <> LvwColumnHeaderSortArrowUp Then lvCompare.Sorted = True
    lvCompare.Redraw = True
    
    cmdCompare.Enabled = True
    txtThreshold.Enabled = True
    cmdBrowseFirst.Enabled = True
    txtFirst.Enabled = True
    lblFirst.Enabled = True
    lblSecond.Enabled = True
    txtSecond.Enabled = (chkUseCurrent = 0)
    cmdBrowseSecond.Enabled = (chkUseCurrent = 0)
    chkUseCurrent.Enabled = True
    lblThreshold.Enabled = True
End Sub

Private Function Stringify() As String
    Stringify = ""
    Dim i As Long
    For i = 1 To lvTestResult.ListItems.Count
        Stringify = Stringify & Join(Split(lvTestResult.ListItems(i).Text, " - "), ",") & "," & lvTestResult.ListItems(i).ListSubItems(1).Text & ","
    Next i
    On Error Resume Next
    Stringify = Left$(Stringify, Len(Stringify) - 1)
End Function

Private Sub cmdSave_Click()
    Dim Path As String
    Path = PromptSave("섹터측정_" & UCase(Left$(lvDrive.Drive, 1)) & "_" & Format(Now, "YYYYMMDD_HHMMSS"), "측정 기록 저장")
    If LenB(Path) = 0 Then Exit Sub
    
    On Error Resume Next
    Kill Path
    On Error GoTo 0
    
    Dim iFileNo As Integer
    iFileNo = FreeFile()
    Open Path For Output As #iFileNo
        Print #iFileNo, Stringify()
    Close #iFileNo
End Sub

Private Sub cmdStart_Click()
    Dim hVol As Long
    Dim StartTime@, EndTime@, Frequency@
    Dim ReadRet As Long
    Dim TotalSectors As Currency
    
    Dim BytesRead As Long
    Dim SectorIndex As Currency
    Dim BlockSize As Currency
    
    Dim LengthInfo As GET_LENGTH_INFORMATION
    Dim BytesReturned As Long
    
    Dim sectorsPerCluster As Long
    Dim bytesPerSector As Long
    Dim numFreeClusters As Long
    Dim totalClusters As Long
    
    Dim SectorsToRead As Currency
    
    Dim DriveLetter As String
    DriveLetter = UCase(Left$(lvDrive.Drive, 2))
    
    '할당 단위 크기
    If GetDiskFreeSpace(DriveLetter & "\", sectorsPerCluster, bytesPerSector, numFreeClusters, totalClusters) = 0 Then
        MsgBox "디스크 할당 단위 크기를 알아내는 데 실패했습니다.", 16
        GoTo EndRead
    End If
    
    Dim allocationUnit As Currency
    allocationUnit = CCur(sectorsPerCluster) * bytesPerSector

    '볼륨 열기
    hVol = CreateFile("\\.\" & DriveLetter, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, FILE_FLAG_NO_BUFFERING Or FILE_FLAG_SEQUENTIAL_SCAN, 0&)
    If hVol = -1 Then
        MsgBox "볼륨을 열 수 없습니다.", 16
        GoTo EndRead
    End If
    
    '총 크기
    If DeviceIoControl(hVol, IOCTL_DISK_GET_LENGTH_INFO, 0&, 0&, LengthInfo, Len(LengthInfo), BytesReturned, 0&) = 0 Then
        MsgBox "볼륨의 크기를 알아내는 데 실패했습니다.", 16
        GoTo EndRead
    End If
    
    TotalSectors = (LengthInfo.Length * 10000) / bytesPerSector
    
    Dim pBuffer As Long
    pBuffer = VirtualAlloc(0, CLng(allocationUnit), MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
    If pBuffer = 0 Then
        MsgBox "메모리가 부족합니다.", 16
        GoTo EndRead
    End If
    
    SectorIndex = 0
    BlockSize = allocationUnit / bytesPerSector
    
    QueryPerformanceFrequency Frequency
    
    pbProgress.Value = 0
    cmdStart.Enabled = False
    cmdStop.Enabled = True
    cmdStop.Visible = True
    cmdStart.Visible = False
    lvTestResult.Sorted = False
    lvTestResult.ListItems.Clear
    FlagStop = False
    Running = True
    cmdSave.Enabled = False
    lblDrive.Enabled = False
    lvDrive.Enabled = False
    
    Dim CurItem As LvwListItem
    Dim i As Byte
    i = 0
    lvTestResult.Redraw = False
    Do While SectorIndex < TotalSectors
        If FlagStop Then
            pbProgress.Value = CInt((SectorIndex / TotalSectors) * 100)
            sbStatusBar.Panels(1).Text = "측정 중... " & CInt((SectorIndex / TotalSectors) * 100) & "%"
            sbStatusBar.Panels(2).Text = TotalSectors & " 중 " & SectorIndex
            Exit Do
        End If
        
        If SetFilePointerEx(hVol, CCur(SectorIndex * bytesPerSector) / 10000@, 0&, FILE_BEGIN) = 0& Then
            MsgBox "포인터 이동 중 문제가 발생했습니다.", 16
            Exit Do
        End If
        
        SectorsToRead = BlockSize
        If SectorIndex + SectorsToRead > TotalSectors Then
            SectorsToRead = TotalSectors - SectorIndex
        End If
        If SectorsToRead < BlockSize Then Exit Do
        If SectorsToRead <= 0 Then Exit Do
        
        QueryPerformanceCounter StartTime
        ReadRet = ReadFile(hVol, ByVal pBuffer, CLng(SectorsToRead * bytesPerSector), BytesRead, 0&)
        QueryPerformanceCounter EndTime
        
        Set CurItem = lvTestResult.ListItems.Add(Text:=SectorIndex & " - " & (SectorIndex + SectorsToRead - 1))
        If ReadRet = 0 Or BytesRead = 0 Then
            CurItem.ForeColor = vbRed
            CurItem.ListSubItems.Add(Text:="오류").ForeColor = vbRed
        Else
            Dim Elapsed#, Speed#
            Elapsed = (EndTime - StartTime) / Frequency
            If Elapsed > 0 Then
                Speed = (BytesRead / 1024# / 1024#) / Elapsed
            Else
                Speed = 0
            End If
            CurItem.ListSubItems.Add Text:=Format$(Speed, "0.00")
        End If
        CurItem.EnsureVisible
        
        SectorIndex = SectorIndex + SectorsToRead
        
        i = i + 1
        If i = 50 Then
            i = 0
            lvTestResult.Redraw = True
            DoEvents
            lvTestResult.Redraw = False
            pbProgress.Value = CInt((SectorIndex / TotalSectors) * 100)
            sbStatusBar.Panels(1).Text = "측정 중... " & CInt((SectorIndex / TotalSectors) * 100) & "%"
            sbStatusBar.Panels(2).Text = TotalSectors & " 중 " & SectorIndex
        End If
    Loop
    
EndRead:
    VirtualFree pBuffer, 0&, MEM_RELEASE
    CloseHandle hVol
    If lvTestResult.ColumnHeaders(1).SortArrow <> LvwColumnHeaderSortArrowUp Then lvTestResult.Sorted = True
    If FlagStop Then
        sbStatusBar.Panels(1).Text = "중단됨"
    Else
        sbStatusBar.Panels(1).Text = "완료"
    End If
    lvTestResult.Redraw = True
    FlagStop = False
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    cmdStart.Visible = True
    cmdStop.Visible = False
    Running = False
    cmdSave.Enabled = True
    lblDrive.Enabled = True
    lvDrive.Enabled = True
End Sub

Private Sub cmdStop_Click()
    If MsgBox("진행 중인 검사를 중지하시겠습니까? 지금까지의 기록은 여전히 저장할 수 있습니다", vbQuestion + vbYesNo) = vbYes Then FlagStop = True
End Sub

Private Sub SetBackColor()
    On Error Resume Next
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.Container.Name = "pbPanel" And (Not TypeOf ctrl Is TextBox) And (Not TypeOf ctrl Is ListView) And (Not TypeOf ctrl Is ComboBox) And (Not TypeOf ctrl Is DriveListBox) And (Not TypeOf ctrl Is SpinBox) Then
            ctrl.BackColor = TabBackgroundHint
        End If
    Next ctrl
End Sub

Private Sub Form_Load()
    DPI = GetDPI()
    UpdateBorderWidth
    
    TabBackgroundHint = GetThemeColor(Me.hWnd, "TAB", 9&, 1&, 3821&, &H8000000F)
    Dim i As Byte
    For i = pbPanel.LBound To pbPanel.UBound
         pbPanel(i).BackColor = TabBackgroundHint
         pbPanel(i).BorderStyle = 0
    Next i
    SetBackColor
    InitPropertySheetDimensions Me, tsTabStrip, pbPanel
    
    lvTestResult.ColumnHeaders.Add Text:="섹터", Width:=15 * 150
    lvTestResult.ColumnHeaders.Add Text:="속도 (MB/초)", Width:=15 * 115, Alignment:=LvwColumnHeaderAlignmentRight
    
    lvCompare.ColumnHeaders.Add Text:="섹터", Width:=15 * 150
    lvCompare.ColumnHeaders.Add Text:="이전 속도", Width:=15 * 100, Alignment:=LvwColumnHeaderAlignmentRight
    lvCompare.ColumnHeaders.Add Text:="이번 속도", Width:=15 * 100, Alignment:=LvwColumnHeaderAlignmentRight
    lvCompare.ColumnHeaders.Add Text:="오차 (MB/초)", Width:=15 * 120, Alignment:=LvwColumnHeaderAlignmentRight
    
    Running = False
    lvDrive_Change
    
    lvTestResult.ColumnHeaders(1).SortArrow = LvwColumnHeaderSortArrowUp
    lvCompare.ColumnHeaders(4).SortArrow = LvwColumnHeaderSortArrowDown
    
    Dim hSysMenu As Long
    Dim MenuCount As Long
    hSysMenu = GetSystemMenu(Me.hWnd, 0&)
    MenuCount = GetMenuItemCount(hSysMenu)
    Dim MII As MENUITEMINFO
    MII.cbSize = MiiSize
    With MII
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
        .fType = MFT_STRING
        .fState = MFS_ENABLED
        .wID = 1000
        .dwTypeData = "프로그램 정보(&A)"
        .cch = Len(.dwTypeData)
    End With
    InsertMenuItem hSysMenu, SC_CLOSE, MF_BYCOMMAND, MII
    With MII
        '.fMask = MIIM_ID Or MIIM_TYPE
        .fType = MFT_SEPARATOR
        .wID = 2000
    End With
    InsertMenuItem hSysMenu, SC_CLOSE, MF_BYCOMMAND, MII
    
    AttachMessage Me, Me.hWnd, WM_GETMINMAXINFO
    AttachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    AttachMessage Me, Me.hWnd, WM_THEMECHANGED
    AttachMessage Me, Me.hWnd, WM_SYSCOMMAND
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    tsTabStrip.Width = Me.Width - 240 - SizingBorderWidth * Screen.TwipsPerPixelX * 2
    tsTabStrip.Height = Me.Height - sbStatusBar.Height - 120 - 120 - SizingBorderWidth * Screen.TwipsPerPixelY * 2 - CaptionHeight * Screen.TwipsPerPixelY
    Dim i As Byte
    For i = pbPanel.LBound To pbPanel.UBound
        pbPanel(i).Width = tsTabStrip.ClientWidth
        pbPanel(i).Height = tsTabStrip.ClientHeight
    Next i
    lvCompare.Height = pbPanel(2).Height - (3735 - 2175)
    lblThreshold.Top = lvCompare.Top + lvCompare.Height + 120
    lblUnit.Top = lblThreshold.Top
    txtThreshold.Top = lblThreshold.Top - 30
    cmdCompare.Top = txtThreshold.Top - 30
    lvTestResult.Height = pbPanel(1).Height - (3735 - 3135)
    pbProgress.Top = lvTestResult.Top + lvTestResult.Height + 120
    cmdSave.Top = pbProgress.Top + pbProgress.Height - cmdSave.Height
    cmdStart.Top = cmdSave.Top - cmdStart.Height - 120
    cmdStop.Top = cmdStart.Top
    lvTestResult.Width = pbPanel(1).Width - (6615 - 3975)
    pbProgress.Width = lvTestResult.Width
    lvCompare.Width = pbPanel(2).Width - (6615 - 6375)
    cmdCompare.Left = lvCompare.Left + lvCompare.Width - cmdCompare.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Running Then
        If MsgBox("진행 중인 검사를 중지하시겠습니까?", vbQuestion + vbYesNo) = vbYes Then
            FlagStop = True
        Else
            Exit Sub
        End If
    End If
    Unload frmMessageBox
    Unload frmData
    Unload frmComparing
    
    IBSSubclass_UnsubclassIt
End Sub

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    DetachMessage Me, Me.hWnd, WM_GETMINMAXINFO
    DetachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    DetachMessage Me, Me.hWnd, WM_THEMECHANGED
    DetachMessage Me, Me.hWnd, WM_SYSCOMMAND
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    On Error Resume Next
    
    Select Case uMsg
        Case WM_GETMINMAXINFO
            Dim lpMMI As MINMAXINFO
            CopyMemory lpMMI, ByVal lParam, Len(lpMMI)
            lpMMI.ptMinTrackSize.X = (465 + SizingBorderWidth * 2) * (DPI / 96)
            lpMMI.ptMinTrackSize.Y = (252 + SizingBorderWidth * 2 + CaptionHeight) * (DPI / 96)
            lpMMI.ptMaxTrackSize.X = (Screen.Width + 1200) * (DPI / 96)
            lpMMI.ptMaxTrackSize.Y = (Screen.Height + 1200) * (DPI / 96)
            CopyMemory ByVal lParam, lpMMI, Len(lpMMI)
            
            IBSSubclass_WindowProc = 1&
            Exit Function
        Case WM_SETTINGCHANGE
            Select Case GetStrFromPtr(lParam)
                Case "WindowMetrics"
                    UpdateBorderWidth
                    Form_Resize
            End Select
        Case WM_THEMECHANGED
            TabBackgroundHint = GetThemeColor(Me.hWnd, "TAB", 9&, 1&, 3821&, &H8000000F)
            SetBackColor
        Case WM_SYSCOMMAND
            If wParam = 1000& Then '항상 위에 표시
                frmAbout.Show vbModal
                
                IBSSubclass_WindowProc = 1&
                Exit Function
            End If
    End Select
    
    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Sub lvCompare_ColumnClick(ColumnHeader As LvwColumnHeader)
    If ColumnHeader.SortArrow <> LvwColumnHeaderSortArrowNone Then
        If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown Then
            ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp
        Else
            ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown
        End If
    Else
        Dim i As Byte
        For i = 1 To lvCompare.ColumnHeaders.Count
            If lvCompare.ColumnHeaders(i) Is ColumnHeader Then
                lvCompare.ColumnHeaders(i).SortArrow = LvwColumnHeaderSortArrowUp
            Else
                lvCompare.ColumnHeaders(i).SortArrow = LvwColumnHeaderSortArrowNone
            End If
        Next i
    End If
    lvCompare.Sorted = False
    If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp Then
        lvCompare.SortOrder = LvwSortOrderAscending
    Else
        lvCompare.SortOrder = LvwSortOrderDescending
    End If
    lvCompare.SortKey = ColumnHeader.Index - 1
    lvCompare.Sorted = True
End Sub

Private Sub lvDrive_Change()
    Dim sectorsPerCluster As Long
    Dim bytesPerSector As Long
    Dim numFreeClusters As Long
    Dim totalClusters As Long
    
    Dim DriveLetter As String
    DriveLetter = UCase(Left$(lvDrive.Drive, 2))
    
    If GetDiskFreeSpace(DriveLetter & "\", sectorsPerCluster, bytesPerSector, numFreeClusters, totalClusters) = 0 Then
        lblAllocUnit = "알 수 없음"
    Else
        Dim allocationUnit As Currency
        allocationUnit = CCur(sectorsPerCluster) * bytesPerSector
        lblAllocUnit = ParseSize(allocationUnit)
    End If
    
    Dim hVol As Long
    
    hVol = CreateFile("\\.\" & DriveLetter, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, FILE_FLAG_NO_BUFFERING Or FILE_FLAG_SEQUENTIAL_SCAN, 0&)
    If hVol = -1 Then
        lblTotalSectors = "알 수 없음"
    Else
        Dim BytesReturned As Long
        Dim LengthInfo As GET_LENGTH_INFORMATION
        If DeviceIoControl(hVol, IOCTL_DISK_GET_LENGTH_INFO, 0&, 0&, LengthInfo, Len(LengthInfo), BytesReturned, 0&) = 0 Then
            lblTotalSectors = "알 수 없음"
        Else
            lblTotalSectors = (LengthInfo.Length * 10000) / bytesPerSector
        End If
    End If
    
    CloseHandle hVol
End Sub

Private Sub lvTestResult_ColumnClick(ColumnHeader As LvwColumnHeader)
    If ColumnHeader.SortArrow <> LvwColumnHeaderSortArrowNone Then
        If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown Then
            ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp
        Else
            ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown
        End If
    Else
        Dim i As Byte
        For i = 1 To lvTestResult.ColumnHeaders.Count
            If lvTestResult.ColumnHeaders(i) Is ColumnHeader Then
                lvTestResult.ColumnHeaders(i).SortArrow = LvwColumnHeaderSortArrowUp
            Else
                lvTestResult.ColumnHeaders(i).SortArrow = LvwColumnHeaderSortArrowNone
            End If
        Next i
    End If
    lvTestResult.Sorted = False
    If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp Then
        lvTestResult.SortOrder = LvwSortOrderAscending
    Else
        lvTestResult.SortOrder = LvwSortOrderDescending
    End If
    lvTestResult.SortKey = ColumnHeader.Index - 1
    lvTestResult.Sorted = True
End Sub

Private Sub tsTabStrip_TabClick(TabItem As TbsTab)
    On Error Resume Next
    Static i As Byte, Show As Boolean
    For i = pbPanel.LBound To pbPanel.UBound
        Show = (i = TabItem.Index)
        pbPanel(i).Visible = Show
        pbPanel(i).Enabled = Show
        If Show Then pbPanel(i).ZOrder 0
    Next i
End Sub
