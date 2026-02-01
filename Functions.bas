Attribute VB_Name = "Functions"
Option Explicit

Private Type LVITEM
    Mask As Long
    iItem As Long
    iSubItem As Long
    State As Long
    StateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
End Type

Private Const LVM_GETITEMTEXT As Long = &H102D&
Private Const LVM_GETITEMCOUNT As Long = &H1004&
Private Const LVIF_TEXT As Long = &H1&
Private Const MAX_TEXT As Long = 256&

Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMII As MENUITEMINFO) As Long
Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMII As MENUITEMINFO) As Long

Public Const SC_MOVE = &HF010&
Public Const SC_RESTORE = &HF120&
Public Const SC_SIZE = &HF000&
Public Const SC_CLOSE = &HF060&
Global Const MF_BYPOSITION = &H400
Global Const MF_BYCOMMAND = &H0&

Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_ENABLED = &H0
Public Const MFS_GRAYED = &H3
Public Const MFS_DISABLED = MFS_GRAYED
Public Const MFS_CHECKED = &H8

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hBmpChecked As Long
    hBmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Const hWnd_TOPMOST As Long = -1&
Public Const hWnd_NOTOPMOST As Long = -2&

Public Const SW_SHOW As Long = 5&

Public Const SWP_NOMOVE As Long = &H2&
Public Const SWP_NOSIZE As Long = &H1&
Public Const SWP_NOZORDER As Long = &H4&
Public Const SWP_FRAMECHANGED As Long = &H20&

Public PaddedBorderWidth As Byte
Public DialogBorderWidth As Byte
Public SizingBorderWidth As Byte
Public ScrollBarWidth As Byte
Public CaptionHeight As Byte
Public DPI As Long

Public Const WM_ERASEBKGND  As Long = &H14&
Public Const WM_NOTIFY As Long = &H4E&
Public Const WM_MOVE As Long = &H3&
'Public Const WM_MOVING As Long = &H216&
Public Const WM_SETCURSOR As Long = &H20&
Public Const WM_NCPAINT As Long = &H85&
Public Const WM_COMMAND As Long = &H111&
Public Const WM_SIZE As Long = &H5&
'Public Const WM_SIZING As Long = &H214&
Public Const WM_GETMINMAXINFO As Long = &H24
Public Const WM_SYSCOMMAND As Long = &H112&
Public Const WM_INITMENU As Long = &H116&
Public Const WM_SETTINGCHANGE As Long = &H1A
Public Const WM_DWMCOMPOSITIONCHANGED As Long = &H31E&
'Public Const WM_DWMCOLORIZATIONCOLORCHANGED As Long = &H320&
Public Const WM_THEMECHANGED As Long = &H31A&
Public Const WM_DPICHANGED As Long = &H2E0&
Public Const WM_CTLCOLORSCROLLBAR As Long = &H137&
Public Const WM_CTLCOLORSTATIC As Long = &H138&
Public Const WM_CTLCOLORBTN As Long = &H135&
Public Const WM_PAINT As Long = &HF&
Public Const WM_PRINTCLIENT As Long = &H318&
Public Const WM_NCCALCSIZE As Long = &H83&
Public Const WM_NCHITTEST As Long = &H84&
Public Const WM_NCACTIVATE As Long = &H86&
Public Const WA_INACTIVE As Long = 0&
Public Const WA_ACTIVE As Long = 1&

Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, ByVal lpInBuffer As Long, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Declare Function SetFilePointerEx Lib "kernel32" (ByVal hFile As Long, ByVal liDistanceToMove As Currency, ByRef lpNewFilePointer As Currency, ByVal dwMoveMethod As Long) As Long

Public Const MEM_COMMIT As Long = &H1000&
Public Const MEM_RESERVE As Long = &H2000&
Public Const MEM_RELEASE As Long = &H8000&
Public Const PAGE_READWRITE As Long = &H4&
Public Const FILE_BEGIN As Long = 0&

Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const FILE_SHARE_READ As Long = &H1&
Public Const FILE_SHARE_WRITE As Long = &H2&
Public Const FILE_FLAG_NO_BUFFERING As Long = &H20000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80&
Public Const CREATE_ALWAYS As Long = 2&
Public Const OPEN_EXISTING As Long = 3&

Public Const FSCTL_LOCK_VOLUME As Long = &H90018
Public Const FSCTL_DISMOUNT_VOLUME As Long = &H90020

Public Const IOCTL_DISK_GET_LENGTH_INFO = &H7405C

Enum OpenSaveMode
    SaveFile = 0
    OpenFile = 1
    OpenDirectory = 2
End Enum

Type GET_LENGTH_INFORMATION
    Length As Currency
End Type

Private Declare Function X_GetThemeColor Lib "uxtheme.dll" Alias "GetThemeColor" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, pColor As Long) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Private Declare Function IsThemeActive Lib "uxtheme.dll" () As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

Public Const TMT_TEXTCOLOR As Long = 3803

Public MsgBoxResults As New Collection
Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const SfiSize As Long = 352&
Public Const MiiSize As Long = 44&

Enum VbMsgBoxResultEx
'    vbAbort = 3
'    vbCancel = 2
'    vbIgnore = 5
'    vbNo = 7
'    vbOK = 1
'    vbRetry = 4
'    vbYes = 6
    vbTryAgain = 10
    vbContinue = 11
End Enum

Enum VbMsgBoxStyleEx
'    vbAbortRetryIgnore = 2
'    vbApplicationModal = 0
'    vbCritical = 16
'    vbDefaultButton1 = 0
'    vbDefaultButton2 = 256
'    vbDefaultButton3 = 512
'    vbDefaultButton4 = 768
'    vbExclamation = 48
'    vbInformation = 64
'    vbMsgBoxHelpButton = 16384
'    vbMsgBoxRight = 524288
'    vbMsgBoxRtlReading = 1048576
'    vbMsgBoxSetForeground = 65536
'    vbOKCancel = 1
'    vbOKOnly = 0
'    vbQuestion = 32
'    vbRetryCancel = 5
'    vbSystemModal = 4096
'    vbYesNo = 4
'    vbYesNoCancel = 3
    vbCancelTryContinue = 6
    vbYesNoEx = 7
End Enum

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As Any, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal lpRootPathName As String) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As String
Private Declare Function ShellExecuteEx Lib "shell32" (ByRef s As SHELLEXECUTEINFO) As Long

Public Const LVM_SETVIEW As Long = 4238&
Public Const LVM_SETIMAGELIST As Long = &H1003&
Public Const LVSIL_SMALL As Long = 1&
Public Const LVSIL_NORMAL As Long = 0&

'Public Const SHGFI_DISPLAYNAME As Long = &H200&
Public Const SHGFI_SYSICONINDEX As Long = &H4000&
Public Const SHGFI_ICON As Long = &H100&
Public Const SHGFI_LARGEICON As Long = &H0&
Public Const SHGFI_SMALLICON As Long = &H1&
Public Const SHGFI_USEFILEATTRIBUTES As Long = &H10&
Public Const SHGFI_TYPENAME As Long = &H400&

Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_INTERNET = &H1
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_COMMON_STARTUP = &H18
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_PRINTHOOD = &H1B
Public Const CSIDL_ALTSTARTUP = &H1D
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E
Public Const CSIDL_COMMON_FAVORITES = &H1F
Public Const CSIDL_INTERNET_CACHE = &H20
Public Const CSIDL_COOKIES = &H21
Public Const CSIDL_HISTORY = &H22

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' optional fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Const INVALID_HANDLE_VALUE As Long = -1&

Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime As Currency
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Private Type ItemID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As ItemID
End Type

Enum DriveTypes
    DRIVE_UNKNOWN = 0
    DRIVE_NO_ROOT_DIR = 1
    DRIVE_REMOVABLE = 2
    DRIVE_FIXED = 3
    DRIVE_REMOTE = 4
    DRIVE_CDROM = 5    'can be a CD or a DVD
    DRIVE_RAMDISK = 6
End Enum

Type POINTAPI
   X As Long
   Y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Const CB_ERR As Long = -1&
Private Const CB_ADDSTRING As Long = &H143&
Private Const CB_RESETCONTENT As Long = &H14B&
Private Const CB_SETITEMDATA As Long = &H151&

Private Const SEE_MASK_INVOKEIDLIST = &HC

Type lParam
    lParam As Long
End Type

Enum ResourceType
    BITMAP = 2
    Icon = 3
    RCData = 10
    Manifest = 24
End Enum

Function LoadResText(ByVal ResourceID As Integer, ByVal ResourceType As ResourceType) As String
    LoadResText = StrConv(LoadResData(ResourceID, ResourceType), vbUnicode)
End Function

Function StartsWith(Str As String, s As String) As Boolean
    StartsWith = (Left$(Str, Len(s)) = s)
End Function

Function EndsWith(Str As String, s As String) As Boolean
    EndsWith = (Right$(Str, Len(s)) = s)
End Function

Sub ShellExecute(sFile As String, Optional Action As String = "open", Optional WorkingDirectory As String)
    Dim shinfo As SHELLEXECUTEINFO
    With shinfo
        .cbSize = LenB(shinfo)
        .lpFile = sFile
        .nShow = SW_SHOW
        If Action = "properties" Then .fMask = SEE_MASK_INVOKEIDLIST
        If LenB(WorkingDirectory) Then .lpDirectory = WorkingDirectory
        .lpVerb = Action
    End With
    ShellExecuteEx shinfo
End Sub

Function GetStrFromPtr(ByVal Ptr As Long) As String
    GetStrFromPtr = SysAllocStringByteLen(Ptr, lstrlen(Ptr))
End Function

Function GetDPI() As Long
    Dim hWndDesktop As Long
    Dim hDCDesktop As Long

    hWndDesktop = GetDesktopWindow()
    hDCDesktop = GetDC(hWndDesktop)
    GetDPI = GetDeviceCaps(hDCDesktop, 88&)
    ReleaseDC hWndDesktop, hDCDesktop
End Function

Sub AddItemToComboBox(cbComboBox As ComboBox, ByVal Text As String)
    SendMessage cbComboBox.hWnd, CB_ADDSTRING, 0&, ByVal Text
End Sub

Sub ClearComboBox(cbComboBox As ComboBox)
    SendMessage cbComboBox.hWnd, CB_RESETCONTENT, 0&, 0&
End Sub

Sub GetDiskSpace(sDrive As String, ByRef dblTotal As Double, ByRef dblFree As Double)
    Dim lResult As Long
    Dim liAvailable As LARGE_INTEGER
    Dim liTotal As LARGE_INTEGER
    Dim liFree As LARGE_INTEGER
    lResult = GetDiskFreeSpaceEx(sDrive, liAvailable, liTotal, liFree)
    dblTotal = CLargeInt(liTotal.LowPart, liTotal.HighPart)
    dblFree = CLargeInt(liFree.LowPart, liFree.HighPart)
End Sub

Private Function CLargeInt(Lo As Long, Hi As Long) As Double
    Dim dblLo As Double, dblHi As Double

    If Lo < 0 Then
        dblLo = 2 ^ 32 + Lo
    Else
        dblLo = Lo
    End If

    If Hi < 0 Then
        dblHi = 2 ^ 32 + Hi
    Else
        dblHi = Hi
    End If
    
    CLargeInt = dblLo + dblHi * 2 ^ 32
End Function

Function ParseSize(ByVal Size As Double, Optional ByVal ShowBytes As Boolean = False, Optional Suffix As String = "") As String
    If Size < 0 Then
        ParseSize = "-"
        Exit Function
    End If

    On Error GoTo ErrLn4
    Dim ret#
    If Size >= 1024# * 1024# * 1024# * 1024# Then
        ret = Fix(Size / 1024# / 1024# / 1024# / 1024# * 100) / 100
        If ret >= 100# Then
            ret = Fix(ret)
        ElseIf ret >= 10# Then
            ret = Fix(ret * 10) / 10
        End If
        ParseSize = ret & "TB" & Suffix
    ElseIf Size >= 1024# * 1024# * 1024# Then
        ret = Fix(Size / 1024# / 1024# / 1024# * 100) / 100
        If ret >= 100# Then
            ret = Fix(ret)
        ElseIf ret >= 10# Then
            ret = Fix(ret * 10) / 10
        End If
        ParseSize = ret & "GB" & Suffix
    ElseIf Size >= 1024# * 1024# Then
        ret = Fix(Size / 1024# / 1024# * 100) / 100
        If ret >= 100# Then
            ret = Fix(ret)
        ElseIf ret >= 10# Then
            ret = Fix(ret * 10) / 10
        End If
        ParseSize = ret & "MB" & Suffix
    ElseIf Size >= 1024# Then
        ret = Fix(Size / 1024# * 100) / 100
        If ret >= 100# Then
            ret = Fix(ret)
        ElseIf ret >= 10# Then
            ret = Fix(ret * 10) / 10
        End If
        ParseSize = ret & "KB" & Suffix
    Else
        ParseSize = CStr(Size) & " " & "바이트" & Suffix
    End If

    If Size >= (1024#) And ShowBytes Then
        ParseSize = ParseSize & " (" & Size & " " & "바이트" & Suffix & ")"
    End If
    Exit Function
ErrLn4:
    ParseSize = "0 " & "바이트" & Suffix
End Function

Function FolderExists(sFullPath As String) As Boolean
    On Error GoTo nonexist
    FolderExists = ((GetAttr(sFullPath) And (vbDirectory Or vbVolume)) <> 0)
    Exit Function
nonexist:
    FolderExists = False
End Function

Function GetShortcutTarget(sPath As String) As String
    Dim shl As Shell, file As FolderItem, fld As shell32.Folder
    Dim lnk As ShellLinkObject, i As Long, folderPath As String
    Dim Shortcutname As String

    On Error GoTo exit_sub
    folderPath = GetParentFolderName(sPath)
    Set shl = New Shell
    Set fld = shl.NameSpace(folderPath)
    Set file = fld.Items.Item(GetFilename(sPath))
    If Err Then
        GetShortcutTarget = "."
        GoTo exit_sub
    Else
        If file.IsLink Then
            Set lnk = file.GetLink
            GetShortcutTarget = lnk.Path
            If Left$(GetShortcutTarget, 1) = """" And Right$(GetShortcutTarget, 1) = """" Then GetShortcutTarget = Mid$(GetShortcutTarget, 2, Len(GetShortcutTarget) - 2)
        Else
            GetShortcutTarget = "."
        End If
    End If
    
exit_sub:
    Set lnk = Nothing
    Set file = Nothing
    Set fld = Nothing
    Set shl = Nothing
End Function

Function FormatModified(DateTime) As String
    FormatModified = Replace(Replace(Format(DateTime, "yyyy-mm-dd AM/PM h:mm"), "AM", "오전"), "PM", "오후")
End Function

Function GetParentFolderName(ByVal Path As String) As String
    On Error GoTo errfso
    Do While Right$(Path, 1) = "\"
        Path = Left$(Path, Len(Path) - 1)
    Loop
    If InStrRev(Path, "\") = 0 Then GoTo errfso
    GetParentFolderName = Left$(Path, InStrRev(Path, "\") - 1)
    Do While Right$(GetParentFolderName, 1) = "\"
        GetParentFolderName = Left$(GetParentFolderName, Len(GetParentFolderName) - 1)
    Loop
    If Len(GetParentFolderName) = 2 And Right$(GetParentFolderName, 1) = ":" Then GetParentFolderName = GetParentFolderName & "\"
    Exit Function
errfso:
    GetParentFolderName = ""
End Function

Function GetFilename(ByVal Path As String) As String
    On Error GoTo errfso
    Do While Right$(Path, 1) = "\"
        Path = Left$(Path, Len(Path) - 1)
    Loop
    GetFilename = Mid$(Path, InStrRev(Path, "\") + 1)
    Exit Function
errfso:
    GetFilename = ""
End Function

Function GetExtensionName(ByVal Path As String) As String
    On Error GoTo errfso
    Path = GetFilename(Path)
    If InStrRev(Path, ".") = 0 Then GoTo errfso
    GetExtensionName = Mid$(Path, InStrRev(Path, ".") + 1)
    Exit Function
errfso:
    GetExtensionName = ""
End Function

Function GetSpecialFolder(CSIDL As Long) As String
    Dim lngRetVal As Long
    Dim IDL As ITEMIDLIST
    Dim strPath As String
    lngRetVal = SHGetSpecialFolderLocation(100&, CSIDL, IDL)
    If lngRetVal = 0& Then
        strPath = Space$(512)
        lngRetVal = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath)
        GetSpecialFolder = Left$(strPath, InStr(strPath, Chr$(0)) - 1)
    End If
End Function

Function GetThemeColor(ByVal hWnd As Long, ClassList As String, Optional ByVal Part As Long = 0&, Optional ByVal State As Long = 0&, Optional ByVal Prop As Long = TMT_TEXTCOLOR, Optional ByVal DefaultColor As Long = 0&) As Long
    On Error GoTo returndefault
    Dim hTheme&, clr&

    If IsAppThemed() = 0& Or IsThemeActive() = 0& Then GoTo returndefault
    hTheme = OpenThemeData(hWnd, StrPtr(ClassList))
    If hTheme = 0& Then GoTo returndefault
    If X_GetThemeColor(hTheme, Part, State, Prop, clr) <> 0 Then GoTo returndefault
    CloseThemeData hTheme
    GetThemeColor = clr
    Exit Function

returndefault:
    If hTheme <> 0& Then CloseThemeData hTheme
    GetThemeColor = DefaultColor
End Function

Sub InitPropertySheetDimensions(frmForm As Form, tsTabStrip As TabStrip, Panels As Object)
    Dim i As Byte
    Dim MaxWidth%, MaxHeight%
    Dim ClientLeft%, ClientTop%, Left%, Top%, Width%, Height%
    ClientLeft = tsTabStrip.ClientLeft
    ClientTop = tsTabStrip.ClientTop
    Left = tsTabStrip.Left
    Top = tsTabStrip.Top
    Width = tsTabStrip.Width
    Height = tsTabStrip.Height
    For i = Panels.LBound To Panels.UBound
        Panels(i).Top = ClientTop + Top
        Panels(i).Left = ClientLeft + Left
        If MaxWidth < Panels(i).Width Then MaxWidth = Panels(i).Width
        If MaxHeight < Panels(i).Height Then MaxHeight = Panels(i).Height
    Next i
    For i = Panels.LBound To Panels.UBound
        Panels(i).Width = MaxWidth
        Panels(i).Height = MaxHeight
    Next i
    Width = MaxWidth + (Width - tsTabStrip.ClientWidth)
    tsTabStrip.Width = Width
    Height = MaxHeight + (Height - tsTabStrip.ClientHeight)
    tsTabStrip.Height = Height
    frmForm.Height = Top + Height + 540 + 330 + 120
    frmForm.Width = Width + 300
End Sub

Function Exists(oCol As Collection, vKey As String) As Boolean
    On Error Resume Next
    oCol.Item CStr(vKey)
    Exists = (Err.Number = 0)
    Err.Clear
End Function

Function StrLen(s As String) As Integer
    StrLen = LenB(StrConv(s, vbFromUnicode))
End Function

Private Function CutLines(Text As String, ByVal Width As Single) As String()
    Dim Paragraphs() As String
    Dim ParagraphX As Long
    Dim Words() As String
    Dim WordX As Long
    Dim CutLine As String
    Dim NewCutLine As String
    Dim SingleWord As Boolean
    Dim ForceX As Long
    Dim Lines() As String
    Dim LineX As Long

    Paragraphs = Split(Text, vbNewLine)
    For ParagraphX = 0 To UBound(Paragraphs)
        Words = Split(Paragraphs(ParagraphX), " ")
        WordX = 0
        Do While WordX <= UBound(Words)
            Do
                If Len(CutLine) = 0 Then
                    NewCutLine = Words(WordX)
                    SingleWord = True
                Else
                    NewCutLine = NewCutLine & " " & Words(WordX)
                End If
                If frmData.TextWidth(NewCutLine) > Width Then Exit Do
                CutLine = NewCutLine
                WordX = WordX + 1
                SingleWord = False
            Loop While WordX <= UBound(Words)
            If SingleWord Then
                For ForceX = Len(Words(WordX)) - 1 To 1 Step -1
                    CutLine = Left$(Words(WordX), ForceX)
                    If frmData.TextWidth(CutLine) <= Width Then
                        Words(WordX) = Mid$(Words(WordX), ForceX + 1)
                        Exit For
                    End If
                Next
            End If
            ReDim Preserve Lines(LineX)
            Lines(LineX) = CutLine
            LineX = LineX + 1
            CutLine = vbNullString
        Loop
    Next
    CutLines = Lines
End Function

Function RandInt(StartNumber, EndNumber)
    RandInt = Int(Rnd * (EndNumber - StartNumber + 1)) + StartNumber
End Function

Function ShowMessageBox(ByVal Content As String, Optional ByVal Title As String, Optional Icon As VbMsgBoxStyle = 64, Optional IsModal As Boolean = True, Optional AlertTimeout As Integer = -1, Optional ByVal DefaultOption As VbMsgBoxResult = vbNo, Optional ByVal MsgBoxMode As VbMsgBoxStyle = vbOKOnly, Optional YesCaption As String = "예(&Y)", Optional NoCaption As String = "아니요(&N)") As VbMsgBoxResult
    If Title = "" Then Title = App.Title

    Dim MessageBox As frmMessageBox
    Set MessageBox = New frmMessageBox
    MessageBox.MsgBoxMode = MsgBoxMode
    MessageBox.ResultID = CStr(Rnd * 1E+15)
    Set MessageBox.MessageBoxObject = MessageBox

    On Error Resume Next

    MessageBox.imgIcon(Icon / 16).Visible = True

    Content = Replace(Content, "&", "&&")
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf & " " & vbCrLf)

    Dim i%
    Dim LineCount As Integer
    Dim LContent As Integer
    Dim MAX_WIDTH As Long
    MAX_WIDTH = Screen.Width / 2
    Content = Join(CutLines(Content, MAX_WIDTH), vbCrLf)
    LContent = 0
    LineCount = UBound(Split(Content, vbLf)) + 1
    Dim s%
    Dim ln$
    Dim CI%, c$
    Dim LineContent$
    For s = 0 To UBound(Split(Content, vbCrLf))
        LineContent = Split(Content, vbCrLf)(s)
        If frmData.TextWidth(LineContent) > LContent Then LContent = frmData.TextWidth(LineContent)
    Next s

    If LContent = 0 Then LContent = StrLen(Content)
    If LineCount > 1 Then MessageBox.lblContent.Top = 280

    Dim MsgBoxMinWidth As Integer
    Select Case MsgBoxMode
        Case vbYesNo, vbRetryCancel, vbOKCancel, vbYesNoEx
            MsgBoxMinWidth = 3480
        Case vbYesNoCancel, vbAbortRetryIgnore, vbCancelTryContinue
            MsgBoxMinWidth = 4920
        Case Else
            MsgBoxMinWidth = 1920
    End Select

    MessageBox.Height = 1615 + LineCount * 180 - 300 + 190 - 60 + IIf(MsgBoxMode = vbYesNoEx, 735, 0)
    MessageBox.Caption = Title
    MessageBox.lblContent.Caption = Content
    If 1175 + LContent > MsgBoxMinWidth Then MessageBox.Width = 1175 + LContent + 60 Else MessageBox.Width = MsgBoxMinWidth

    Select Case MsgBoxMode
        Case vbYesNo
            MessageBox.cmdYes.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdYes.Width / 2
            MessageBox.cmdYes.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdNo.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdYes.Width / 2 - 120 + MessageBox.cmdYes.Width + 240 - 30
            MessageBox.cmdNo.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdYes.Top = MessageBox.cmdYes.Top + 180
                MessageBox.cmdNo.Top = MessageBox.cmdNo.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdYes.Top = MessageBox.cmdYes.Top - 210
'                MessageBox.cmdNo.Top = MessageBox.cmdNo.Top - 210
'            End If
        Case vbYesNoEx
            MessageBox.cmdOK.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdOK.Width / 2
            MessageBox.cmdOK.Top = 840 + (LineCount * 185) - 350 + 705
            MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdOK.Width / 2 - 120 + MessageBox.cmdOK.Width + 240 - 30
            MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350 + 705
            MessageBox.optYes.Top = MessageBox.cmdOK.Top - 620
            MessageBox.optNo.Top = MessageBox.cmdOK.Top - 320
            If LineCount > 1 Then
                MessageBox.optYes.Top = MessageBox.optYes.Top - 80
                MessageBox.optNo.Top = MessageBox.optNo.Top - 80
            End If
            If IsEmpty(DefaultOption) Then
                MessageBox.optYes.Value = False
                MessageBox.optNo.Value = False
                MessageBox.cmdOK.Enabled = False
            ElseIf DefaultOption = vbYes Then
                MessageBox.optYes.Value = True
                MessageBox.cmdOK.Enabled = True
            Else
                MessageBox.optNo.Value = True
                MessageBox.cmdOK.Enabled = True
            End If
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top + 180
                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top - 210
'                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
'                MessageBox.optYes.Top = MessageBox.optYes.Top - 210
'                MessageBox.optNo.Top = MessageBox.optNo.Top - 210
'            End If
        Case vbYesNoCancel
            MessageBox.cmdYes.Left = MessageBox.Width / 2 - 900 - MessageBox.cmdYes.Width
            MessageBox.cmdYes.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdNo.Left = MessageBox.Width / 2 - 810 + 15
            MessageBox.cmdNo.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 900 + MessageBox.cmdYes.Width + 190 + 30
            MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdYes.Top = MessageBox.cmdYes.Top + 180
                MessageBox.cmdNo.Top = MessageBox.cmdNo.Top + 180
                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
'                MessageBox.cmdYes.Top = MessageBox.cmdYes.Top - 210
'                MessageBox.cmdNo.Top = MessageBox.cmdNo.Top - 210
'            End If
        Case vbRetryCancel
            MessageBox.cmdRetry.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdRetry.Width / 2
            MessageBox.cmdRetry.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdCancel.Width / 2 - 120 + MessageBox.cmdRetry.Width + 240 - 30
            MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top + 180
                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top - 210
'                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
'            End If
        Case vbAbortRetryIgnore
            MessageBox.cmdAbort.Left = MessageBox.Width / 2 - 900 - MessageBox.cmdAbort.Width
            MessageBox.cmdAbort.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdRetry.Left = MessageBox.Width / 2 - 810 + 15
            MessageBox.cmdRetry.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdIgnore.Left = MessageBox.Width / 2 - 900 + MessageBox.cmdAbort.Width + 190 + 30
            MessageBox.cmdIgnore.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdAbort.Top = MessageBox.cmdAbort.Top + 180
                MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top + 180
                MessageBox.cmdIgnore.Top = MessageBox.cmdIgnore.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdIgnore.Top = MessageBox.cmdIgnore.Top - 210
'                MessageBox.cmdAbort.Top = MessageBox.cmdAbort.Top - 210
'                MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top - 210
'            End If
        Case vbOKCancel
            MessageBox.cmdOK.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdOK.Width / 2
            MessageBox.cmdOK.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdCancel.Width / 2 - 120 + MessageBox.cmdOK.Width + 240 - 30
            MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top + 180
                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top - 210
'                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
'            End If
        Case vbCancelTryContinue
            MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 900 - MessageBox.cmdCancel.Width
            MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdTryAgain.Left = MessageBox.Width / 2 - 810 + 15
            MessageBox.cmdTryAgain.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdContinue.Left = MessageBox.Width / 2 - 900 + MessageBox.cmdCancel.Width + 190 + 30
            MessageBox.cmdContinue.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
                MessageBox.cmdTryAgain.Top = MessageBox.cmdTryAgain.Top + 180
                MessageBox.cmdContinue.Top = MessageBox.cmdContinue.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdContinue.Top = MessageBox.cmdContinue.Top - 210
'                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
'                MessageBox.cmdTryAgain.Top = MessageBox.cmdTryAgain.Top - 210
'            End If
        Case Else 'vbOKOnly
            MessageBox.cmdOK.Left = MessageBox.Width / 2 - 810 + 30
            MessageBox.cmdOK.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top - 210
'            End If
    End Select
    
    MessageBox.lblContent.Height = MessageBox.Height

    MessageBeep Icon

    If MsgBoxMode = vbOKOnly And AlertTimeout >= 0 Then
        MessageBox.timeout.Interval = AlertTimeout
        MessageBox.timeout.Enabled = -1
    End If

    MessageBox.cmdOK.Visible = (MsgBoxMode = vbOKOnly Or MsgBoxMode = vbYesNoEx Or MsgBoxMode = vbOKCancel)
    MessageBox.cmdCancel.Visible = (MsgBoxMode = vbYesNoEx Or MsgBoxMode = vbYesNoCancel Or MsgBoxMode = vbRetryCancel Or MsgBoxMode = vbOKCancel Or MsgBoxMode = vbCancelTryContinue)
    MessageBox.cmdYes.Visible = (MsgBoxMode = vbYesNo Or MsgBoxMode = vbYesNoCancel)
    MessageBox.cmdNo.Visible = (MsgBoxMode = vbYesNo Or MsgBoxMode = vbYesNoCancel)
    MessageBox.optYes.Visible = (MsgBoxMode = vbYesNoEx)
    MessageBox.optNo.Visible = (MsgBoxMode = vbYesNoEx)

    MessageBox.cmdAbort.Visible = (MsgBoxMode = vbAbortRetryIgnore)
    MessageBox.cmdRetry.Visible = (MsgBoxMode = vbAbortRetryIgnore Or MsgBoxMode = vbRetryCancel)
    MessageBox.cmdIgnore.Visible = (MsgBoxMode = vbAbortRetryIgnore)
    MessageBox.cmdContinue.Visible = (MsgBoxMode = vbCancelTryContinue)
    MessageBox.cmdTryAgain.Visible = (MsgBoxMode = vbCancelTryContinue)
    MessageBox.cmdHelp.Visible = False

    MessageBox.cmdCancel.Cancel = (MsgBoxMode = vbYesNoEx Or MsgBoxMode = vbYesNoCancel Or MsgBoxMode = vbRetryCancel Or MsgBoxMode = vbOKCancel Or MsgBoxMode = vbCancelTryContinue)
    MessageBox.cmdCancel.Default = False
    MessageBox.cmdYes.Cancel = False
    MessageBox.cmdYes.Default = False
    MessageBox.cmdNo.Cancel = False
    MessageBox.cmdNo.Default = False
    MessageBox.cmdOK.Cancel = (MsgBoxMode = vbOKOnly)
    MessageBox.cmdOK.Default = (MsgBoxMode = vbOKOnly Or MsgBoxMode = vbYesNoEx)

    MessageBox.Init
    If MsgBoxMode = vbOKOnly Then
        If IsModal Then
            MessageBox.Show vbModal
            Unload MessageBox
            Set MessageBox = Nothing
        Else
            MessageBox.Show
        End If
        ShowMessageBox = vbOK
    Else
        MessageBox.Show vbModal
        ShowMessageBox = MsgBoxResults(MessageBox.ResultID)
        MsgBoxResults.Remove MessageBox.ResultID
        Unload MessageBox
        Set MessageBox = Nothing
    End If
End Function

Function ConfirmEx(ByVal Content As String, Optional ByVal Title As String, Optional ByVal Icon As VbMsgBoxStyle = 32, Optional ByVal DefaultOption As VbMsgBoxResult = vbNo, Optional YesCaption As String = "예(&Y)", Optional NoCaption As String = "아니요(&N)") As VbMsgBoxResult
    ConfirmEx = ShowMessageBox(Content, Title, Icon, DefaultOption:=DefaultOption, MsgBoxMode:=vbYesNoEx, YesCaption:=YesCaption, NoCaption:=NoCaption)
End Function

Function MsgBox(ByVal Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String) As VbMsgBoxResult
    If Title = "" Then Title = App.Title
    If Buttons > 70 Then
        GoTo nativemsgbox
    ElseIf Buttons < 16 Then
        MsgBox = ShowMessageBox(Prompt, Title, 0, MsgBoxMode:=Buttons)
    ElseIf (Buttons And vbInformation) = vbInformation Then
        MsgBox = ShowMessageBox(Prompt, Title, vbInformation, MsgBoxMode:=(Buttons And (Not vbInformation)))
    ElseIf (Buttons And vbExclamation) = vbExclamation Then
        MsgBox = ShowMessageBox(Prompt, Title, vbExclamation, MsgBoxMode:=(Buttons And (Not vbExclamation)))
    ElseIf (Buttons And vbQuestion) = vbQuestion Then
        MsgBox = ShowMessageBox(Prompt, Title, vbQuestion, MsgBoxMode:=(Buttons And (Not vbQuestion)))
    ElseIf (Buttons And vbCritical) = vbCritical Then
        MsgBox = ShowMessageBox(Prompt, Title, vbCritical, MsgBoxMode:=(Buttons And (Not vbCritical)))
    Else
        GoTo nativemsgbox
    End If

    Exit Function
nativemsgbox:
    MsgBox = VBA.MsgBox(Prompt, Buttons, Title)
End Function

Function PromptSave(Optional PresetPath As String, Optional Title As String) As String
    Dim Explorer As frmExplorer
    Set Explorer = New frmExplorer
    Explorer.BrowseMode = SaveFile
    Explorer.PresetPath = PresetPath
    If LenB(Title) Then Explorer.Caption = Title
    Explorer.Show vbModal
    PromptSave = Explorer.ReturnPath
    Unload Explorer
    Set Explorer = Nothing
End Function

Function PromptOpen(Optional PresetPath As String, Optional Title As String) As String
    Dim Explorer As frmExplorer
    Set Explorer = New frmExplorer
    Explorer.BrowseMode = OpenFile
    Explorer.PresetPath = PresetPath
    If LenB(Title) Then Explorer.Caption = Title
    Explorer.Show vbModal
    PromptOpen = Explorer.ReturnPath
    Unload Explorer
    Set Explorer = Nothing
End Function

Sub UpdateBorderWidth()
    DialogBorderWidth = GetSystemMetrics(8&)
    SizingBorderWidth = GetSystemMetrics(33&)
    PaddedBorderWidth = SizingBorderWidth - DialogBorderWidth
    CaptionHeight = GetSystemMetrics(31&)
    ScrollBarWidth = GetSystemMetrics(2&)
End Sub

Function Right(Str As String, Length As Long) As String
    On Error GoTo errproc
    Right = VBA.Right$(Str, Length)
    Exit Function
errproc:
    Right = ""
End Function

Function ListItemText(hWnd As Long, ByVal iItem As Long, Optional ByVal iSubItem As Long = 0&) As String
    Dim sText As String: sText = String$(MAX_TEXT, vbNullChar)
    
    Dim Item As LVITEM
    Item.Mask = LVIF_TEXT
    Item.iItem = iItem
    Item.iSubItem = iSubItem
    Item.pszText = sText
    Item.cchTextMax = MAX_TEXT
    
    SendMessage hWnd, LVM_GETITEMTEXT, iItem - 1&, Item
    ListItemText = Left$(Item.pszText, InStr(Item.pszText, vbNullChar) - 1)
End Function

Function ReadLine(ByVal Path As String) As Collection
    Dim hFile As Long
    Dim BytesRead As Long
    Dim BufferSize As Long
    Dim LineStart As Long
    Set ReadLine = New Collection
    Dim i As Long
    BufferSize = 65536
    Dim Buffer() As Byte
    ReDim Buffer(BufferSize - 1)
    Dim Carry() As Byte
    Dim CarryLen As Long
    ReDim Carry(0)
    CarryLen = 0&
    hFile = CreateFile(Path, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    If hFile = -1 Then Exit Function
    Dim CurLine As String
    Do
        If ReadFile(hFile, Buffer(0), BufferSize, BytesRead, 0&) = 0& Then Exit Do
        If BytesRead = 0& Then Exit Do
        LineStart = 0&
        For i = 0& To BytesRead - 1&
            If Buffer(i) = 10 Then
                Dim LineBytes() As Byte
                Dim LineLen As Long
                LineLen = CarryLen + (i - LineStart)
                If LineLen > 0 Then
                    ReDim LineBytes(LineLen - 1)
                    If CarryLen Then CopyMemory LineBytes(0), Carry(0), CarryLen
                    CopyMemory LineBytes(CarryLen), Buffer(LineStart), i - LineStart - 1&
                    CurLine = StrConv(LineBytes, vbUnicode)
                    ReadLine.Add CurLine
                End If
                CarryLen = 0&
                LineStart = i + 1&
            End If
        Next i
        If LineStart < BytesRead Then
            CarryLen = BytesRead - LineStart
            ReDim Carry(CarryLen - 1)
            CopyMemory Carry(0), Buffer(LineStart), CarryLen
        End If
    Loop
    If CarryLen Then ReadLine.Add StrConv(Carry, vbUnicode)
    CloseHandle hFile
End Function
