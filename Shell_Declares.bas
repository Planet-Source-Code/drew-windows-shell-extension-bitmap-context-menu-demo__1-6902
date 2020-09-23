Attribute VB_Name = "Shell_Declares"
Option Explicit

'----------------------------------
'- Shell Extension API Declares...
'----------------------------------
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpstring As Any) As Long

Public Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal ProgID As String, rClsId As Any) As Long
Public Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
Public Declare Function ReleaseStgMedium Lib "ole32.dll" (pMedium As STGMEDIUM) As Long
Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal pUINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare Function GetShortPathNameA Lib "kernel32" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Public Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As String) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Public Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Public Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function InsertMenuBmp Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long


'----------------------------------
'- Shell Extension Public Types...
'----------------------------------
Public Type STGMEDIUM
    tymed               As Long     ' DWORD
    hGlobal             As Long     ' [case(TYMED_HGLOBAL)]  HGLOBAL        hGlobal;
    pUnkForRelease      As IUnknown ' [unique] IUnknown *pUnkForRelease;
End Type

Public Type FORMATETC
    cfFormat            As Long
    ptd                 As Long
    dwAspect            As Long
    lindex              As Long
    tymed               As Long
End Type

Public Type uGUID       '{xxxx-xx-xx-xxxxxxxx}
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(7)            As Byte
End Type

Public Type TEXTMETRIC
    tmHeight            As Long
    tmAscent            As Long
    tmDescent           As Long
    tmInternalLeading   As Long
    tmExternalLeading   As Long
    tmAveCharWidth      As Long
    tmMaxCharWidth      As Long
    tmWeight            As Long
    tmOverhang          As Long
    tmDigitizedAspectX  As Long
    tmDigitizedAspectY  As Long
    tmFirstChar         As Byte
    tmLastChar          As Byte
    tmDefaultChar       As Byte
    tmBreakChar         As Byte
    tmItalic            As Byte
    tmUnderlined        As Byte
    tmStruckOut         As Byte
    tmPitchAndFamily    As Byte
    tmCharSet           As Byte
End Type

Public Type BITMAP
    bmType              As Long
    bmWidth             As Long
    bmHeight            As Long
    bmWidthBytes        As Long
    bmPlanes            As Integer
    bmBitsPixel         As Integer
    bmBits              As Long
End Type

Public Type CMINVOKECOMMANDINFO
    cbSize              As Long    ' sizeof(CMINVOKECOMMANDINFO)
    fMask               As Long    ' any combination of CMIC_MASK_*
    hwnd                As Long    ' might be NULL (indicating no owner window)
    lpVerb              As Long    ' either a string or MAKEINTRESOURCE(idOffset)
    lpParameters        As Long    ' might be NULL (indicating no parameter)
    lpDirectory         As Long    ' might be NULL (indicating no specific directory)
    nShow               As Long    ' one of SW_ values for ShowWindow() API
    dwHotKey            As Long
    hIcon               As Long
End Type

'----------------------------------------------------------
'- Shell Extension Public Constants...
'----------------------------------------------------------
Public Const CF_HDROP = 15
Public Const DVASPECT_CONTENT = 1
Public Const TYMED_HGLOBAL = 1

Public Const PAGE_NOACCESS = &H1&
Public Const PAGE_READONLY = &H2&
Public Const PAGE_READWRITE = &H4&
Public Const PAGE_WRITECOPY = &H8&
Public Const PAGE_EXECUTE = &H10&
Public Const PAGE_EXECUTE_READ = &H20&
Public Const PAGE_EXECUTE_READWRITE = &H40&
Public Const PAGE_EXECUTE_WRITECOPY = &H80&
Public Const PAGE_GUARD = &H100&
Public Const PAGE_NOCACHE = &H200&
Public Const RT_BITMAP = 2&
Public Const REG_SZ = 1&            ' Unicode null terminated string

'System Parameter Constants
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPIF_UPDATEINIFILE = &H1


' ;win40  -- A lot of MF_* flags have been renamed as MFT_* and MFS_* flags */
' Menu flags for Add/Check/EnableMenuItem()
Public Const MF_INSERT = &H0&
Public Const MF_CHANGE = &H80&
Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_REMOVE = &H1000&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_SEPARATOR = &H800&
Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_DISABLED = &H2&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_USECHECKBITMAPS = &H200&
Public Const MF_STRING = &H0&
Public Const MF_BITMAP = &H4&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public Const MF_UNHILITE = &H0&
Public Const MF_HILITE = &H80&
Public Const MF_DEFAULT = &H1000&
Public Const MF_SYSMENU = &H2000&
Public Const MF_HELP = &H4000&
Public Const MF_RIGHTJUSTIFY = &H4000&
Public Const MF_MOUSESELECT = &H8000&
Public Const MF_END = &H80&                      ' Obsolete -- only used by old RES files */

Public Const MFT_STRING = MF_STRING
Public Const MFT_BITMAP = MF_BITMAP
Public Const MFT_MENUBARBREAK = MF_MENUBARBREAK
Public Const MFT_MENUBREAK = MF_MENUBREAK
Public Const MFT_OWNERDRAW = MF_OWNERDRAW
Public Const MFT_RADIOCHECK = &H200&
Public Const MFT_SEPARATOR = MF_SEPARATOR
Public Const MFT_RIGHTORDER = &H2000&
Public Const MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY

'- Menu flags for Add/Check/EnableMenuItem()
Public Const MFS_GRAYED = &H3&
Public Const MFS_DISABLED = MFS_GRAYED
Public Const MFS_CHECKED = MF_CHECKED
Public Const MFS_HILITE = MF_HILITE
Public Const MFS_ENABLED = MF_ENABLED
Public Const MFS_UNCHECKED = MF_UNCHECKED
Public Const MFS_UNHILITE = MF_UNHILITE
Public Const MFS_DEFAULT = MF_DEFAULT

'- QueryContextMenu uFlags
Public Const CMF_NORMAL = &H0&
Public Const CMF_DEFAULTONLY = &H1&
Public Const CMF_VERBSONLY = &H2&
Public Const CMF_EXPLORE = &H4&
Public Const CMF_NOVERBS = &H8&
Public Const CMF_CANRENAME = &H10&
Public Const CMF_NODEFAULT = &H20&
Public Const CMF_INCLUDESTATIC = &H40&
Public Const CMF_RESERVED = &HFFFF0000          ' View specific


' GetCommandString uFlags
Public Const GCS_VERBA = &H0&                   ' canonical verb
Public Const GCS_HELPTEXTA = &H1&               ' help text (for status bar)
Public Const GCS_VALIDATEA = &H2&               ' validate command exists
Public Const GCS_VERBW = &H4&                   ' canonical verb (Unicode)
Public Const GCS_HELPTEXTW = &H5&               ' help text (Unicode version)
Public Const GCS_VALIDATEW = &H6&               ' validate command exists (Unicode)

Public Const CMDSTR_NEWFOLDER = "NewFolder"
Public Const CMDSTR_VIEWLIST = "ViewList"
Public Const CMDSTR_VIEWDETAILS = "ViewDetails"

'#define SEE_MASK_ICON           0x00000010
'#define SEE_MASK_HOTKEY         0x00000020
'#define SEE_MASK_FLAG_NO_UI     0x00000400
'#define SEE_MASK_UNICODE        0x00004000
'#define SEE_MASK_NO_CONSOLE     0x00008000
'#define SEE_MASK_ASYNCOK        0x00100000
'
'#define CMIC_MASK_HOTKEY        SEE_MASK_HOTKEY
'#define CMIC_MASK_ICON          SEE_MASK_ICON
'#define CMIC_MASK_FLAG_NO_UI    SEE_MASK_FLAG_NO_UI
'#define CMIC_MASK_UNICODE       SEE_MASK_UNICODE
'#define CMIC_MASK_NO_CONSOLE    SEE_MASK_NO_CONSOLE
'#define CMIC_MASK_HASLINKNAME   SEE_MASK_HASLINKNAME
'#define CMIC_MASK_FLAG_SEP_VDM  SEE_MASK_FLAG_SEPVDM
'#define CMIC_MASK_HASTITLE      SEE_MASK_HASTITLE
'#define CMIC_MASK_ASYNCOK       SEE_MASK_ASYNCOK
