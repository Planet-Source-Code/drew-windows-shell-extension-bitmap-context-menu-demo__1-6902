Attribute VB_Name = "Shell_Functions"
Option Explicit

'----------------------------------------------------------
' Shell Extension Constants...
'----------------------------------------------------------
' CHANGE HERE:
'    HERE YOU CAN CHANGE THE
'       * PROGRAM'S CLASS NAME
'       * MENU TEXT
'----------------------------------------------------------
Private Const kPROGRAM_CLASS_NAME = "BitmapShellMenu.ShellExt"
 Public Const kMENU_ITEM_TEXT = "Set Wallpaper"
 
Private Const KMENU_ITEM_IMAGE = 101&

'----------------------------------------------------------
' Shell Extension Public Variable...
'----------------------------------------------------------
Public hBitmap(5)   As Long
Public pOldFunction As Long


Sub Main()
    Dim ClassId As String

    LoadLibrary "oleaut32.dll"
    GetClassId kPROGRAM_CLASS_NAME, ClassId
End Sub

Public Function ReplaceVtableEntry(pObj As Long, _
                                   EntryNumber As Integer, _
                                   ByVal lpfn As Long) As Long
    Dim lOldAddr        As Long
    Dim lpVtableHead    As Long
    Dim lpfnAddr        As Long
    Dim lOldProtect     As Long
    
    CopyMemory lpVtableHead, ByVal pObj, 4
    lpfnAddr = lpVtableHead + (EntryNumber - 1) * 4
    CopyMemory lOldAddr, ByVal lpfnAddr, 4
    
    Call VirtualProtect(lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect)
    CopyMemory ByVal lpfnAddr, lpfn, 4
    Call VirtualProtect(lpfnAddr, 4, lOldProtect, lOldProtect)
    
    ReplaceVtableEntry = lOldAddr
End Function

Private Sub GetClassId(ByVal ProgID As String, ClassId As String)
    Dim rc          As Long
    Dim IdLen       As Integer
    Dim rClsId      As uGUID
    Dim sProgId     As String
    Dim bClassId()  As Byte
    
    sProgId = StrConv(ProgID, vbUnicode)
    rc = CLSIDFromProgID(sProgId, rClsId)
    
    If (rc < 0) Then
        ClassId = "Bad ProgId rc::" & Str$(rc)
        Exit Sub
    End If

    IdLen = 40
    bClassId = String$(IdLen, 0)
    rc = StringFromGUID2(rClsId, VarPtr(bClassId(0)), IdLen)
    
    If (rc <= 0) Then
        ClassId = "Bad ClassID rc::" & Str$(rc)
        Exit Sub
    End If
    ClassId = bClassId

    If (Asc(Mid$(ClassId, rc, 1)) = 0) Then rc = rc - 1
    ClassId = Left$(ClassId, rc)
End Sub

Public Function sc_QueryContextMenu(ByVal This As IContextMenu, _
                                    ByVal hMenu As Long, _
                                    ByVal indexMenu As Long, _
                                    ByVal idCmdFirst As Long, _
                                    ByVal idCmdLast As Long, _
                                    ByVal uFlags As Long) As Long
    Dim SE              As ShellExt
    Dim rc              As Long
    Dim idCmd           As Long
    Dim szMenu          As String
    Dim szMenuText      As String    ' [64]
    Dim bAppendItems    As Boolean
    Dim szTemp          As String    ' [32]
    Dim bBitmap()       As Byte
    Dim mnuPict         As Picture
    
    idCmd = idCmdFirst
    bAppendItems = True

    If ((uFlags And &HF&) = CMF_NORMAL) Then     ' Check == here, since CMF_NORMAL=0
        szMenuText = " &" & kMENU_ITEM_TEXT
    ElseIf ((uFlags And CMF_VERBSONLY) = CMF_VERBSONLY) Then
        szMenuText = " &" & kMENU_ITEM_TEXT
    ElseIf ((uFlags And CMF_EXPLORE) = CMF_EXPLORE) Then
        szMenuText = " &" & kMENU_ITEM_TEXT
    ElseIf ((uFlags And CMF_DEFAULTONLY) = CMF_DEFAULTONLY) Then
        bAppendItems = False
    Else
        bAppendItems = False
    End If
    
    If bAppendItems Then
        '=======================================================================
        ' Insert 1st Menu...
        '=======================================================================
        Call InsertMenu(hMenu, indexMenu, MF_USECHECKBITMAPS Or MF_STRING Or MF_BYPOSITION, idCmd, szMenuText)
        
        ' Add checked bitmap...
        Set mnuPict = LoadResPicture(KMENU_ITEM_IMAGE, vbResBitmap)
        hBitmap(0) = ScaleBMP(mnuPict.Handle)
        Set mnuPict = Nothing
        Call SetMenuItemBitmaps(hMenu, indexMenu, MF_USECHECKBITMAPS Or MF_BITMAP Or MF_BYPOSITION, hBitmap(0), hBitmap(0))

        ' Increment Index and menu count...
        indexMenu = indexMenu + 1
        idCmd = idCmd + 1
        

        '=======================================================================
        ' Insert Menu Separator...
        '=======================================================================
        Call InsertMenu(hMenu, indexMenu, MF_SEPARATOR Or MF_BYPOSITION, 0, vbNullString)
        indexMenu = indexMenu + 1
        
        sc_QueryContextMenu = (idCmd - idCmdFirst)      ' Must return number of menu items inserted
    Else
        sc_QueryContextMenu = 0                         ' Must return number of menu items inserted
    End If
End Function

Public Function ScaleBMP(hBitmap1 As Long) As Long
    Dim hdc         As Long
    Dim hBitmap2    As Long
    Dim tm          As TEXTMETRIC
    Dim bm1         As BITMAP
    Dim bm2         As BITMAP
    Dim hdcMem1     As Long
    Dim hdcMem2     As Long
    Dim hBmOld1     As Long
    Dim hBmOld2     As Long
    
    hdc = CreateIC("DISPLAY", vbNullChar, vbNullChar, vbNullChar)
    Call GetTextMetrics(hdc, tm)
    
    hdcMem1 = CreateCompatibleDC(hdc)
    hdcMem2 = CreateCompatibleDC(hdc)
    DeleteDC hdc
  
    GetObject hBitmap1, LenB(bm1), bm1
  
    LSet bm2 = bm1
    
    bm2.bmWidth = tm.tmMaxCharWidth
    bm2.bmHeight = tm.tmHeight
    
    bm2.bmWidthBytes = (((bm2.bmWidth * bm2.bmBitsPixel) + 15) \ 16) * 2
  
    hBitmap2 = CreateBitmapIndirect(bm2)
  
    hBmOld1 = SelectObject(hdcMem1, hBitmap1)
    hBmOld2 = SelectObject(hdcMem2, hBitmap2)
  
    Call StretchBlt(hdcMem2, 0, 0, bm2.bmWidth, bm2.bmHeight, hdcMem1, 0, 0, bm1.bmWidth, bm1.bmHeight, vbSrcCopy)
  
    SelectObject hdcMem1, hBmOld1
    SelectObject hdcMem2, hBmOld2
    DeleteDC hdcMem1
    DeleteDC hdcMem2

    ScaleBMP = hBitmap2
End Function

Public Function CopyBMP(hBitmap1 As Long) As Long
    Dim hdc         As Long
    Dim hBitmap2    As Long
    Dim bm1         As BITMAP
    Dim bm2         As BITMAP
    Dim hdcMem1     As Long
    Dim hdcMem2     As Long
    Dim hBmOld1     As Long
    Dim hBmOld2     As Long
    
    hdc = CreateIC("DISPLAY", vbNullChar, vbNullChar, vbNullChar)
    
    hdcMem1 = CreateCompatibleDC(hdc)
    hdcMem2 = CreateCompatibleDC(hdc)
    DeleteDC hdc
  
    GetObject hBitmap1, LenB(bm1), bm1
  
    LSet bm2 = bm1
    hBitmap2 = CreateBitmapIndirect(bm2)
  
    hBmOld1 = SelectObject(hdcMem1, hBitmap1)
    hBmOld2 = SelectObject(hdcMem2, hBitmap2)
  
    Call StretchBlt(hdcMem2, 0, 0, bm2.bmWidth, bm2.bmHeight, hdcMem1, 0, 0, bm1.bmWidth, bm1.bmHeight, vbSrcCopy)
  
    SelectObject hdcMem1, hBmOld1
    SelectObject hdcMem2, hBmOld2
    DeleteDC hdcMem1
    DeleteDC hdcMem2

    CopyBMP = hBitmap2
End Function


