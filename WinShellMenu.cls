VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_szFile As String

'-----------------------------
'- Implemented Interfaces...
'-----------------------------
Implements IContextMenu
Implements IShellExtInit

Private Sub Class_Initialize()
    Dim pImp As Long
    Dim oImp As IContextMenu
    
    Set oImp = Me
    Call CopyMemory(pImp, oImp, 4)
    pOldFunction = ReplaceVtableEntry(pImp, 4, AddressOf sc_QueryContextMenu)
End Sub

Private Sub Class_Terminate()
    Dim pImp As Long
    Dim oImp As IContextMenu

    Set oImp = Me
    Call CopyMemory(pImp, oImp, 4)
    pOldFunction = ReplaceVtableEntry(pImp, 4, pOldFunction)
    
    DeleteObject hBitmap(0)
    DeleteObject hBitmap(1)
    DeleteObject hBitmap(2)
End Sub

Private Sub IShellExtInit_Initialize(ByVal pidlFolder As Long, _
                                     ByVal lpdobj As Long, _
                                     ByVal hKeyProgID As Long)
    Dim strVal      As String
    Dim strSize     As Long
    Dim varType     As Long
    Dim pIDataObj   As IUnknown
    Dim oIDataObj   As IDataObject
    Dim FmtEtc      As FORMATETC
    Dim pMedium     As STGMEDIUM
    Dim szTemp      As String
    Dim szTemp2     As String
    Dim iFile       As Long
    Dim iFiles      As Long
    Dim rc          As Long

    '------------------------------------------------------------
    ' Get ptr to IDataObject and determine what file(s) where
    ' Right-Clicked on inside of explorer...
    '------------------------------------------------------------
    CopyMemory pIDataObj, lpdobj, 4         ' Copy dataobject pointer to IUnknown variable
                                            ' WARNING: This does not addref the actual object.
    Set oIDataObj = pIDataObj               ' Make IDataObject Variable QueryInterface from IUnknown
    CopyMemory pIDataObj, 0&, 4             ' Set IUnknown Variable to nothing with calling release.
                                            ' Setting PIDataObj = Nothing will trash Explorer...
    With FmtEtc
        .cfFormat = CF_HDROP
        .ptd = 0
        .dwAspect = DVASPECT_CONTENT
        .lindex = -1
        .tymed = TYMED_HGLOBAL
    End With
    
    oIDataObj.GetData ByVal VarPtr(FmtEtc), ByVal VarPtr(pMedium)
    rc = Err.LastDllError

    iFiles = DragQueryFile(pMedium.hGlobal, &HFFFFFFFF, vbNullString, 0)
    For iFile = 0 To (iFiles - 1)
        szTemp = String(500, 0)
        rc = DragQueryFile(pMedium.hGlobal, iFile, szTemp, Len(szTemp))
        If (rc > 0) Then
            szTemp = Left$(szTemp, rc)
            
            szTemp2 = String(500, 0)
            rc = GetShortPathNameA(szTemp, szTemp2, 500)
            If (rc > 0) Then                                ' Is char. count of ShortPathName is > 0?
                szTemp2 = Left$(szTemp2, rc)
                m_szFile = m_szFile & " " & szTemp2
            End If
        End If
    Next
    
    rc = ReleaseStgMedium(pMedium)
    Set oIDataObj = Nothing
    
    '------------------------------------
    ' Get registry info from hKeyProgID
    '------------------------------------
    strSize = 512                   ' Variable Size
    strVal = String$(strSize, 0)    ' Allocate Variable Space
    varType = REG_SZ
    
    rc = RegQueryValueEx(hKeyProgID, vbNullString, 0&, varType, strVal, strSize)

    If (Asc(Mid(strVal, strSize, 1)) = 0) Then strSize = strSize - 1
    If ((rc = 0) And (strSize > 0)) Then
        strVal = Left(strVal, strSize)
    End If
End Sub

Private Sub IContextMenu_QueryContextMenu(ByVal hMenu As Long, _
                                          ByVal indexMenu As Long, _
                                          ByVal idCmdFirst As Long, _
                                          ByVal idCmdLast As Long, _
                                          ByVal uFlags As Long)
    '------------------------------------------------------------
    ' This sub has been replaced in the IContextMenu VTable with SC_QueryContextMenu
    ' This sub is used as a stub in the Class_Initialize Event
    '------------------------------------------------------------
End Sub

Private Sub IContextMenu_GetCommandString(ByVal idCmd As Long, _
                                          ByVal uFlags As Long, _
                                          ByVal pwReserved As Long, _
                                          ByVal pszName As Long, _
                                          ByVal cchMax As Long)
    'Referencing pwReserved will cause a GPF, it is a null pointer.
    Dim bszName()   As Byte
    Dim cChars      As Long
    Dim szName      As String
    Dim temp        As String

    Select Case uFlags
        Case GCS_VERBA, GCS_VERBW
            Select Case idCmd
                Case 0
                    szName = "(GCS_VERB)MENUITEM1"
            End Select
        Case GCS_HELPTEXTA, GCS_HELPTEXTW
            Select Case idCmd
                Case 0
                    szName = kMENU_ITEM_TEXT
            End Select
        Case GCS_VALIDATEA, GCS_VALIDATEW
            Select Case idCmd
                Case 0
                    szName = "(GCS_VALIDATE)New menu item number 1"
            End Select
    End Select
    
    szName = Left$(szName, cchMax) & vbNullChar
    
    Select Case uFlags
        Case GCS_VERBA, GCS_HELPTEXTA, GCS_VALIDATEA
            If (szName <> "") Then
                bszName = StrConv(szName, vbFromUnicode)
                Call CopyMemory(ByVal pszName, bszName(0), UBound(bszName) + 1)
            End If
        Case GCS_VERBW, GCS_HELPTEXTW, GCS_VALIDATEW
            If (szName <> "") Then
                bszName = szName
                Call CopyMemory(ByVal pszName, bszName(0), UBound(bszName) + 1)
            End If
    End Select
End Sub

Private Sub IContextMenu_InvokeCommand(ByVal lpcmi As Long)
    '--------------------------------------
    ' EXECUTE YOUR PROGRAM/COMMAND HERE...
    '--------------------------------------
    Dim idCmd       As Long
    Dim CMI         As CMINVOKECOMMANDINFO
    Dim Directory   As String
    Dim Parameters  As String
    Dim Verb        As String
    Dim sLen        As Long

    CopyMemory CMI, ByVal lpcmi, Len(CMI)
    
    Directory = String(lstrlen(ByVal CMI.lpDirectory), Chr$(0))
    lstrcpy ByVal Directory, ByVal CMI.lpDirectory

    Parameters = String(lstrlen(ByVal CMI.lpParameters), Chr$(0))
    
    sLen = lstrlen(ByVal CMI.lpParameters)
    If (sLen > 0) Then
        Dim bData() As Byte
        ReDim bData(sLen) As Byte

        CopyMemory bData(0), ByVal CMI.lpParameters, sLen
    End If
    
    lstrcpy ByVal Parameters, ByVal CMI.lpParameters

    ' If the high-order word of lpcmi->lpVerb is not NULL, this
    ' function was called by an application and lpVerb is a command
    ' that should be activated. Otherwise, the shell has called this
    ' function, and the low-order word of lpcmi->lpVerb is the
    ' identifier of the menu item that the user selected.
    If ((CMI.lpVerb And &HFFFF000) = 0) Then
        idCmd = CMI.lpVerb
        Select Case idCmd
            Case 0
                '--------------------------------------
                ' MenuItem 0 selected...
                ' EXECUTE YOUR PROGRAM/COMMAND HERE...
                '--------------------------------------
                
                '--------------------------------------
                ' CHANGE HERE:
                '    Change the following code to run your code,
                '           or call your program...
                '--------------------------------------
                
                'I WANT TO SET THE DESKTOP WALLPAPER...
                SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, Trim$(m_szFile), SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
        
        End Select
    Else
        Verb = String(lstrlen(ByVal CMI.lpVerb), Chr$(0))
        lstrcpy ByVal Verb, ByVal CMI.lpVerb
    End If
End Sub
