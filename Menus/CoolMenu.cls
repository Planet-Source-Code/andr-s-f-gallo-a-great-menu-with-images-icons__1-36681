VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoolMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
'Declares for getting the menu(s) form the form
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal ByPosition As Long, ByRef lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CYCAPTION = 4
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CXDLGFRAME = 7
Private Const SM_CYDLGFRAME = 8
Private Const SM_CYVTHUMB = 9
Private Const SM_CXHTHUMB = 10
Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXCURSOR = 13
Private Const SM_CYCURSOR = 14
Private Const SM_CYMENU = 15
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
Private Const SM_CYKANJIWINDOW = 18
Private Const SM_MOUSEPRESENT = 19
Private Const SM_CYVSCROLL = 20
Private Const SM_CXHSCROLL = 21
Private Const SM_DEBUG = 22
Private Const SM_SWAPBUTTON = 23
Private Const SM_RESERVED1 = 24
Private Const SM_RESERVED2 = 25
Private Const SM_RESERVED3 = 26
Private Const SM_RESERVED4 = 27
Private Const SM_CXMIN = 28
Private Const SM_CYMIN = 29
Private Const SM_CXSIZE = 30
Private Const SM_CYSIZE = 31
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33
Private Const SM_CXMINTRACK = 34
Private Const SM_CYMINTRACK = 35
Private Const SM_CXDOUBLECLK = 36
Private Const SM_CYDOUBLECLK = 37
Private Const SM_CXICONSPACING = 38
Private Const SM_CYICONSPACING = 39
Private Const SM_MENUDROPALIGNMENT = 40
Private Const SM_PENWINDOWS = 41
Private Const SM_DBCSENABLED = 42
Private Const SM_CMOUSEBUTTONS = 43
Private Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Private Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Private Const SM_CXSIZEFRAME = SM_CXFRAME
Private Const SM_CYSIZEFRAME = SM_CYFRAME
Private Const SM_SECURE = 44
Private Const SM_CXEDGE = 45
Private Const SM_CYEDGE = 46
Private Const SM_CXMINSPACING = 47
Private Const SM_CYMINSPACING = 48
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
Private Const SM_CYSMCAPTION = 51
Private Const SM_CXSMSIZE = 52
Private Const SM_CYSMSIZE = 53
Private Const SM_CXMENUSIZE = 54
Private Const SM_CYMENUSIZE = 55
Private Const SM_ARRANGE = 56
Private Const SM_CXMINIMIZED = 57
Private Const SM_CYMINIMIZED = 58
Private Const SM_CXMAXTRACK = 59
Private Const SM_CYMAXTRACK = 60
Private Const SM_CXMAXIMIZED = 61
Private Const SM_CYMAXIMIZED = 62
Private Const SM_NETWORK = 63
Private Const SM_CLEANBOOT = 67
Private Const SM_CXDRAG = 68
Private Const SM_CYDRAG = 69
Private Const SM_SHOWSOUNDS = 70
Private Const SM_CXMENUCHECK = 71
Private Const SM_CYMENUCHECK = 72
Private Const SM_SLOWMACHINE = 73
Private Const SM_MIDEASTENABLED = 74
Private Const ILD_NORMAL = &H0
Private Const ILD_TRANSPARENT = &H1
Private Const ILD_MASK = &H10
Private Const ILD_IMAGE = &H20
Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4
Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_ADJ_MAX = 100
Private Const COLOR_ADJ_MIN = -100
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_MENU = 4
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_WINDOWTEXT = 8
Private Const ODA_DRAWENTIRE = &H1
Private Const ODA_SELECT = &H2
Private Const ODA_FOCUS = &H4
Private Const ODS_SELECTED = &H1
Private Const ODS_GRAYED = &H2
Private Const ODS_DISABLED = &H4
Private Const ODS_CHECKED = &H8
Private Const ODS_FOCUS = &H10
Private Const ODS_DEFAULT = &H20
Private Const ODS_COMBOBOXEDIT = &H1000
Private Const LF_FACESIZE = 32
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const GWL_WNDPROC = -4
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_CHARSTREAM = 4
Private Const DT_DISPFILE = 6
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_INTERNAL = &H1000
Private Const DT_LEFT = &H0
Private Const DT_METAFILE = 5
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_PLOTTER = 0
Private Const DT_RASCAMERA = 3
Private Const DT_RASDISPLAY = 1
Private Const DT_RASPRINTER = 2
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const ODT_MENU = 1
Private Const MNC_IGNORE = 0
Private Const MNC_CLOSE = 1
Private Const MNC_EXECUTE = 2
Private Const MNC_SELECT = 3
Private Const MIIM_STATE = &H1&
Private Const MIIM_ID = &H2
Private Const MIIM_SUBMENU = &H4
Private Const MIIM_CHECKMARKS = &H8
Private Const MIIM_TYPE = &H10
Private Const MIIM_DATA = &H20
Private Const MIIM_STRING = &H40
Private Const MIIM_BITMAP = &H80
Private Const MIIM_FTYPE = &H100
Private Const MF_INSERT = &H0
Private Const MF_CHANGE = &H80
Private Const MF_APPEND = &H100
Private Const MF_DELETE = &H200
Private Const MF_REMOVE = &H1000
Private Const MF_BYCOMMAND = &H0
Private Const MF_BYPOSITION = &H400
Private Const MF_SEPARATOR = &H800
Private Const MF_ENABLED = &H0
Private Const MF_GRAYED = &H1
Private Const MF_DISABLED = &H2
Private Const MF_UNCHECKED = &H0
Private Const MF_CHECKED = &H8
Private Const MF_USECHECKBITMAPS = &H200
Private Const MF_STRING = &H0
Private Const MF_BITMAP = &H4
Private Const MF_OWNERDRAW = &H100
Private Const MF_POPUP = &H10
Private Const MF_MENUBARBREAK = &H20
Private Const MF_MENUBREAK = &H40
Private Const MF_UNHILITE = &H0
Private Const MF_HILITE = &H80
Private Const MF_DEFAULT = &H1000
Private Const MF_SYSMENU = &H2000
Private Const MF_HELP = &H4000
Private Const MF_RIGHTJUSTIFY = &H4000
Private Const MF_MOUSESELECT = &H8000
Private Const MF_END = &H80
Private Const MFT_STRING = MF_STRING
Private Const MFT_BITMAP = MF_BITMAP
Private Const MFT_MENUBARBREAK = MF_MENUBARBREAK
Private Const MFT_MENUBREAK = MF_MENUBREAK
Private Const MFT_OWNERDRAW = MF_OWNERDRAW
Private Const MFT_RADIOCHECK = &H200
Private Const MFT_SEPARATOR = MF_SEPARATOR
Private Const MFT_RIGHTORDER = &H2000
Private Const MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY
Private Const MFS_GRAYED = &H3
Private Const MFS_DISABLED = MFS_GRAYED
Private Const MFS_CHECKED = MF_CHECKED
Private Const MFS_HILITE = MF_HILITE
Private Const MFS_ENABLED = MF_ENABLED
Private Const MFS_UNCHECKED = MF_UNCHECKED
Private Const MFS_UNHILITE = MF_UNHILITE
Private Const MFS_DEFAULT = MF_DEFAULT
Private Const CXGAP = 1
Private Const CXTEXTMARGIN = 2
Private Const CXBUTTONMARGIN = 2
Private Const CYBUTTONMARGIN = 2
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_DIAGONAL = &H10
Private Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Private Const BF_MIDDLE = &H800
Private Const BF_SOFT = &H1000
Private Const BF_ADJUST = &H2000
Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000
Private Const WM_SYSCOLORCHANGE = &H15
Private Const WM_NCMOUSEMOVE = &HA0
Private Const WM_COMMAND = &H111
Private Const WM_CLOSE = &H10
Private Const WM_DRAWITEM = &H2B
Private Const WM_GETFONT = &H31
Private Const WM_MEASUREITEM = &H2C
Private Const WM_NCHITTEST = &H84
Private Const WM_MENUSELECT = &H11F
Private Const WM_MENUCHAR = &H120
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_WININICHANGE = &H1A
Private Const WM_SETCURSOR = &H20
Private Const WM_SETTINGCHANGE = WM_WININICHANGE
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Type MEASUREITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemWidth As Long
  itemHeight As Long
  ItemData As Long
End Type
Private Type DRAWITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemAction As Long
  itemState As Long
  hwndItem As Long
  hdc As Long
  rcItem As RECT
  ItemData As Long
End Type
Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As Long
  cch As Long
End Type
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 32
End Type
Private Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type
Private Type IMAGEINFO
  hbmImage As Long
  hbmMask As Long
  Unused1 As Long
  Unused2 As Long
  rcImage As RECT
End Type
Private m_iBitmapWidth As Integer
Private m_ilHandle As Object
Private m_wndHandle As Long
Private m_MenuFont As Long
Private lpPrevWndProc As Long
Private pmds As CMyItemDatas
Private menuHeads() As Long
Private m_bColorEmbossed As Boolean
Public Property Get PrevWndProc() As Long
PrevWndProc = lpPrevWndProc&
End Property
Public Sub ConvertMenu(hMenu As Long, nIndex As Long, bSysMenu As Boolean, bShowButtons As Boolean, Optional Permanent As Boolean = False)
'Sub for convert the menu(s) into graphical menus
On Error GoTo ErrorHandle
Dim I As Long, k As Byte
Dim nItem As Long
Dim dwItemData As Long
Dim Text As String
Dim ByteBuffer() As Byte
Dim info As MENUITEMINFO
nItem& = GetMenuItemCount(hMenu&)
For I& = 0 To nItem& - 1
  ReDim ByteBuffer(0 To 200) As Byte
  For k = 0 To 200
    ByteBuffer(k) = 0
  Next k
  info.fMask = MIIM_DATA Or MIIM_ID Or MIIM_TYPE
  info.dwTypeData = VarPtr(ByteBuffer(0))
  info.cch = UBound(ByteBuffer)
  info.cbSize = LenB(info)
  Call GetMenuItemInfo(hMenu&, I&, True, info)
  dwItemData = info.dwItemData
  If bSysMenu And (info.wID >= &HF000) Then
    GoTo NextGoto
  End If
  info.fMask = 0&
  If bShowButtons Then
    If Not CBool(info.fType And MFT_OWNERDRAW) Then
      info.fType = info.fType Or MFT_OWNERDRAW
      info.fMask = info.fMask Or MIIM_TYPE
      If dwItemData = 0& Then
        info.dwItemData = CLng(pmds.Count + 1)
        info.fMask = info.fMask Or MIIM_DATA
        Dim pmd As CMyItemData
        Set pmd = pmds.Add(, , , CStr(info.dwItemData))
        pmd.fType = info.fType
        pmd.sMenuText = Left(StrConv(ByteBuffer, vbUnicode), info.cch)
        pmd.iButton = GetButtonIndex(pmd.sMenuText)
      Else
        Set pmd = pmds.Item(CStr(dwItemData))
      End If
      pmd.bMainMenu = Permanent
      End If
      If Not Permanent Then Call AddmenuHead(hMenu)
    Else
      Text$ = Left(StrConv(ByteBuffer, vbUnicode), info.cch)
      If info.fType And MFT_OWNERDRAW Then
        info.fType = info.fType And (Not MFT_OWNERDRAW)
        info.fMask = info.fMask Or MIIM_TYPE
        If dwItemData& <> 0& Then Text$ = pmds.Item(CStr(dwItemData&)).sMenuText
        info.cch = BSTRtoLPSTR(Text$, ByteBuffer, info.dwTypeData)
      End If
      If dwItemData <> 0& Then
        info.dwItemData = 0&
        info.fMask = info.fMask Or MIIM_DATA
        pmds.Remove CStr(dwItemData)
      End If
    End If
    If info.fMask Then Call SetMenuItemInfo(hMenu, I, True, info)
NextGoto:
Next I&
Exit Sub
ErrorHandle:
Debug.Print Err.Number; Err.Description
Err.Clear
End Sub
Friend Sub OnInitMenuPopup(hMenu As Long, nIndex As Long, bSysMenu As Boolean)
Call ConvertMenu(hMenu, nIndex, bSysMenu, True)
End Sub
Friend Function OnMenuChar(nChar As Long, nFlags As Long, hMenu As Long) As Long
Dim I As Long
Dim nItem As Long
Dim dwItemData As Long
Dim info As MENUITEMINFO
Dim Count As Integer
Dim iCurrent As Integer
Dim ItemMatch() As Integer
ReDim ItemMatch(0 To 0) As Integer
Count = 0
nItem = GetMenuItemCount(hMenu&)
For I& = 0 To nItem - 1
  info.cbSize = LenB(info)
  info.fMask = MIIM_DATA Or MIIM_TYPE Or MIIM_STATE
  Call GetMenuItemInfo(hMenu, I&, True, info)
  dwItemData = info.dwItemData
  If (info.fType And MFT_OWNERDRAW) And dwItemData& <> 0& Then
    Dim Text As String
    Dim iAmpersand As Integer
    Text$ = pmds(CStr(dwItemData&)).sMenuText
    iAmpersand = InStr(Text$, "&")
    If (iAmpersand > 0) And (UCase(Chr(nChar&)) = UCase(Mid(Text$, iAmpersand + 1, 1))) Then
      If Count > UBound(ItemMatch) Then ReDim Preserve ItemMatch(0 To Count)
      ItemMatch(Count) = I&
      Count = Count + 1
    End If
  End If
  If info.fState And MFS_HILITE Then iCurrent = I&
Next I&
Count = Count - 1
If Count = -1 Then
  OnMenuChar = 0&
  Exit Function
End If
If Count = 0 Then
  
  OnMenuChar = MakeLong(ItemMatch(0), MNC_EXECUTE)
End If
Dim iSelect As Integer
For I& = 0 To Count
  If ItemMatch(I&) = iCurrent Then
    iSelect = I&
    Exit For
  End If
Next I&
OnMenuChar = MakeLong(ItemMatch(iSelect), MNC_SELECT)
End Function
Private Sub DrawMenuText(hdc As Long, rc As RECT, Text As String, Color As Long, Optional bLeftAlign As Boolean = True)
'Draw the menu text
Dim LeftStr As String
Dim RightStr As String
Dim iTabPos As Integer
Dim OldFont As Long
LeftStr$ = Text$
iTabPos = InStr(LeftStr$, Chr(9))
If iTabPos > 0 Then
  RightStr$ = Right$(LeftStr$, Len(LeftStr$) - iTabPos)
  LeftStr$ = Left$(LeftStr$, iTabPos - 1)
End If
Call SetTextColor(hdc&, Color&)
OldFont& = SelectObject(hdc&, GetMenuFont)
Call DrawText(hdc&, LeftStr$, Len(LeftStr$), rc, IIf(bLeftAlign, DT_LEFT, DT_CENTER) Or DT_VCENTER Or DT_SINGLELINE)
If iTabPos > 0 Then Call DrawText(hdc&, RightStr$, Len(RightStr$), rc, DT_RIGHT Or DT_VCENTER Or DT_SINGLELINE)
  Call SelectObject(hdc&, OldFont&)
End Sub
Friend Function OnDrawItem(ByRef dsPtr As Long) As Boolean
'This is the most important sub in this program
'When it draw each sub menu and main menu in the form
'DON'T STOP HERE OR YOUR VISUAL BASIC WILL CRASH
Dim lpds As DRAWITEMSTRUCT
Call CopyMemory(lpds, ByVal dsPtr&, Len(lpds))
Dim rt As RECT
Dim rtItem As RECT
Dim rtText As RECT
Dim rtButn As RECT
Dim rtIcon As RECT
Dim dwColorBG As Long
Dim rtHighlight As RECT
Dim dwItemData As Long
dwItemData& = lpds.ItemData
If (dwItemData& = 0&) Or (lpds.CtlType <> ODT_MENU) Then
  OnDrawItem = False
  Exit Function
End If
Dim pmd As CMyItemData
Set pmd = pmds.Item(CStr(dwItemData&))
Dim hdc As Long
hdc& = lpds.hdc
LSet rtItem = lpds.rcItem
If pmd.fType And MFT_SEPARATOR Then
  LSet rt = rtItem
  rt.Top = rt.Top + ((rt.Bottom - rt.Top) \ 2)
  Call DrawEdge(hdc&, rt, EDGE_ETCHED, BF_TOP)
Else
  Dim bDisabled As Boolean
  Dim bSelected As Boolean
  Dim bChecked As Boolean
  Dim bHaveButn As Boolean
  bDisabled = lpds.itemState And ODS_GRAYED
  bSelected = lpds.itemState And ODS_SELECTED
  bChecked = lpds.itemState And ODS_CHECKED
  bHaveButn = False
  Dim iButton As Integer
  iButton = pmd.iButton
  If iButton > 0 Then
    bHaveButn = True
    LSet rtButn = rtItem
    rtButn.Right = rtButn.Left + m_iBitmapWidth + CXBUTTONMARGIN
    rtIcon.Left = ((rtButn.Right - rtButn.Left) \ 2) - (m_iBitmapWidth \ 2)
    rtIcon.Right = rtIcon.Left + m_iBitmapWidth
    rtIcon.Top = rtButn.Top + ((rtButn.Bottom - rtButn.Top) - m_iBitmapWidth) \ 2
    rtIcon.Bottom = rtIcon.Top + m_iBitmapWidth
    Dim hIcon As Long
    If bDisabled = False Then 'if it is enabled
      hIcon = m_ilHandle.ListImages(iButton).ExtractIcon
      DrawState hdc&, 0&, 0&, hIcon, 0&, rtIcon.Left, rtIcon.Top, rtIcon.Left + m_iBitmapWidth%, rtIcon.Top + m_iBitmapWidth%, DST_ICON Or DSS_NORMAL
    ElseIf bChecked = False Then 'if it is disabled but not checked
      hIcon = m_ilHandle.ListImages(iButton).ExtractIcon
      DrawState hdc&, 0&, 0&, hIcon, 0&, rtIcon.Left, rtIcon.Top, rtIcon.Left + m_iBitmapWidth%, rtIcon.Top + m_iBitmapWidth%, DST_ICON Or DSS_DISABLED
    End If
  Else
  End If
  Dim iButnWidth As Integer
  iButnWidth% = m_iBitmapWidth% + CXBUTTONMARGIN
  dwColorBG = GetSysColor(IIf(bSelected, COLOR_HIGHLIGHT, COLOR_MENU))
  LSet rtText = rtItem
  If pmd.bMainMenu Then Call FillRectEx(hdc&, rtItem, GetSysColor(COLOR_MENU))
  'if it is selected and enabled, the cantrol draws it normal
  If (bSelected Or (lpds.itemAction = ODA_SELECT)) And (Not bDisabled) Then
    LSet rtHighlight = rtItem
    If pmd.bMainMenu And bSelected Then
      Call DrawEdge(hdc&, rtHighlight, BDR_SUNKENOUTER, BF_RECT)
    Else
      Call FillRectEx(hdc&, rtHighlight, dwColorBG&)
    End If
    If pmd.iButton > 0 Then
      hIcon = m_ilHandle.ListImages(pmd.iButton).ExtractIcon
      DrawState hdc&, 0&, 0&, hIcon, 0&, rtIcon.Left, rtIcon.Top, rtIcon.Left + m_iBitmapWidth%, rtIcon.Top + m_iBitmapWidth%, DST_ICON Or DSS_NORMAL
    End If
  'else it draws the image disabled
  ElseIf (bSelected Or (lpds.itemAction = ODA_SELECT)) And bDisabled Then
    LSet rtHighlight = rtItem
    If pmd.bMainMenu And bSelected Then
      Call DrawEdge(hdc&, rtHighlight, BDR_SUNKENOUTER, BF_RECT)
    Else
      Call FillRectEx(hdc&, rtHighlight, dwColorBG&)
    End If
    If pmd.iButton > 0 Then
      hIcon = m_ilHandle.ListImages(iButton).ExtractIcon
      DrawState hdc&, 0&, 0&, hIcon, 0&, rtIcon.Left, rtIcon.Top, rtIcon.Left + m_iBitmapWidth%, rtIcon.Top + m_iBitmapWidth%, DST_ICON Or DSS_DISABLED
    End If
  End If
  'if it is checked
  If bChecked Then
    LSet rtHighlight = rtIcon
    rtHighlight.Top = rtHighlight.Top - 2
    rtHighlight.Left = rtHighlight.Left - 2
    rtHighlight.Bottom = rtHighlight.Bottom + 2
    rtHighlight.Right = rtHighlight.Right + 2
    dwColorBG = GetSysColor(IIf(bChecked, COLOR_HIGHLIGHT, COLOR_MENU))
    FillRectEx hdc, rtHighlight, dwColorBG
    hIcon = m_ilHandle.ListImages(iButton).ExtractIcon
    DrawState hdc&, 0&, 0&, hIcon, 0&, rtIcon.Left, rtIcon.Top, rtIcon.Left + m_iBitmapWidth%, rtIcon.Top + m_iBitmapWidth%, DST_ICON Or DSS_NORMAL
  End If
  If Not pmd.bMainMenu Then
    rtText.Left = rtItem.Left + iButnWidth% + CXGAP + CXTEXTMARGIN
    rtText.Right = rtItem.Right - iButnWidth%
  End If
  Call SetBkMode(hdc&, TRANSPARENT)
  Dim dwColorText As Long
  dwColorText = GetSysColor(IIf(bDisabled, COLOR_GRAYTEXT, IIf(bSelected And (Not pmd.bMainMenu), COLOR_HIGHLIGHTTEXT, COLOR_MENUTEXT)))
  'here it draws the menu text(disbled/normal)
  DrawMenuText hdc&, rtText, pmd.sMenuText, dwColorText&, Not pmd.bMainMenu
End If
CopyMemory ByVal dsPtr&, lpds, Len(lpds)
OnDrawItem = True
End Function
Friend Function OnMeasureItem(ByRef miPtr As Long) As Boolean
Dim lpms As MEASUREITEMSTRUCT
Call CopyMemory(lpms, ByVal miPtr, Len(lpms))
Dim dwItemData As Long
dwItemData& = lpms.ItemData
If dwItemData& = 0 Then Exit Function
Dim pmd As CMyItemData
Set pmd = pmds.Item(CStr(dwItemData&))
If lpms.CtlType <> ODT_MENU Then
  OnMeasureItem = False
  Exit Function
End If
Dim iCYMENU As Integer
iCYMENU% = GetSystemMetrics(SM_CYMENU)
If pmd.fType And MFT_SEPARATOR Then
  lpms.itemHeight = iCYMENU% \ 2
  lpms.itemWidth = 0
Else
  Dim temphdc As Long
  Dim rc As RECT
  Dim rcHeight As Integer
  Dim itemWidth As Long
  Dim OldFont As Long
  temphdc& = GetDC(m_wndHandle&)
  OldFont& = SelectObject(temphdc&, GetMenuFont)
  Call DrawText(temphdc&, pmd.sMenuText, Len(pmd.sMenuText), rc, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_CALCRECT)
  Call SelectObject(temphdc&, OldFont&)
  rcHeight = rc.Bottom - rc.Top
  lpms.itemHeight = IIf(rcHeight > iCYMENU%, rcHeight, iCYMENU%)
  itemWidth& = (rc.Right - rc.Left)
  If Not pmd.bMainMenu Then
    itemWidth& = itemWidth& + (CXTEXTMARGIN * 2) + CXGAP + (m_iBitmapWidth% + CXBUTTONMARGIN) * 2
    itemWidth& = itemWidth& - (GetSystemMetrics(SM_CXMENUCHECK) - 1)
  End If
  lpms.itemWidth = itemWidth&
End If
Call CopyMemory(ByVal miPtr, lpms, Len(lpms))
OnMeasureItem = True
End Function
Friend Function GetMenuFont(Optional bForceReset As Boolean = False) As Long
If (m_MenuFont& = 0) Or bForceReset Then
  Dim whdc As Long
  Dim sText As String
  Dim TextLen As Long
  Dim tLF As LOGFONT
  Dim tm As TEXTMETRIC
  sText$ = Space$(255)
  TextLen& = Len(sText$)
  whdc& = GetDC(m_wndHandle)
  TextLen& = GetTextFace(whdc&, TextLen&, sText$)
  Call GetTextMetrics(whdc&, tm)
  tLF.lfFaceName = Left$(sText$, TextLen& + 1)
  tLF.lfHeight = tm.tmHeight
  tLF.lfWeight = tm.tmWeight
  m_MenuFont& = CreateFontIndirect(tLF)
End If
GetMenuFont& = m_MenuFont&
End Function
Public Function Install(wndHandle As Long, Optional ilHandle As Object, Optional bColorEmbossed As Boolean = True, Optional bMainMenu As Boolean = False) As Boolean
m_iBitmapWidth% = 16
If wndHandle <> 0 Then
  m_wndHandle& = wndHandle&
  m_bColorEmbossed = bColorEmbossed
  lpPrevWndProc& = GetWindowLong(m_wndHandle&, GWL_WNDPROC)
  Call SetWindowLong(m_wndHandle&, GWL_WNDPROC, AddressOf WindowProc)
  Set pmds = New CMyItemDatas
  ReDim menuHeads(0 To 0) As Long
  If bMainMenu Then Call ConvertMenu(GetMenu(m_wndHandle&), 0&, False, True, True)
  If Not (ilHandle Is Nothing) Then
    Set m_ilHandle = ilHandle
  End If
Else
  If m_wndHandle& <> 0 Then
    Call SetWindowLong(m_wndHandle&, GWL_WNDPROC, lpPrevWndProc&)
  End If
  Set pmds = Nothing
  Set m_ilHandle = Nothing
End If
Install = True
End Function
Private Sub FillRectEx(hdc As Long, rc As RECT, Color As Long)
'Draw th back rectangle of a selected menu
Dim OldBrush As Long
Dim NewBrush As Long
NewBrush& = CreateSolidBrush(Color&)
Call FillRect(hdc&, rc, NewBrush&)
Call DeleteObject(NewBrush&)
End Sub
Private Function OffsetRect(InRect As RECT, xOffset As Long, yOffset As Long) As RECT
OffsetRect.Left = InRect.Left + xOffset&
OffsetRect.Right = InRect.Right + xOffset&
OffsetRect.Top = InRect.Top + yOffset&
OffsetRect.Bottom = InRect.Bottom + yOffset&
End Function
Friend Sub OnMenuSelect(nItemID As Integer, nFlags As Integer, hSysMenu As Long)
Dim I As Integer
If (hSysMenu = 0&) And (nFlags = &HFFFF) Then
  For I% = 0 To UBound(menuHeads) - 1
    Call ConvertMenu(menuHeads(I%), 0&, False, False)
  Next I%
End If
End Sub
Private Sub AddmenuHead(hMenu As Long)
Dim I As Integer
Dim Bound As Integer
Bound% = UBound(menuHeads)
For I% = 0 To Bound%
  If menuHeads(I%) = hMenu Then Exit Sub
Next I%
ReDim Preserve menuHeads(0 To Bound% + 1) As Long
menuHeads(Bound%) = hMenu
End Sub
Public Function GetButtonIndex(sMenuText As String) As Integer
Dim sKeyText As String
Dim I As Integer
Dim Pos As Integer
Pos = InStr(1, sMenuText, vbTab, vbTextCompare)
If Pos <> 0 Then
  sMenuText = Left(sMenuText, Pos - 1)
End If
If Not (m_ilHandle Is Nothing) Then
  For I = 1 To m_ilHandle.ListImages.Count
    sKeyText = m_ilHandle.ListImages(I).Key
    If UCase(sKeyText) = UCase(sMenuText) Then
      GetButtonIndex = I
      Exit Function
    End If
  Next I
End If
GetButtonIndex = -1
End Function
Private Function BSTRtoLPSTR(sBSTR As String, b() As Byte, lpsz As Long) As Long
Dim cBytes As Long
Dim sABSTR As String
cBytes = LenB(sBSTR)
ReDim b(1 To cBytes + 2) As Byte
sABSTR = StrConv(sBSTR, vbFromUnicode)
lpsz = StrPtr(sABSTR)
CopyMemory b(1), ByVal lpsz, cBytes + 2
lpsz = VarPtr(b(1))
BSTRtoLPSTR = cBytes
End Function

