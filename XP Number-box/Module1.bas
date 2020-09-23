Attribute VB_Name = "Module1"
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type TOOLINFO
    cbSize As Long
    dwFlags As Long
    hwnd As Long
    dwID As Long
    rtRect As RECT
    hInst As Long
    lpszText As Long
    lParam  As Long
End Type

Public Type ICCEX
    dwSize As Long
    dwICC As Long
End Type

Public Enum EditTipIcon
    etiNone = 0
    etiInfo = 1
    etiWarning = 2
    etiError = 3
End Enum

Public Type EDITBALLOONTIP
    cbStruct As Long
    pszTitle As Long
    pszText As Long
    ttiIcon As Long
End Type

Public Enum TOOLSTYLE
    szClassic = 1
    szBalloon = 64
End Enum

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (ByRef iccInit As ICCEX) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long


' Set Window Pos Flags
Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1

' Init Common Controls
Public Const ICC_WIN95_CLASSES As Long = &HFF

' Misc
Public Const CCM_FIRST As Long = &H2000
Public Const CCM_SETWINDOWTHEME As Long = (CCM_FIRST + &HB)
Public Const WM_USER As Long = &H400
Public Const CW_USEDEFAULT As Long = &H80000000
Public Const ECM_FIRST As Long = &H1500

' Edit Box Tip
Public Const EM_SHOWBALLOONTIP = ECM_FIRST + 3

' Window Styles
Public Const WS_POPUP As Long = &H80000000
Public Const WS_EX_TOPMOST As Long = &H8&


' ToolTips Class
Public Const TOOLTIPS_CLASSA As String = "tooltips_class32"

' ToolTips Flags
Public Const TTF_ABSOLUTE As Long = &H80
Public Const TTF_CENTERTIP As Long = &H2
Public Const TTF_DI_SETITEM As Long = &H8000
Public Const TTF_IDISHWND As Long = &H1
Public Const TTF_RTLREADING As Long = &H4
Public Const TTF_SUBCLASS As Long = &H10
Public Const TTF_TRACK As Long = &H20
Public Const TTF_TRANSPARENT As Long = &H100

' ToolTips Icon
Public Const TTI_ERROR As Long = 3
Public Const TTI_INFO As Long = 1
Public Const TTI_NONE As Long = 0
Public Const TTI_WARNING As Long = 2

' ToolTips Message
Public Const TTM_ACTIVATE As Long = (WM_USER + 1)
Public Const TTM_ADDTOOL As Long = (WM_USER + 4)
Public Const TTM_ADJUSTRECT As Long = (WM_USER + 31)
Public Const TTM_DELTOOL As Long = (WM_USER + 5)
Public Const TTM_ENUMTOOLS As Long = (WM_USER + 14)
Public Const TTM_GETBUBBLESIZE As Long = (WM_USER + 30)
Public Const TTM_GETCURRENTTOOL As Long = (WM_USER + 15)
Public Const TTM_GETDELAYTIME As Long = (WM_USER + 21)
Public Const TTM_GETMARGIN As Long = (WM_USER + 27)
Public Const TTM_GETMAXTIPWIDTH As Long = (WM_USER + 25)
Public Const TTM_GETTEXT As Long = (WM_USER + 11)
Public Const TTM_GETTIPBKCOLOR As Long = (WM_USER + 22)
Public Const TTM_GETTIPTEXTCOLOR As Long = (WM_USER + 23)
Public Const TTM_GETTOOLCOUNT As Long = (WM_USER + 13)
Public Const TTM_GETTOOLINFO As Long = (WM_USER + 8)
Public Const TTM_HITTEST As Long = (WM_USER + 10)
Public Const TTM_NEWTOOLRECT As Long = (WM_USER + 6)
Public Const TTM_POP As Long = (WM_USER + 28)
Public Const TTM_POPUP As Long = (WM_USER + 34)
Public Const TTM_RELAYEVENT As Long = (WM_USER + 7)
Public Const TTM_SETDELAYTIME As Long = (WM_USER + 3)
Public Const TTM_SETMARGIN As Long = (WM_USER + 26)
Public Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)
Public Const TTM_SETTIPBKCOLOR As Long = (WM_USER + 19)
Public Const TTM_SETTIPTEXTCOLOR As Long = (WM_USER + 20)
Public Const TTM_SETTITLE As Long = (WM_USER + 32)
Public Const TTM_SETTOOLINFO As Long = (WM_USER + 9)
Public Const TTM_SETWINDOWTHEME As Long = CCM_SETWINDOWTHEME
Public Const TTM_TRACKACTIVATE As Long = (WM_USER + 17)
Public Const TTM_TRACKPOSITION As Long = (WM_USER + 18)
Public Const TTM_UPDATE As Long = (WM_USER + 29)
Public Const TTM_UPDATETIPTEXT As Long = (WM_USER + 12)
Public Const TTM_WINDOWFROMPOINT As Long = (WM_USER + 16)

' ToolTips Notification
Public Const TTN_FIRST As Long = (-520)
Public Const TTN_GETDISPINFO As Long = (TTN_FIRST - 0)
Public Const TTN_LAST As Long = (-549)
Public Const TTN_LINKCLICK As Long = (TTN_FIRST - 3)
Public Const TTN_NEEDTEXT As Long = TTN_GETDISPINFO
Public Const TTN_POP As Long = (TTN_FIRST - 2)
Public Const TTN_SHOW As Long = (TTN_FIRST - 1)

' ToolTips Creation Flags
Public Const TTS_ALWAYSTIP As Long = &H1
Public Const TTS_BALLOON As Long = &H40
Public Const TTS_NOANIMATE As Long = &H10
Public Const TTS_NOFADE As Long = &H20
Public Const TTS_NOPREFIX As Long = &H2

Global ghWndTip As Long, ghWndParent As Long

Public Function StartTip(hWndParent As Long, Style As Long)
    
    Dim hWndTip As Long, dwFlags As Long, ICEx As ICCEX
    
    dwFlags = TTS_NOPREFIX Or TTS_ALWAYSTIP Or Style
    
    With ICEx
        .dwSize = Len(ICEx)
        .dwICC = ICC_WIN95_CLASSES
    End With
    
    InitCommonControlsEx ICEx
    
    hWndTip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASSA, "", WS_POPUP Or dwFlags, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, hWndParent, 0, App.hInstance, ByVal 0&)
    
    If hWndTip = 0 Then Exit Function
    
    SetWindowPos hWndTip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    StartTip = hWndTip
    ghWndTip = hWndTip
    ghWndParent = hWndParent
    
End Function

Public Sub CreateBalloon(Object1 As Object, hWndOwner As Long, szText As String, Style As TOOLSTYLE, szCentered As Boolean, Optional szTitle As String, Optional mvarIcon As EditTipIcon, Optional BackColor As String, Optional ForeColor As String)
    
Object1.Tag = StartTip(hWndOwner, Style)
    Dim tiInfo As TOOLINFO
    
    With tiInfo
    
        If szCentered = True Then
        
        .dwFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_TRANSPARENT
        
        Else
        
        .dwFlags = TTF_SUBCLASS Or TTF_TRANSPARENT
        
        End If
        
        .hwnd = hWndOwner
        .lpszText = StrPtr(StrConv(szText, vbFromUnicode))
        .hInst = App.hInstance
        GetClientRect hWndOwner, .rtRect
        
        .cbSize = Len(tiInfo)

    End With
    
    If szTitle <> "" Then
    
    SendMessage ghWndTip, TTM_ADDTOOL, 0&, tiInfo
    SendMessage ghWndTip, TTM_SETTITLE, CLng(mvarIcon), ByVal szTitle
    SendMessage ghWndTip, TTM_SETTITLE, CLng(mvarIcon), ByVal szTitle

    Else
    
    SendMessage ghWndTip, TTM_ADDTOOL, 0&, tiInfo
    
    End If
    
    If BackColor <> "" Then

    SendMessage ghWndTip, TTM_SETTIPBKCOLOR, BackColor, 0&
    
    End If
    
    If ForeColor <> "" Then
    
    SendMessage ghWndTip, TTM_SETTIPTEXTCOLOR, ForeColor, 0&
    
    End If
    

    
End Sub

Public Sub ShowBalloonTip(hwndEdit As Long, szTitle As String, szText As String, tipIcon As EditTipIcon, Optional BackColor As String, Optional ForeColor As String)
    
  
    Dim ebtTip As EDITBALLOONTIP
    
    With ebtTip

        .cbStruct = Len(ebtTip)
        .pszText = StrPtr(szText)
        .pszTitle = StrPtr(szTitle)
        .ttiIcon = tipIcon
    End With
    

    
    SendMessage hwndEdit, EM_SHOWBALLOONTIP Or TTF_CENTERTIP, 0&, ebtTip
    
    
End Sub

Public Sub KillBalloonTip(Id As Long)
DestroyWindow Id
End Sub
