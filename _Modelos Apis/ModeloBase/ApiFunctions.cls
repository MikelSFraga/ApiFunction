VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=========================================================================================
' DECLARAÇÃO DAS FUNÇÕES DAS API´S DO WINDOWS PARA A CLASSE.
#If VBA7 Then         ' Avalia as versões do Office 2010+
  #If Win64 Then      ' Arquitetura 64 Bits do Office
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongLong
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As LongLong
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As LongLong
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As LongLong
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongLong, lpRect As Rect) As LongLong
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongLong, ByVal nCmdShow As LongLong) As LongLong
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongLong) As LongLong
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As LongLong, ByVal lpfn As LongLong, ByVal hmod As LongLong, ByVal dwThreadId As LongLong) As LongLong
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongLong, ByVal nCode As LongLong, ByVal wParam As LongLong, lParam As Any) As LongLong
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongLong) As LongLong
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As LongLong, ByVal yPoint As LongLong) As LongLong
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As PointApi) As LongLong
  #ElseIf Win32 Then  ' Arquitetura 32 Bits do Office
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As PointApi) As Long
  #End If
#ElseIf VBA6 Then     ' Avalia as versões do Office 2007-
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
  Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
  Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
  Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
  Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
  Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
  Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As PointApi) As Long
#End If
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS PARA APIS DA CLASSE.
Enum SysMetrics
  SM_CXSCREEN = 0 'X Size of screen
  SM_CYSCREEN = 1 'Y Size of Screen
  'SM_CXVSCROLL = 2 'X Size of arrow in vertical scroll bar.
  'SM_CYHSCROLL = 3 'Y Size of arrow in horizontal scroll bar
  'SM_CYCAPTION = 4 'Height of windows caption
  'SM_CXBORDER = 5 'Width of no-sizable borders
  'SM_CYBORDER = 6 'Height of non-sizable borders
  'SM_CXDLGFRAME = 7 'Width of dialog box borders
  'SM_CYDLGFRAME = 8 'Height of dialog box borders
  'SM_CYVTHUMB = 9 'Height of scroll box on horizontal scroll bar
  'SM_CXHTHUMB = 10 ' Width of scroll box on horizontal scroll bar
  'SM_CXICON = 11 'Width of standard icon
  'SM_CYICON = 12 'Height of standard icon
  'SM_CXCURSOR = 13 'Width of standard cursor
  'SM_CYCURSOR = 14 'Height of standard cursor
  'SM_CYMENU = 15 'Height of menu
  'SM_CXFULLSCREEN = 16 'Width of client area of maximized window
  'SM_CYFULLSCREEN = 17 'Height of client area of maximized window
  'SM_CYKANJIWINDOW = 18 'Height of Kanji window
  'SM_MOUSEPRESENT = 19 'True is a mouse is present
  'SM_CYVSCROLL = 20 'Height of arrow in vertical scroll bar
  'SM_CXHSCROLL = 21 'Width of arrow in vertical scroll bar
  'SM_DEBUG = 22 'True if deugging version of windows is running
  'SM_SWAPBUTTON = 23 'True if left and right buttons are swapped.
  'SM_CXMIN = 28 'Minimum width of window
  'SM_CYMIN = 29 'Minimum height of window
  'SM_CXSIZE = 30 'Width of title bar bitmaps
  'SM_CYSIZE = 31 'height of title bar bitmaps
  'SM_CXMINTRACK = 34 'Minimum tracking width of window
  'SM_CYMINTRACK = 35 'Minimum tracking height of window
  'SM_CXDOUBLECLK = 36 'double click width
  'SM_CYDOUBLECLK = 37 'double click height
  'SM_CXICONSPACING = 38 'width between desktop icons
  'SM_CYICONSPACING = 39 'height between desktop icons
  'SM_MENUDROPALIGNMENT = 40 'Zero if popup menus are aligned to the left of the memu bar item. True if it is aligned to the right.
  'SM_PENWINDOWS = 41 'The handle of the pen windows DLL if loaded.
  'SM_DBCSENABLED = 42 'True if double byte characteds are enabled
  'SM_CMOUSEBUTTONS = 43 'Number of mouse buttons.
  'SM_CMETRICS = 44 'Number of system metrics
  'SM_CLEANBOOT = 67 'Windows 95 boot mode. 0 = normal, 1 = safe, 2 = safe with network
  'SM_CXMAXIMIZED = 61 'default width of win95 maximised window
  'SM_CXMAXTRACK = 59 'maximum width when resizing win95 windows
  'SM_CXMENUCHECK = 71 'width of menu checkmark bitmap
  'SM_CXMENUSIZE = 54 'width of button on menu bar
  'SM_CXMINIMIZED = 57 'width of rectangle into which minimised windows must fit.
  'SM_CYMAXIMIZED = 62 'default height of win95 maximised window
  'SM_CYMAXTRACK = 60 'maximum width when resizing win95 windows
  'SM_CYMENUCHECK = 72 'height of menu checkmark bitmap
  'SM_CYMENUSIZE = 55 'height of button on menu bar
  'SM_CYMINIMIZED = 58 'height of rectangle into which minimised windows must fit.
  'SM_CYSMCAPTION = 51 'height of windows 95 small caption
  'SM_MIDEASTENABLED = 74 'Hebrw and Arabic enabled for windows 95
  'SM_NETWORK = 63 'bit o is set if a network is present.   SM_SECURE = 44 'True if security is present on windows 95 system
  'SM_SLOWMACHINE = 73 'true if machine is too slow to run win95
End Enum
Enum ClassStyles
  'CS_VREDRAW = &H1
  'CS_HREDRAW = &H2
  'CS_KEYCVTWINDOW = &H4
  'CS_DBLCLKS = &H8
  'CS_OWNDC = &H20
  'CS_CLASSDC = &H40
  'CS_PARENTDC = &H80
  'CS_NOKEYCVT = &H100
  'CS_NOCLOSE = &H200
  'CS_SAVEBITS = &H800
  'CS_BYTEALIGNCLIENT = &H1000
  'CS_BYTEALIGNWINDOW = &H2000
  'CS_PUBLICCLASS = &H4000
End Enum
Enum WindowStyles
  'WS_OVERLAPPED = &H0&
  'WS_POPUP = &H80000000
  'WS_CHILD = &H40000000
  'WS_MINIMIZE = &H20000000
  'WS_VISIBLE = &H10000000
  'WS_DISABLED = &H8000000
  'WS_CLIPSIBLINGS = &H4000000
  'WS_CLIPCHILDREN = &H2000000
  'WS_MAXIMIZE = &H1000000
  'WS_CAPTION = &HC00000                  ' WS_BORDER Or WS_DLGFRAME
  'WS_BORDER = &H800000
  'WS_DLGFRAME = &H400000
  'WS_VSCROLL = &H200000
  'WS_HSCROLL = &H100000
  'WS_SYSMENU = &H80000
  'WS_THICKFRAME = &H40000
  'WS_GROUP = &H20000
  'WS_TABSTOP = &H10000
  WS_MINIMIZEBOX = &H20000
  WS_MAXIMIZEBOX = &H10000
  WS_FULLSIZING = &H70000
  'WS_TILED = WS_OVERLAPPED
  'WS_ICONIC = WS_MINIMIZE
  'WS_SIZEBOX = WS_THICKFRAME
  'WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
  'WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
  'WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
  'WS_CHILDWINDOW = (WS_CHILD)
End Enum
Enum ExWindowStyles
  'WS_EX_APPWINDOW = &H40000
  'WS_EX_TOOLWINDOW = &H80
  'WS_EX_DLGMODALFRAME = &H1&
  'WS_EX_NOPARENTNOTIFY = &H4&
  'WS_EX_TOPMOST = &H8&
  'WS_EX_ACCEPTFILES = &H10&
  'WS_EX_TRANSPARENT = &H20&
End Enum
Enum ColorConstants
  'COLOR_SCROLLBAR = 0
  'COLOR_BACKGROUND = 1
  'COLOR_ACTIVECAPTION = 2
  'COLOR_INACTIVECAPTION = 3
  'COLOR_MENU = 4
  'COLOR_WINDOW = 5
  'COLOR_WINDOWFRAME = 6
  'COLOR_MENUTEXT = 7
  'COLOR_WINDOWTEXT = 8
  'COLOR_CAPTIONTEXT = 9
  'COLOR_ACTIVEBORDER = 10
  'COLOR_INACTIVEBORDER = 11
  'COLOR_APPWORKSPACE = 12
  'COLOR_HIGHLIGHT = 13
  'COLOR_HIGHLIGHTTEXT = 14
  'COLOR_BTNFACE = 15
  'COLOR_BTNSHADOW = 16
  'COLOR_GRAYTEXT = 17
  'COLOR_BTNTEXT = 18
  'COLOR_INACTIVECAPTIONTEXT = 19
  'COLOR_BTNHIGHLIGHT = 20
End Enum
Enum WindowMessages
  'WM_NULL = &H0
  'WM_CREATE = &H1
  'WM_DESTROY = &H2
  'WM_MOVE = &H3
  'WM_SIZE = &H5
End Enum
Enum ShowWindowCommands
  'SW_HIDE = 0
  'SW_SHOWNORMAL = 1
  'SW_NORMAL = 1
  'SW_SHOWMINIMIZED = 2
  SW_SHOWMAXIMIZED = 3
  'SW_MAXIMIZE = 3
  'SW_SHOWNOACTIVATE = 4
  'SW_SHOW = 5
  'SW_MINIMIZE = 6
  'SW_SHOWMINNOACTIVE = 7
  'SW_SHOWNA = 8
  'SW_RESTORE = 9
  'SW_SHOWDEFAULT = 10
  'SW_MAX = 10
End Enum
Enum StandardIdCursors
  'IDC_ARROW = 32512&
  'IDC_IBEAM = 32513&
  'IDC_WAIT = 32514&
  'IDC_CROSS = 32515&
  'IDC_UPARROW = 32516&
  'IDC_SIZE = 32640&
  'IDC_ICON = 32641&
  'IDC_SIZENWSE = 32642&
  'IDC_SIZENESW = 32643&
  'IDC_SIZEWE = 32644&
  'IDC_SIZENS = 32645&
  'IDC_SIZEALL = 32646&
  'IDC_NO = 32648&
  'IDC_APPSTARTING = 32650&
End Enum
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS PERSONALIZADAS DA CLASSE.
Enum ClassName
  NullUserform = -1
  notNullUserform = 0
  XLMAIN = 1
End Enum
Enum StyleMode
  smCloseButton = 1
  smTitleBar = 2
  smMinimize = 3
  smMaximize = 4
  smFullScreen = 5
End Enum
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type PointApi
    X As Long
    Y As Long
End Type
Private Type MouseHookStruct
    pt As PointApi
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS CONSTANTES DA CLASSE.
Private Const GWL_STYLE As Long = (-16)
'Private Const GWL_EXSTYLE As Long = (-20)
'Private Const SC_CLOSE As Long = &HF060
'Private Const WM_SETICON = &H80
Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0
Private Const GWL_HINSTANCE As Long = (-6)
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS/PROPRIEDADES PRIVADAS DA CLASSE.

'Declare Controls
  Private aForm As MSForms.UserForm
  Private aList As MSForms.Control
'Declare UDV (User Definition Variable)
  Private aRect As Rect
  Private aWndStyle As WindowStyles
  Private aCmdSWnd As ShowWindowCommands
  Private aPt As PointApi
'Declare Validate
  Private aMbHideTittleBar As Boolean
  Private aMbHook As Boolean
  Private aClassName As String
'Declare API Values
  Private aHWndForm As Long
  Private aHWndMenu As Long
  Private aHwndListBox As Long
  Private aLngMouseHook As Long
  Private aIStyle As Long
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS/PROPRIEDADES PUBLICAS DA CLASSE.
Public Property Set FormStart(ByVal pForm As MSForms.UserForm): Set aForm = pForm: End Property
Public Property Set FormList(ByVal pList As MSForms.Control): Set aList = pList: End Property
Public Property Let RemoveTitleBar(ByVal pRemove As Boolean)
  If Not pRemove Then Exit Property                               'Se parâmetro for falso, sai da propriedade.
  Call ApplyStyle(notNullUserform, aForm.Caption, smTitleBar)     'Aplica o estilo com a classe do Userform.
  aMbHideTittleBar = pRemove                                      'Modifica validador da barra de título do Userform.
End Property
Public Property Let HideCloseButton(ByVal pHide As Boolean)
  If Not pHide Or aMbHideTittleBar Then Exit Property             'Se parâmetro for falso, sai da propriedade.
  Call ApplyStyle(NullUserform, aForm.Caption, smCloseButton)     'Aplica o estilo com a classe nula.
  aMbHideTittleBar = Not pHide                                    'Modifica validador botão close do Userform.
End Property
Public Property Let ActivateMinimizeButton(ByVal pActivate As Boolean)
  If Not pActivate Or aMbHideTittleBar Then Exit Property         'Se parâmetro for falso, sai da propriedade.
  aWndStyle = WS_MINIMIZEBOX                                      'Define o estilo do botão que esta sendo usado.
  Call ApplyStyle(notNullUserform, aForm.Caption, smMinimize)     'Aplica o estilo com a classe do Userform.
  aMbHideTittleBar = Not pActivate                                'Modifica validador botão minimize do Userform.
End Property
Public Property Let ActivateMaximizeButton(ByVal pActivate As Boolean)
  If Not pActivate Or aMbHideTittleBar Then Exit Property         'Se parâmetro for falso, sai da propriedade.
  aWndStyle = WS_MAXIMIZEBOX                                      'Define o estilo do botão que esta sendo usado.
  Call ApplyStyle(notNullUserform, aForm.Caption, smMaximize)     'Aplica o estilo com a classe do Userform.
  aMbHideTittleBar = Not pActivate                                'Modifica validador botão maximize do Userform.
End Property
Public Property Let ActivateDualButtons(ByVal pActivate As Boolean)
  If Not pActivate Or aMbHideTittleBar Then Exit Property         'Se parâmetro for falso, sai da propriedade.
  aWndStyle = WS_FULLSIZING                                       'Define o estilo do botão que esta sendo usado.
  Call ApplyStyle(notNullUserform, aForm.Caption, smFullScreen)   'Aplica o estilo fullscreen para o Userform.
  aMbHideTittleBar = Not pActivate                                'Modifica validador botão maximize do Userform.
End Property
Public Property Let WindowShowTaskBar(ByVal pActivate As Boolean)
  If Not pActivate Then Exit Property                             'Se parâmetro for falso, sai da propriedade.
  Call TaskBarWindow(notNullUserform, aForm.Caption)              'Aplica o estilo fullscreen para o Userform.
End Property
Public Property Let WindowFocusForm(ByVal pFocus As Boolean)
  If Not pFocus Then Exit Property                                'Se parâmetro for falso, sai da propriedade.
  aClassName = SetClassName(notNullUserform)                      'Captura a ClassName, baseada na versão do Office.
  aHWndForm = SetHandleWindow(aClassName, aForm.Caption)          'Define o valor para a variável aHWndForm.
  SetFocus aHWndForm
End Property
Private Property Let FormWindowSizeExcel(ByVal pActivate As Boolean)
  'If Not pActivate Then Exit Property                            'Se parâmetro for falso, sai da propriedade.
  'Call CaptureSizeApp(XLMAIN, Application.Caption)               'Captura tamanho das medidas da janela do Excel.
End Property
Public Property Let FormWindowState(ByVal pDefine As ShowWindowCommands)
  aClassName = SetClassName(notNullUserform)                      'Captura a ClassName, baseada na versão do Office.
  aHWndForm = SetHandleWindow(aClassName, aForm.Caption)          'Define o valor para a variável aHWndForm.
  ShowWindow aHWndForm, pDefine
End Property
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS MÉTODOS PUBLICOS DA CLASSE.
Private Function DisplaySize(ByVal pMetrics As SysMetrics) As Long: DisplaySize = GetSystemMetrics(pMetrics): End Function
Public Sub HookListBoxScroll(pForm As MSForms.UserForm, pList As MSForms.Control)
  
End Sub
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS MÉTODOS PRIVADOS DA CLASSE.
Private Sub ApplyStyle(ByVal pClass As ClassName, ByVal pCaption As String, ByVal pStyle As StyleMode)
  aClassName = SetClassName(pClass)
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  If aWndStyle = 0 Then aIStyle = SetStyle(pStyle) Else aIStyle = SetStyle(pStyle, aWndStyle)
  SetWindowLong aHWndForm, GWL_STYLE, aIStyle
End Sub
Private Sub CaptureSizeApp(ByVal pClass As ClassName, ByVal pCaption As String)
  aClassName = SetClassName(pClass)
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  GetWindowRect aHWndForm, aRect
End Sub
Private Function TaskBarWindow(ByVal pClass As ClassName, ByVal pCaption As String)
  aClassName = SetClassName(pClass): aCmdSWnd = SW_SHOW
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  ShowWindow aHWndForm, aCmdSWnd
End Function
Private Function SetClassName(ByVal pClass As ClassName) As String
  Select Case pClass
    Case -1: SetClassName = VBA.vbNullString
    Case 0: SetClassName = VBA.IIf(VBA.Val(Application.Version) < 9, "ThunderXFrame", "ThunderDFrame")
    Case 1: SetClassName = "XLMAIN"
  End Select
End Function
Private Function SetHandleWindow(ByVal pClass As String, ByVal pCaptionWindow As String) As Long
  SetHandleWindow = FindWindow(pClass, pCaptionWindow)
End Function
Private Function SetStyle(ByVal pStyle As StyleMode, Optional pWndStyle As WindowStyles) As Long
  Select Case pStyle
    Case 1: SetStyle = GetWindowLong(aHWndForm, GWL_STYLE) And &HFFF7FFFF
    Case 2: SetStyle = GetWindowLong(aHWndForm, GWL_STYLE) And Not &HC00000
    Case 3, 4, 5: SetStyle = GetWindowLong(aHWndForm, GWL_STYLE) Or pWndStyle
  End Select
End Function
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS EVENTOS DA CLASSE.
Private Sub Class_Initialize()
  'Initialize the variables.
  aMbHideTittleBar = False
  aClassName = ""
End Sub
Private Sub Class_Terminate()
  'Clear the variables.
  Set aList = Nothing
  Set aForm = Nothing
End Sub
'=========================================================================================

