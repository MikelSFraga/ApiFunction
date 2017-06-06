VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=========================================================================================
' DECLARAÇÃO DAS FUNÇÕES DAS API´S DO WINDOWS PARA A CLASSE.
#If VBA7 Then         ' Avalia as versões do Office 2010+
  #If Win64 Then      ' Arquitetura 64 Bits do Office
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Longlong
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Longlong
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Longlong
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Longlong, ByVal nCmdShow As Longlong) As Longlong
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As Longlong) As Longlong
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Longlong
  #ElseIf Win32 Then  ' Arquitetura 32 Bits do Office
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
  #End If
#ElseIf VBA6 Then     ' Avalia as versões do Office 2007-
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#End If
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS PARA APIS DA CLASSE.
Enum WindowStyles
  WS_MINIMIZEBOX = &H20000
  WS_MAXIMIZEBOX = &H10000
  WS_FULLSIZING = &H70000
End Enum
Enum ExWindowStyles
  WS_EX_DLGMODALFRAME = &H1
  WS_EX_APPWINDOW = &H40000
End Enum
Enum GetWindowLong
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
End Enum
Enum ShowWindowCommands
  SW_SHOWMAXIMIZED = 3
  SW_SHOWTASKBAR = 5 ' SW_SHOW ORIGINALLY
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
  smDualMButton = 5
  smAppModal = 6
End Enum
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS CONSTANTES DA CLASSE.
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS/PROPRIEDADES PRIVADAS DA CLASSE.
'Declare Controls
  Private aForm As Object
'Declare UDV (User Definition Variable)
  Private aWndStyle As WindowStyles
  Private aCmdSWnd As ShowWindowCommands
  Private aGetWin As GetWindowLong
'Declare Validate
  Private aClassName As String
  Private aCaption As String
  Private aTitleBarEnable As Boolean
'Declare API Values
  Private aHWndForm As Long
  Private aIStyle As Long
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS/PROPRIEDADES PUBLICAS DA CLASSE.
Public Property Set FormStart(ByVal pForm As Object)
  'Inicia o objeto aForm e captura o seu Caption.
  Set aForm = pForm: aCaption = aForm.Caption
End Property
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS MÉTODOS PUBLICOS DA CLASSE.
Public Sub RemoveTitleBar()
  'Testa se o atributo TitleBarEnable é True.
  If Not aTitleBarEnable Then Exit Sub
  'Aribui o estilo definido no método, para a janela selecionada.
  Call SetWindowStyle(notNullUserform, aCaption, smTitleBar, GWL_STYLE)
End Sub
Public Sub HideCloseButton()
  'Testa se o atributo TitleBarEnable é True.
  If Not aTitleBarEnable Then Exit Sub
  'Aribui o estilo definido no método, para a janela selecionada.
  Call SetWindowStyle(NullUserform, aCaption, smCloseButton, GWL_STYLE)
End Sub
Public Sub ActivateButtons(ByVal pWndStyle As WindowStyles)
  'Testa se o atributo TitleBarEnable é True.
  If Not aTitleBarEnable Then Exit Sub
  'Aribui o estilo definido no método, para a janela selecionada.
  Call SetWindowStyle(notNullUserform, aCaption, smDualMButton, GWL_STYLE, pWndStyle)
End Sub
Public Sub ActivateDualButtons()
  'Testa se o atributo TitleBarEnable é True.
  If Not aTitleBarEnable Then Exit Sub
  'Aribui o estilo definido no método, para a janela selecionada.
  Call SetWindowStyle(notNullUserform, aCaption, smDualMButton, GWL_STYLE, WS_FULLSIZING)
End Sub
Public Sub ActivateMaximizeOnly()
  'Testa se o atributo TitleBarEnable é True.
  If Not aTitleBarEnable Then Exit Sub
  'Aribui o estilo definido no método, para a janela selecionada.
  Call SetWindowStyle(notNullUserform, aCaption, smMaximize, GWL_STYLE, WS_MAXIMIZEBOX)
End Sub
Public Sub ActivateMinimizeOnly()
  'Testa se o atributo TitleBarEnable é True.
  If Not aTitleBarEnable Then Exit Sub
  'Aribui o estilo definido no método, para a janela selecionada.
  Call SetWindowStyle(notNullUserform, aCaption, smMinimize, GWL_STYLE, WS_MINIMIZEBOX)
End Sub
Public Sub ShowFormTaskBar()
  'Exibe a janela/Userform na Barra de Tarefas do Windows.
  Call SetWindowShow(notNullUserform, aCaption, SW_SHOWTASKBAR)
End Sub
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS MÉTODOS PRIVADOS DA CLASSE.
Private Sub SetWindowStyle(ByVal pClass As ClassName, ByVal pCaption As String, _
            ByVal pStyle As StyleMode, ByVal pGetWin As GetWindowLong, _
            Optional ByVal pWndStyle As WindowStyles = 0)
  aClassName = SetClassName(pClass)
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  If pWndStyle = 0 Then aIStyle = SetStyle(pStyle, pGetWin) _
                   Else aIStyle = SetStyle(pStyle, pGetWin, pWndStyle)
  SetWindowLong aHWndForm, pGetWin, aIStyle
End Sub
Private Sub SetWindowShow(ByVal pClass As ClassName, ByVal pCaption As String, _
                          Optional ByVal pCmdSWnd As ShowWindowCommands = -1)
  aClassName = SetClassName(pClass)
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  ShowWindow aHWndForm, pCmdSWnd
  DrawMenuBar aHWndForm
  SetFocus aHWndForm
End Sub
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
Private Function SetStyle(ByVal pStyle As StyleMode, ByVal pGetWin As GetWindowLong, _
                          Optional pWndStyle As WindowStyles, _
                          Optional pExWndStyle As ExWindowStyles) As Long
  Select Case pStyle
    Case 1: SetStyle = GetWindowLong(aHWndForm, pGetWin) And &HFFF7FFFF
    Case 2: SetStyle = GetWindowLong(aHWndForm, pGetWin) And Not &HC00000
    Case 3, 4, 5: SetStyle = GetWindowLong(aHWndForm, pGetWin) Or pWndStyle
    Case 6: SetStyle = GetWindowLong(aHWndForm, pGetWin) And Not WS_EX_DLGMODALFRAME Or WS_EX_APPWINDOW 'Or pExWndStyle
  End Select
End Function
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS EVENTOS DA CLASSE.
Private Sub Class_Initialize()
  'Initialize the variables.
  Set aForm = Nothing
  aClassName = ""
  aTitleBarEnable = True
End Sub
Private Sub Class_Terminate()
  'Clear the variables.
  Set aForm = Nothing
End Sub
'=========================================================================================
