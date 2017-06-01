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
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongLong
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As LongLong
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As LongLong
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongLong, ByVal nCmdShow As LongLong) As LongLong
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongLong) As LongLong
  #ElseIf Win32 Then  ' Arquitetura 32 Bits do Office
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
  #End If
#ElseIf VBA6 Then     ' Avalia as versões do Office 2007-
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
#End If
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS PARA APIS DA CLASSE.
Enum WindowStyles
  WS_MINIMIZEBOX = &H20000
  WS_MAXIMIZEBOX = &H10000
  WS_FULLSIZING = &H70000
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
End Enum
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS CONSTANTES DA CLASSE.
Private Const GWL_STYLE As Long = (-16)
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS/PROPRIEDADES PRIVADAS DA CLASSE.
'Declare Controls
  Private aForm As Object
'Declare UDV (User Definition Variable)
  Private aWndStyle As WindowStyles
  Private aCmdSWnd As ShowWindowCommands
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
  If Not aTitleBarEnable Then Exit Sub                                          'Testa se o atributo TitleBarEnable é True.
  Call SetWindowStyle(notNullUserform, aCaption, smTitleBar)                    'Aribui o estilo definido no método, para a janela selecionada.
End Sub
Public Sub HideCloseButton()
  If Not aTitleBarEnable Then Exit Sub                                          'Testa se o atributo TitleBarEnable é True.
  Call SetWindowStyle(NullUserform, aCaption, smCloseButton)                    'Aribui o estilo definido no método, para a janela selecionada.
End Sub
Public Sub ActivateButtons(ByVal pWndStyle As WindowStyles)
  If Not aTitleBarEnable Then Exit Sub                                          'Testa se o atributo TitleBarEnable é True.
  Call SetWindowStyle(notNullUserform, aCaption, smDualMButton, pWndStyle)      'Aribui o estilo definido no método, para a janela selecionada.
End Sub
Public Sub ActivateDualButtons()
  If Not aTitleBarEnable Then Exit Sub                                          'Testa se o atributo TitleBarEnable é True.
  Call SetWindowStyle(notNullUserform, aCaption, smDualMButton, WS_FULLSIZING)  'Aribui o estilo definido no método, para a janela selecionada.
End Sub
Public Sub ActivateMaximizeOnly()
  If Not aTitleBarEnable Then Exit Sub                                          'Testa se o atributo TitleBarEnable é True.
  Call SetWindowStyle(notNullUserform, aCaption, smMaximize, WS_MAXIMIZEBOX)    'Aribui o estilo definido no método, para a janela selecionada.
End Sub
Public Sub ActivateMinimizeOnly()
  If Not aTitleBarEnable Then Exit Sub                                          'Testa se o atributo TitleBarEnable é True.
  Call SetWindowStyle(notNullUserform, aCaption, smMinimize, WS_MINIMIZEBOX)    'Aribui o estilo definido no método, para a janela selecionada.
End Sub

Public Sub ShowWindowForm(Optional ByVal pSetFocus As Boolean = False, _
                          Optional ByVal pSetFullScreen As Boolean = False, _
                          Optional ByVal pSetTaskBar As Boolean = False)
  If pSetFocus Then Call SetWindowShow(notNullUserform, aCaption)
  If pSetFullScreen And Not pSetTaskBar Then Call SetWindowShow(notNullUserform, aCaption, SW_SHOWMAXIMIZED)
  If Not pSetFullScreen And pSetTaskBar Then Call SetWindowShow(notNullUserform, aCaption, SW_SHOWTASKBAR)
End Sub
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS MÉTODOS PRIVADOS DA CLASSE.
Private Sub SetWindowStyle(ByVal pClass As ClassName, ByVal pCaption As String, _
            ByVal pStyle As StyleMode, Optional ByVal pWndStyle As WindowStyles = 0)
  aClassName = SetClassName(pClass)
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  If pWndStyle = 0 Then aIStyle = SetStyle(pStyle) Else aIStyle = SetStyle(pStyle, pWndStyle)
  SetWindowLong aHWndForm, GWL_STYLE, aIStyle
End Sub
Private Sub SetWindowShow(ByVal pClass As ClassName, ByVal pCaption As String, _
            Optional ByVal pCmdSWnd As ShowWindowCommands = 0) '??? Confirmar valor = 0.
  aClassName = SetClassName(pClass)
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  If pCmdSWnd = 0 Then SetFocus aHWndForm Else ShowWindow aHWndForm, pCmdSWnd
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
  Set aForm = Nothing
  aClassName = ""
  aTitleBarEnable = True
End Sub
Private Sub Class_Terminate()
  'Clear the variables.
  Set aForm = Nothing
End Sub
'=========================================================================================
