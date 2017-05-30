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
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As LongLong
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongLong, ByVal nCmdShow As LongLong) As LongLong
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongLong) As LongLong
  #ElseIf Win32 Then  ' Arquitetura 32 Bits do Office
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
  #End If
#ElseIf VBA6 Then     ' Avalia as versões do Office 2007-
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
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
  SW_SHOW = 5
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
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS CONSTANTES DA CLASSE.
Private Const GWL_STYLE As Long = (-16)
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS/PROPRIEDADES PRIVADAS DA CLASSE.
'Declare Controls
'  Private aForm As MSForms.UserForm
'Declare UDV (User Definition Variable)
  Private aWndStyle As WindowStyles
  Private aCmdSWnd As ShowWindowCommands
'Declare Validate
  Private aClassName As String
'Declare API Values
  Private aHWndForm As Long
  Private aIStyle As Long
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS/PROPRIEDADES PUBLICAS DA CLASSE.
'Public Property Set FormStart(ByVal pForm As MSForms.UserForm): Set aForm = pForm: End Property
Public Property Let WindowFocusForm(ByVal pFocus As Boolean)
  aClassName = SetClassName(notNullUserform)                        'Captura a ClassName, baseada na versão do Office.
  aHWndForm = SetHandleWindow(aClassName, aForm.Caption)            'Define o valor para a variável aHWndForm.
  SetFocus aHWndForm
End Property
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS MÉTODOS PUBLICOS DA CLASSE.
Public Sub RemoveTitleBar(ByVal pCaption As String)
  Call ApplyStyle(notNullUserform, pCaption, smTitleBar)            'Aplica o estilo com a classe do Userform.
End Sub
Public Sub HideCloseButton(ByVal pCaption As String)
  Call ApplyStyle(NullUserform, pCaption, smCloseButton)            'Aplica o estilo com a classe nula.
End Sub
Public Sub ActivateButtons(ByVal pCaption As String, ByVal pWndStyle As WindowStyles)
  Call ApplyStyle(notNullUserform, pCaption, smMinimize, pWndStyle) 'Aplica o estilo com a classe do Userform.
End Sub
Public Sub ActivateMinimizeButton(ByVal pCaption As String)
  aWndStyle = WS_MINIMIZEBOX                                        'Define o estilo do botão que esta sendo usado.
  Call ApplyStyle(notNullUserform, pCaption, smMinimize)            'Aplica o estilo com a classe do Userform.
End Sub
Public Sub ActivateMaximizeButton(ByVal pCaption As String)
  aWndStyle = WS_MAXIMIZEBOX                                        'Define o estilo do botão que esta sendo usado.
  Call ApplyStyle(notNullUserform, pCaption, smMaximize)            'Aplica o estilo com a classe do Userform.
End Sub
Public Sub ActivateDualButtons(ByVal pCaption As String)
  aWndStyle = WS_FULLSIZING                                         'Define o estilo do botão que esta sendo usado.
  Call ApplyStyle(notNullUserform, pCaption, smFullScreen)          'Aplica o estilo fullscreen para o Userform.
End Sub
Public Sub FormWindowState(ByVal pCaption As String, ByVal pSwWndCmd As ShowWindowCommands)
  aClassName = SetClassName(notNullUserform)                        'Captura a ClassName, baseada na versão do Office.
  aHWndForm = SetHandleWindow(aClassName, pCaption)                 'Define o valor para a variável aHWndForm.
  ShowWindow aHWndForm, pSwWndCmd
End Sub
Public Sub WindowShowTaskBar(ByVal pCaption As String)
  Call TaskBarWindow(notNullUserform, pCaption)                     'Aplica o estilo fullscreen para o Userform.
End Sub
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS MÉTODOS PRIVADOS DA CLASSE.
Private Sub ApplyStyle(ByVal pClass As ClassName, ByVal pCaption As String, _
                       ByVal pStyle As StyleMode, Optional ByVal pWndStyle As WindowStyles = 0)
  aClassName = SetClassName(pClass)
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  If pWndStyle = 0 Then aIStyle = SetStyle(pStyle) Else aIStyle = SetStyle(pStyle, pWndStyle)
  SetWindowLong aHWndForm, GWL_STYLE, aIStyle
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
  'Set aForm = Nothing
  aClassName = ""
End Sub
Private Sub Class_Terminate()
  'Clear the variables.
  'Set aForm = Nothing
End Sub
'=========================================================================================
