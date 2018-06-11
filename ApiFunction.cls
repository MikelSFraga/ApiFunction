VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
<<<<<<< HEAD
=======
'-----------------------------------------------------------------------------------------
' Classe          : ApiFunction Class
' Versão          : v1.0.0
' Data            : 16/05/2017
' Developer       : Mikel S Fraga
' Contato/GitHub  : https://github.com/MikelSFraga/ApiFunction
' Descrição       : Desenvolvimento de funções que utilizam API´s do Windows,
'                   para personalização de Userforms e outras funções diversas.
'-----------------------------------------------------------------------------------------
>>>>>>> develop
Option Explicit
'=========================================================================================
' DECLARAÇÃO DAS FUNÇÕES DAS API´S DO WINDOWS PARA A CLASSE.
#If VBA7 Then         ' Avalia as versões do Office 2010+
  Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
  Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
  Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
  Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
  Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
#ElseIf VBA6 Then     ' Avalia as versões do Office 2007-
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
  Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
  Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
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
  WS_EX_LAYERED = &H80000
End Enum
Enum GetWindowLong
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
End Enum
Enum ShowWindowCommands
  SW_SHOWMAXIMIZED = 3
  SW_SHOWTASKBAR = 5 ' SW_SHOW ORIGINALLY
End Enum
Enum WindowMessages
  WM_SETICON = &H80
End Enum
Enum LayerWindowAtribute
  LWA_ALPHA = &H2&
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
  smTransparency = 6
End Enum
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS CONSTANTES DA CLASSE.
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS/PROPRIEDADES PRIVADAS DA CLASSE.
'Declare Controls
  Private aForm                         As Object
'Declare UDV (User Definition Variable)
  Private aWndStyle                     As WindowStyles
  Private aCmdSWnd                      As ShowWindowCommands
  Private aGetWin                       As GetWindowLong
  Private aLyrWndAtr                    As LayerWindowAtribute
'Declare Validate
  Private aClassName                    As String
  Private aCaption                      As String
  Private aTitleBarEnable               As Boolean
  Private Const aOpacityMax             As Byte = 250
  Private aOpacityNow                   As Byte
'Declare API Values
  Private aHWndForm, aHWndFormChild     As Long
  Private aIStyle                       As Long
  Private aIcon                         As Long
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS/PROPRIEDADES PUBLICAS DA CLASSE.
Public Property Let ActivateButtons(ByVal pWndStyle As WindowStyles)
  'Testa se o atributo TitleBarEnable é True.
  If Not aTitleBarEnable Then Exit Property
  'Aribui o estilo definido no método, para a janela selecionada.
  Call SetWindowStyle(notNullUserform, aCaption, smDualMButton, GWL_STYLE, pWndStyle)
End Property
Public Property Set FormStart(ByVal pForm As Object)
  'Inicia o objeto aForm e captura o seu Caption.
  Set aForm = pForm: aCaption = aForm.Caption
End Property
Public Property Let IconTitleBarForm(ByVal pIconPath As String)
  'Testa se o atributo TitleBarEnable é True.
  If Not aTitleBarEnable Then Exit Property
  'Inseri o ícone informado na barra de título do userform.
  Call SetFormIcon(notNullUserform, aCaption, pIconPath, WM_SETICON)
End Property
Public Property Let OpacityPercent(ByVal pPercent As Byte)
  'Verifica se o valor do parâmetro não ultrapasse 100%.
  If pPercent < 0 Or pPercent > 100 Then
    MsgBox "Valor informado não atende a estrutura desta função."
    Exit Property
  End If
  'Chama rotina de atributo e envia o percentual da transparência.
  Call SetFormTransparency(notNullUserform, aCaption, smTransparency, GWL_EXSTYLE, , WS_EX_LAYERED, pPercent)
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
Public Sub ParentForms(ByVal pChildCaption As String, ByVal pDadCaption As String)
  'Testa se o atributo TitleBarEnable é True.
  If Not aTitleBarEnable Then Exit Sub
  'Obtem a classe a ser utilizada no processo.
  aClassName = SetClassName(notNullUserform)
  'Obtem o identificação das janelas dos formulários que estarão ligados.
  aHWndFormChild = SetHandleWindow(aClassName, pChildCaption)
  aHWndForm = SetHandleWindow(aClassName, pDadCaption)
  'Executa a atribuição/relação entre os Uerforms.
  SetParent aHWndFormChild, aHWndForm
End Sub
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS MÉTODOS PRIVADOS DA CLASSE.
Private Sub SetWindowStyle(ByVal pClass As ClassName, ByVal pCaption As String, _
            ByVal pStyle As StyleMode, ByVal pGetWin As GetWindowLong, _
            Optional ByVal pWndStyle As WindowStyles = 0)
  'Obtem a classe a ser utilizada no processo.
  aClassName = SetClassName(pClass)
  'Obtem o identificado da janela do formulário.
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  If pWndStyle = 0 Then aIStyle = SetStyle(pStyle, pGetWin) _
                   Else aIStyle = SetStyle(pStyle, pGetWin, pWndStyle)
  SetWindowLong aHWndForm, pGetWin, aIStyle
End Sub
Private Sub SetWindowShow(ByVal pClass As ClassName, ByVal pCaption As String, _
                          Optional ByVal pCmdSWnd As ShowWindowCommands = -1)
  'Obtem a classe a ser utilizada no processo.
  aClassName = SetClassName(pClass)
  'Obtem o identificado da janela do formulário.
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  ShowWindow aHWndForm, pCmdSWnd
  DrawMenuBar aHWndForm
  SetFocus aHWndForm
End Sub
Private Sub SetFormIcon(ByVal pClass As ClassName, ByVal pCaption As String, _
                        ByVal pIconPath As String, ByVal pWndMsg As WindowMessages)
  'Obtem o ícone da fonte.
  aIcon = ExtractIcon(0, pIconPath, 0)
  'Obtem a classe a ser utilizada no processo.
  aClassName = SetClassName(pClass)
  'Obtem o identificado da janela do formulário.
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  'Define os ícones grandes (32x32) e pequenos (16x16).
  SendMessage aHWndForm, pWndMsg, True, aIcon
  SendMessage aHWndForm, pWndMsg, False, aIcon
End Sub
Private Sub SetFormTransparency(ByVal pClass As ClassName, _
                                ByVal pCaption As String, ByVal pStyle As StyleMode, _
                                ByVal pGetWin As GetWindowLong, _
                                Optional ByVal pWndStyle As WindowStyles = 0, _
                                Optional ByVal pExWndStyle As ExWindowStyles, _
                                Optional ByVal pPercent As Byte = 100)
  'Faz o calculo proporcional do percentual e o tipo de camada.
  aOpacityNow = VBA.Round(pPercent / 100 * aOpacityMax): aLyrWndAtr = LWA_ALPHA
  'Obtem a classe a ser utilizada no processo.
  aClassName = SetClassName(pClass)
  'Obtem o identificado da janela do formulário.
  aHWndForm = SetHandleWindow(aClassName, pCaption)
  aIStyle = SetStyle(pStyle, pGetWin, , pExWndStyle)
  SetWindowLong aHWndForm, pGetWin, aIStyle
  SetLayeredWindowAttributes aHWndForm, 0, aOpacityNow, aLyrWndAtr
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
                          Optional ByVal pWndStyle As WindowStyles, _
                          Optional ByVal pExWndStyle As ExWindowStyles) As Long
  Select Case pStyle
    Case 1: SetStyle = GetWindowLong(aHWndForm, pGetWin) And &HFFF7FFFF
    Case 2: SetStyle = GetWindowLong(aHWndForm, pGetWin) And Not &HC00000
    Case 3, 4, 5: SetStyle = GetWindowLong(aHWndForm, pGetWin) Or pWndStyle
    Case 6: SetStyle = GetWindowLong(aHWndForm, pGetWin) Or pExWndStyle
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
