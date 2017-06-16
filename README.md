# ApiFunction Class V1.0.0
	
Com o intuito de facilitar o desenvolvimento de ferramentas, utilizando o processo de programação em __VBA do pacote Microsoft Office__, iniciou-se o desenvolvimeno desta __Classe__.

Seu objetivo é a unificação de várias rotinas que utilizam as __Funções dos Windows (API´s)__, através das quais é possível realizar alterações na estrutura dos __Formulários (Userform´s)__ que utilizamos na estrutura do __Visual Basic for Application (VBA)__.

Esta __Classe__ atenderá as diferentes _Arquiteturas do Sistema Operacional Windows (32 bits e 64 bits)_, como também as _Arquiteturas do Pacote Microsoft Office (VBA6 e VBA7)_.

### Funções Windows (API´s)

Nesta versão atual do __ApiFunction (v1.0.0)__, são utilizadas as seguintes __Funções/Api´s do Windows__:

- __FindWindow (user32.dll):__ A função FindWindow recupera o identificador da janela que possui o nome da classe e da janela  combinando com textos específicos. Esta função não pesquisa janelas dependentes.

- __GetWindowLong (user32.dll):__ A função GetWindowLong recupara informação sobre a janela especificada. A função também recupera valores 32-bit (long) específico de uma janela extra da memória de uma janela.

- __SetWindowLong (user32.dll):__ A função SetWindowLong modifica um atribudo da janela específica. A função também define  valores 32-bit (long) específico de uma janela extra da memória de uma janela.

- __ShowWindow (user32.dll):__ A função ShowWindow define o status específico de exibição da janela.

- __SetFocus (user32.dll):__ A função SetFocus define o foco do cursor para a janela especificada. A janela deveria estar associada com a fila de mensagens do threads de chamadas.

- __DrawMenuBar (user32.dll):__ A função DrawMenuBar redesenha a barra de menu da janela especificada. Se a barra de menu for alterada apos o Windows ter creado a janela, esta função deveria ser chamado para desenhar as modificações do menu bar

- __ExtractIcon (shell32.dll):__ A função ExtractIcon recupera o identificador de um ícone exetutado de um arquivo específicado,de uma biblioteca (DLL) ou de uma imagem do tipo ___ico___.

- __SendMessage (user32.dll):__ A função SendMessage envia a mensagem especificada para uma janela(s). A função chama o procedimento da janela para a janela especificada e não retorna até que o procedimento da janela tenha processado a mensagem. A função PostMessage, em contrapartida, posta uma mensagem para uma lista de mensagem em thread e retorna imediatamente.

- __SetLayeredWindowAttributes (user32.dll):__ A função SetLayeredWindowAttributes define a opacidade e cor chave de transparência de uma camada da janela.

- __SetParent (user32.dll):__ A função SetParent altera a janela pai de uma janela filha especificada.

Caso tenham interesse em conhecer mais sobre as API´s do Windows, podem acessar o site [AllAPI.net](http://allapi.mentalis.org/index2.shtml) e acessar a [Lista de Api´s](http://allapi.mentalis.org/apilist/apilist.php).

### Métodos e Propriedades Classe

Com base nas Funções acima, foram criados ___Propriedades___ e ___Métodos___ para esta __Classe__, que irá funcionar como um ___Framework de Projetos VBA___, para realizar alterações na estrutura física do __Formulário (Userform)__. Seguem relação e descrição, de todos os recursos que a __Classe__ oferece. 
Ressaltando que os nomes das ___Propriedades___ e ___Métodos___ são em Inglês, para seguir o padrão da língua estrangeira, utilizada como base para a __Programação VBA__.

#### Propriedades

##### FormStart

Essa propriedade nos permite instanciar o ___Objeto da Classe___, definindo o __Formulário (Userform)__ que será manipulado/personalizado. Definindo essa propriedade, todas as demais já estarão direcionando seus resultados para o __Formulário (Userform)__ definido.

__Exemplo:__
```vb
Option Explicit

Private Sub UserForm_Initialize()
  ' Declaração do objeto da classe.
  Dim objApi As New ApiFunction
  ' Define o Userform para sub-objeto da classe.
  Set objApi.FormStart = UserForm1
End Sub
```

##### ActivateButtons

Essa propriedade permite que o desenvolvedor insira os botões de Maximizar e Minimizar, sejam individualmente ou ambos. Para isso, é necessário a selação de uma das seguintes opções:
- __WS_FULLSIZING:__ ativa, de uma só vez, os dois botões (Maximizar e Minimizar);
- __WS_MAXIMIZE:__ ativa somente o botão de Maximizar. O botão de Minimizar fica visível, mas desabilitado para uso;
- __WS_MINIMIZE:__ ativa somente o botão de Minimizar. O botão de Maximizar fica visível, mas desabilitado para uso.

__Exemplo:__
```vb
Option Explicit

Private Sub UserForm_Initialize()
  ' Declaração do objeto da classe.
  Dim objApi As New ApiFunction
  ' Define o Userform para sub-objeto da classe.
  Set objApi.FormStart = UserForm1
  ' Chama a propriedade que irá ativar os botões Minimizar 
  ' e Maximizar os botões na estrutura do Userform.
  objApi.ActivateButtons = WS_FULLSIZING
End Sub
```

##### IconTitleBarForm

Através desta propriedade, é possível inserir um __Ícone na Barra de Título do Userform__. É preciso apenas informar o caminho onde o arquivo se encontra, para que possa ser passado para as rotinas da propriedade realizarem a transação da imagem para o userform.

__Exemplo:__
```vb
Option Explicit

Private Sub UserForm_Initialize()
  ' Declaração do objeto da classe.
  Dim objApi As New ApiFunction
  ' Define o Userform para sub-objeto da classe.
  Set objApi.FormStart = UserForm1
  ' Passa para a propriedade a localização da imagem
  ' que será inserida na Barra de título do Userform.
  objApi.IconTitleBarForm = ThisWorkbook.Path & "\xyz.ico"
End Sub
```

##### OpacityPercent

Essa propriedade define o __Percentual de Opacidade__, ou seja, a ___Transparência que um Userform___ pode possuir. Geralmente esse é iniciado com __100%__, mas o desenvolvedor pode quere desenvolver um formulário com apenas __60% de Opacidade__. Basta apenas passar o valor da porcentagem desejada.

__Exemplo:__
```vb
Option Explicit

Private Sub UserForm_Initialize()
  ' Declaração do objeto da classe.
  Dim objApi As New ApiFunction
  ' Define o Userform para sub-objeto da classe.
  Set objApi.FormStart = UserForm1
  ' Define o percentual de opacidade para Userform.
  objApi.OpacityPercent = 60
End Sub
```

#### Métodos

##### RemoveTitleBar

Esse método, quando chamado, remove a __Barra de Título do Userform__ definido, quando o __Objeto da Classe__ foi iniciado.

__Exemplo:__
```vb
Option Explicit

Private Sub UserForm_Initialize()
  ' Declaração do objeto da classe.
  Dim objApi As New ApiFunction
  ' Define o Userform para sub-objeto da classe.
  Set objApi.FormStart = UserForm1
  ' Remove a Barra de Título do Userform.
  objApi.RemoveTitleBar
End Sub
```

##### HideCloseButton

Esse método, quando chamado, oculta o __Botão Fechar da Barra de Título do Userform__ definido, quando o __Objeto da Classe__ foi iniciado.

__Exemplo:__
```vb
Option Explicit

Private Sub UserForm_Initialize()
  ' Declaração do objeto da classe.
  Dim objApi As New ApiFunction
  ' Define o Userform para sub-objeto da classe.
  Set objApi.FormStart = UserForm1
  ' Esconde o botão Fechar (X) do Userform.
  objApi.HideCloseButton
End Sub
```

##### ActivateDualButtons

Esse método, quando chamado, ativa os botões de __Maximizar e Minimizar na Barra de Título do Userform__ definido. Executa a mesma função que a ___Propriedade ActiveButtons___, mas não é preciso passar parâmetro para a execução.

__Exemplo:__
```vb
Option Explicit

Private Sub UserForm_Initialize()
  ' Declaração do objeto da classe.
  Dim objApi As New ApiFunction
  ' Define o Userform para sub-objeto da classe.
  Set objApi.FormStart = UserForm1
  ' Ativa os botões de Minimizar e Maximizar do
  ' Userform, como na Propriedade ActivateButtons.
  objApi.ActivateDualButtons
End Sub
```


##### ActivateMaximizeOnly

Esse método, quando chamado, ativa somente o botão de __Maximizar na Barra de Título do Userform__ definido. Executa a mesma função que a ___Propriedade ActiveButtons___, mas não é preciso passar parâmetro para a execução.

__Exemplo:__
```vb
Option Explicit

Private Sub UserForm_Initialize()
  ' Declaração do objeto da classe.
  Dim objApi As New ApiFunction
  ' Define o Userform para sub-objeto da classe.
  Set objApi.FormStart = UserForm1
  ' Ativa somente o botão de Maximizar do
  ' Userform, como na Propriedade ActivateButtons.
  objApi.ActivateMaximizeOnly
End Sub
```


##### ActivateMinimizeOnly

Esse método, quando chamado, ativa somente o botão de __Minimizar na Barra de Título do Userform__ definido. Executa a mesma função que a ___Propriedade ActiveButtons___, mas não é preciso passar parâmetro para a execução.

__Exemplo:__
```vb
Option Explicit

Private Sub UserForm_Initialize()
  ' Declaração do objeto da classe.
  Dim objApi As New ApiFunction
  ' Define o Userform para sub-objeto da classe.
  Set objApi.FormStart = UserForm1
  ' Ativa somente o botão de Minimizar do
  ' Userform, como na Propriedade ActivateButtons.
  objApi.ActivateMinimizeOnly
End Sub
```


##### ShowFormTaskBar

Esse método, quando chamado, extende o acesso ao __Userform para a Barra de Tarefas do Windows__. Desta forma, não se torna obrigatório o acesso a esse userform, unica e exclusivamente pelo __Applicativo Office__.

__Exemplo:__
```vb
Option Explicit

Private Sub UserForm_Initialize()
  ' Declaração do objeto da classe.
  Dim objApi As New ApiFunction
  ' Define o Userform para sub-objeto da classe.
  Set objApi.FormStart = UserForm1
  ' Ativa o Userform na Barra de Tarefas do Window.
  objApi.ShowFormTaskBar
End Sub
```


##### ParentForms

Esse método, quando chamado, define um segundo userform como um __Userform Pai__ do userform atual. Isso significa que o __Userform Filho__ fica limitado a área do __Userform Pai__.

__Exemplo:__
```vb
Option Explicit

Private Sub UserForm_Initialize()
  ' Declaração do objeto da classe.
  Dim objApi As New ApiFunction
  ' Instancia o novo objeto, a partir da classe.
  Set objApi = New ApiFunction
  ' Define relação enre dos Userforms.
  objApi.ParentForms UserForm2.Caption, UserForm1.Caption
End Sub
```
