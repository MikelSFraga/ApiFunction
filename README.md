# ApiFunction Class V1.0.0
	
Com o intuito de facilitar o desenvolvimento de ferramentas, utilizando o processo de programação em __VBA do pacote Microsoft Office__, iniciou-se o desenvolvimeno desta __Classe__.

Seu objetivo é a unificação de várias rotinas que utilizam as __Funções dos Windows (API´s)__, através das quais é possível realizar alterações na estrutura dos __Formulários (Userform´s)__ que utilizamos na estrutura do __Visual Basic for Application (VBA)__.

Esta __Classe__ atenderá as diferentes _Arquiteturas do Sistema Operacional Windows (32 bits e 64 bits)_, como também as _Arquiteturas do Pacote Microsoft Office (VBA6 e VBA7)_.

### Funções Windows (API´s)

Nesta versão atual do __ApiFunction (v1.0.0)__, são utilizadas as seguintes __Funções/Api´s do Windows_:

- __FindWindow (user32.dll):__ The FindWindow function retrieves the handle to the top-level window whose class name and window name match the specified strings. This function does not search child windows.

- __GetWindowLong (user32.dll):__ The GetWindowLong function retrieves information about the specified window. The function also retrieves the 32-bit (long) value at the specified offset into the extra window memory of a window.

- __SetWindowLong (user32.dll):__ The SetWindowLong function changes an attribute of the specified window. The function also sets a 32-bit (long) value at the specified offset into the extra window memory of a window.

- __ShowWindow (user32.dll):__ The ShowWindow function sets the specified window’s show state.

- __SetFocus (user32.dll):__ The SetFocus function sets the keyboard focus to the specified window. The window must be associated with the calling thread’s message queue.

- __DrawMenuBar (user32.dll):__ The DrawMenuBar function redraws the menu bar of the specified window. If the menu bar changes after Windows has created the window, this function must be called to draw the changed menu bar.

- __ExtractIcon (shell32.dll):__ The ExtractIcon function retrieves the handle of an icon from the specified executable file, dynamic-link library (DLL), or icon file.

- __SendMessage (user32.dll):__ The SendMessage function sends the specified message to a window or windows. The function calls the window procedure for the specified window and does not return until the window procedure has processed the message. The PostMessage function, in contrast, posts a message to a thread’s message queue and returns immediately.

- __SetLayeredWindowAttributes (user32.dll):__ The SetLayeredWindowAttributes function sets the opacity and transparency color key of a layered window.

- __SetParent (user32.dll):__ The SetParent function changes the parent window of the specified child window.

Caso tenham interesse em conhecer mais sobre as API´s do Windows, podem acessar o site [AllAPI.net](http://allapi.mentalis.org/index2.shtml) e acessar a [Lista de Api´s](http://allapi.mentalis.org/apilist/apilist.php).

### Métodos e Propriedades Classe

Com base nas Funções acima, foram criados ___Propriedades___ e ___Métodos___ para esta __Classe__, que irá funcionar como um ___Framework de Projetos VBA___, para realizar alterações na estrutura física do __Formulário (Userform)__. Seguem relação e descrição, de todos os recursos que a __Classe__ oferece.

#### Propriedades

##### FormStart


```vb

```

##### ActivateButtons


```vb

```

##### IconTitleBarForm


```vb

```

##### OpacityPercent


```vb

```

#### Métodos

##### RemoveTitleBar


```vb

```

##### HideCloseButton


```vb

```

##### ActivateDualButtons


```vb

```


##### ActivateMaximizeOnly


```vb

```


##### ActivateMinimizeOnly


```vb

```


##### ShowFormTaskBar


```vb

```


##### ParentForms


```vb

```




