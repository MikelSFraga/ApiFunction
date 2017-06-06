<!DOCTYPE html>
<html>
<head>
	<title></title>
</head>
<body>
<h1>ApiFunction Class V1.0.0</h1>

Com o intuito de facilitar o desenvolvimento de ferramentas, utilizando o processo de programação em VBA do pacote Microsoft Office, iniciou-se o trabalho de criação desta classe.

Seu objetivo é unificar as rotinas, que utiliza as API´s dos Windows, que tornam possível realizar alterações na estrutura dos Formulários (Userform), que são criados e manipulados pelo Visual Basic for Application (VBA). 

Esta classe também atenderá as diferentes Arquiteturas do Sistema Operacional Windows (32 bits e 64 bits), como também as Arquiteturas do Pacote Microsoft Office (VBA6 e VBA7). e de  Criação de uma Classe em VBA em múltipla Arquitetura (32 bits e 64 bits).

Na atual versão do ApiFunction (v1.0.0), as funções já fazem uso das seguintes rotinas:

- FindWindow			(user32.dll)
	- The FindWindow function retrieves the handle to the top-level window whose class name and window name match the specified strings. This function does not search child windows.

- GetWindowLong			(user32.dll)
	- The GetWindowLong function retrieves information about the specified window. The function also retrieves the 32-bit (long) value at the specified offset into the extra window memory of a window.

 - SetWindowLong		(user32.dll)
 	- The SetWindowLong function changes an attribute of the specified window. The function also sets a 32-bit (long) value at the specified offset into the extra window memory of a window.

 - ShowWindow			(user32.dll)
 	- The ShowWindow function sets the specified window’s show state.

 - SetFocus				(user32.dll)
 	- The SetFocus function sets the keyboard focus to the specified window. The window must be associated with the calling thread’s message queue.
 - DrawMenuBar			(user32.dll)
    - The DrawMenuBar function redraws the menu bar of the specified window. If the menu bar changes after Windows has created the window, this function must be called to draw the changed menu bar.

 Através das APIs acima mencionadas, foi possível desenvolver algumas funções, que irão funcionar como um framework, para realizar alterações na estrutura física do Formulário (Userform).
 
</body>
</html>



