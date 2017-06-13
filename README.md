<html>
<head>
	
</head>
<body>
	<h1>ApiFunction Class V1.0.0</h1>

	<hr/>
	
	<p>Com o intuito de facilitar o desenvolvimento de ferramentas, utilizando o processo de programação em VBA do pacote Microsoft Office, iniciou-se o desenvolvimeno desta classe.</p>

	<p>Seu objetivo é a unificação de várias rotinas que utilizam as Funções dos Windows (API´s), através das quais é possível realizar alterações na estrutura dos Formulários (Userform´s) que utilizamos na estrutura do Visual Basic for Application (VBA).</p>

	<p>Esta classe atenderá as diferentes Arquiteturas do Sistema Operacional Windows (32 bits e 64 bits), como também as Arquiteturas do Pacote Microsoft Office (VBA6 e VBA7).</p>

	<p>Nesta versão atual do ApiFunction (v1.0.0), são utilizadas as seguintes Funções/Api´s do Windows:</p>

	<ol><li>FindWindow			(user32.dll)</li>
	<ol><li>The FindWindow function retrieves the handle to the top-level window whose class name and window name match the specified strings. This function does not search child windows.</li></ol></ol>

	<ol><li>GetWindowLong			(user32.dll)</li>
	<ol><li>The GetWindowLong function retrieves information about the specified window. The function also retrieves the 32-bit (long) value at the specified offset into the extra window memory of a window.</li></ol></ol>


</body>
</html>

 - SetWindowLong		(user32.dll)
 	- The SetWindowLong function changes an attribute of the specified window. The function also sets a 32-bit (long) value at the specified offset into the extra window memory of a window.

 - ShowWindow			(user32.dll)
 	- The ShowWindow function sets the specified window’s show state.

 - SetFocus				(user32.dll)
 	- The SetFocus function sets the keyboard focus to the specified window. The window must be associated with the calling thread’s message queue.


 Através das APIs acima mencionadas, foi possível desenvolver algumas funções, que irão funcionar como um framework, para realizar alterações na estrutura física do Formulário (Userform).