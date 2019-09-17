<!--#include file="conn.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title>::<%=sitenome%></title>
		<link href="css/ag.css" rel="styleSheet" type="text/css">
		<% Response.Charset="ISO-8859-1" %>
	</head>
	<body style="margin: 10 10 10 10; background-color: #FFFFFF;">
		<table width="1000px" height='20px' border="0" align="center" cellpadding="0" cellspacing="0">
			<tr><td style='border-bottom: solid 1px #ccc;'><b>Usuários</b></td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>
<%
	dim Usuario, data
	Usuario = Session("idUsuario")
	data    = date()
	'//====================================================================================================================
		if isempty(request.querystring("type")) then 
			formulario 
		end if
	'//====================================================================================================================
		if request.querystring("type") = "alt" or request.querystring("type") = "exc" then
			set objtable = objdatabase.execute("select nmUsuario, idUsuario, senhaUsuario, perfilUsuario, ativoUsuario, deptoUsuario, emailUsuario, codigoUsuario from Usuario where idUsuario = "&request.querystring("cod")&"")
			nmUsuario     = objtable("nmUsuario")
			id            = objtable("idUsuario")
			senhaUsuario  =  AsciiToString(objtable("senhaUsuario"))
			perfilUsuario = objtable("perfilUsuario")
			ativoUsuario  = objtable("ativoUsuario")
			deptoUsuario  = objtable("deptoUsuario")
			emailUsuario  = objtable("emailUsuario")
			codigoUsuario = objtable("codigoUsuario")
			if ativoUsuario = true  then ativoUsuario=1 end if
			if ativoUsuario = false then ativoUsuario=0 end if
			formulario
		end if
	'//====================================================================================================================
		if request("botao") = "incluir" then
			set objtable = objdatabase.execute("select * from Usuario where nmUsuario = '"&request("nmUsuario")&"';")
			if objtable.bof or objtable.eof then
				senhaini   = StringToAscii(request("senhaUsuario"))
				nrsenha    = len(senhaini)
				if nrsenha = 3  then senha2 = senhaini&"032032032032" end if
				if nrsenha = 6  then senha2 = senhaini&"032032032"    end if
				if nrsenha = 9  then senha2 = senhaini&"032032"       end if
				if nrsenha = 12 then senha2 = senhaini&"032"          end if
				if nrsenha = 15 then senha2 = senhaini                end if	
				senhaUsuario = senha2
				set objtable = objdatabase.execute("INSERT INTO Usuario (UsuarioInclusao, dtInclusao, UsuarioUltimaAlteracao, dtUltimaAlteracao, nmUsuario, senhaUsuario, ativoUsuario, perfilUsuario, deptoUsuario, emailUsuario, codigoUsuario) VALUES ('"&Usuario&"', '"&data&"', '"&Usuario&"', '"&data&"', '"&request("nmUsuario")&"', '"&senhaUsuario&"', '"&request("ativoUsuario")&"', '"&request("perfilUsuario")&"', '"&request("deptoUsuario")&"', '"&request("emailUsuario")&"', '"&request("codigoUsuario")&"')")
			else %><script language="javaScript" type="text/javaScript">alert('Já existe um dado cadastrado com este descrição.');</script><%
			end if
		end if
	'//====================================================================================================================
		if request("botao") = "alterar" then
			senhaini = StringToAscii(request("senhaUsuario"))
			nrsenha   = len(senhaini)
			if nrsenha = 3  then senha2 = senhaini&"032032032032" end if
			if nrsenha = 6  then senha2 = senhaini&"032032032"    end if
			if nrsenha = 9  then senha2 = senhaini&"032032"       end if
			if nrsenha = 12 then senha2 = senhaini&"032"          end if
			if nrsenha = 15 then senha2 = senhaini                end if	
			senhaUsuario = senha2			
			set objtable = objdatabase.execute("update Usuario set UsuarioUltimaAlteracao= '"&Usuario&"', dtUltimaAlteracao = '"&data&"', nmUsuario = '"&request("nmUsuario")&"', senhaUsuario = '"&senhaUsuario&"', ativoUsuario = '"&request("ativoUsuario")&"', perfilUsuario = '"&request("perfilUsuario")&"', deptoUsuario = '"&request("deptoUsuario")&"', emailUsuario = '"&request("emailUsuario")&"', codigoUsuario = '"&request("codigoUsuario")&"' where idUsuario = "&request("id")&";")
		end if
	'//====================================================================================================================
		if (request("botao") = "excluir" and trim(request("id")) = "") then
			%><script language="javaScript" type="text/javaScript">alert('Impossível deletar um campo em branco!');document.location="Usuario.asp";</script><%
		end if
	'//====================================================================================================================
		if request("botao") = "excluir" then
			if  trim(request("id") = "") then
			%><script language="javaScript" type="text/javaScript">alert('Impossível deletar um campo em branco!');document.location="Usuario.asp";</script><%
			else
				set objtable = objdatabase.execute("delete from Usuario where idUsuario = "&request("id")&"")
			end if
		end if
	'//====================================================================================================================
	sub formulario %>
		<form name='Usuario' action="Usuario.asp">
			<table border='0' cellspacing='0' width='600px' style="margin-left:20;" height='20' border="0" cellpadding="0" class="font">
				<tr><td colspan='3' style='border-bottom: solid 1px #ccc;width:600px'><b>Cadastro de Usuário</b></td></tr>
				<tr><td colspan='3' width='5px'>&nbsp;</td></tr>
				<tr>
					<td width='25%' align='RIGHT' >Digite o login:&nbsp;</td>
					<td width='40%' align='left'><input type="text" name="nmUsuario" value="<%=nmUsuario%>" class="input" SIZE='15' maxlength='15'/></td>
					<td width='35%' align='LEFT'  >&nbsp;</td>
				</tr>
				<tr>
					<td width='25%' align='RIGHT' >Digite a senha:&nbsp;</td>
					<td width='40%' align='left'><input type="text" name="senhaUsuario" value="<%=senhaUsuario%>" class="input" SIZE='10' maxlength='5'/></td>
					<td width='35%' align='LEFT'  >&nbsp;</td>
				</tr>
				<tr>
					<td width='25%' align='RIGHT' >Digite o Setor:&nbsp;</td>
					<td width='40%' align='left'><input type="text" name="deptoUsuario" value="<%=deptoUsuario%>" class="input" SIZE='50' maxlength='50'/></td>
					<td width='35%' align='LEFT'  >&nbsp;</td>
				</tr>
				<tr>
					<td width='25%' align='RIGHT' >Digite o Código:&nbsp;</td>
					<td width='40%' align='left'><input type="text" name="codigoUsuario" value="<%=codigoUsuario%>" class="input" SIZE='50' maxlength='50'/></td>
					<td width='35%' align='LEFT'  >&nbsp;</td>
				</tr>
				<tr>
					<td width='25%' align='RIGHT' >Digite o E-mail:&nbsp;</td>
					<td width='40%' align='left'><input type="text" name="emailUsuario" value="<%=emailUsuario%>" class="input" SIZE='50' maxlength='50'/></td>
					<td width='35%' align='LEFT'  >&nbsp;</td>
				</tr>

				<tr>
					<td width='25%' align='RIGHT' >Ativo?&nbsp;</td>
					<td width='40%' align='left'>
						<input type='radio' name="ativoUsuario" value="1" <%if ativoUsuario=1 then response.write "Checked='checked'" end if%>/>Sim 
					    <input type='radio' name="ativoUsuario" value="0" <%if ativoUsuario=0 then response.write "Checked='checked'" end if%>/>Não </td>
					<td width='35%' align='LEFT'  >&nbsp;</td>
				</tr>
				<tr>
					<td width='25%' align='RIGHT' valign='top'>Perfil:&nbsp;</td>
					<td colspan="2" align='left'>
					         <input type='radio' name="PerfilUsuario" value="B" <%if perfilUsuario="B" or perfilUsuario = "" then response.write "Checked='checked'" end if%>/>Básico 
						<br/><input type='radio' name="PerfilUsuario" value="A" <%if perfilUsuario="A" then response.write "Checked='checked'" end if%>/>Administrador
					</td>
				</tr>
				<tr><td colspan='3'>&nbsp;</td></tr>
				<tr>
					<td width='25%' align='RIGHT' >&nbsp;</td>
					<td width='40%' align='CENTER'>&nbsp;</td>
					<td width='35%' align='LEFT'  >&nbsp;
						<input type="hidden" name="id" value="<%=id%>">
						<input type="submit" name="botao" id="botao" value="incluir" class="botao"/>
						<input type="submit" name="botao" id="botao" value="alterar" class="botao"/>
						<input type="submit" name="botao" id="botao" value="excluir" class="botao"/>
					</td>
				</tr>
			</table>
		</form>
	</body>
<% 	end sub%>	
<% lista%>
<% sub lista%>
	<table border='0' cellspacing='0' width='600px' style="margin-left:20;" height='20' border="0" cellpadding="0" class="font">
		<tr><td colspan='3' style="width:600px;border-bottom: solid 1px #CCC;width:600px;'><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Lista de dados cadastrados:</font></strong></td></tr>
		<tr><td>&nbsp;</td></tr>		
<%
set objtable = objdatabase.execute("select * from Usuario order by nmUsuario asc")
if objtable.eof or objtable.bof then
%>
		<tr><td  colspan='3' ><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Nenhum dado cadastrado.</font></strong></td></tr>
	</table>
<%
else
%>
		<tr> 
			<td><b>Usuário</b></td>
			<td>&nbsp;</td>
			<td><b>Alterar - Excluir</b></td>
		</tr>
		<TR><TD colspan="3">&nbsp;</TD></TR>
		<%		do while not objtable.eof %>
		<TR> 
			<TD style='border-bottom: solid 1px #ccc;width:600px'><%=objtable("nmUsuario")%></TD>
			<TD style='border-bottom: solid 1px #ccc;width:600px'>&nbsp;</TD>
			<TD style='border-bottom: solid 1px #ccc;width:600px'><a href="Usuario.asp?cod=<%=objtable("idUsuario")%>&type=alt">Alterar</a> - <a href="Usuario.asp?cod=<%=objtable("idUsuario")%>&type=exc">Excluir</a></TD>
		</TR>
	<%			objtable.movenext
			loop %>
	</table>
<%
end if
objtable.close
Set objtable = nothing
end sub
%>
			</td>
		</tr>
		<tr><td>&nbsp;</td></tr>
		</table>
	</body>
</html>