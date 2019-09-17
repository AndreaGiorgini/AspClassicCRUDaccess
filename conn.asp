<%
'****************************************************************************************************************
'Base de dados Principal
   Set objdatabase = Server.CreateObject("ADODB.Connection")
   objdatabase.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("usuario.mdb"))
'***************************************************************************************************************
Response.Buffer         = True
Response.CacheControl   = "private"
Session.LCID            = 1046
Response.Expires        = -54000
Response.Charset        = "UTF-8"
Server.ScriptTimeout    = 86400
'****************************************************************************************************************
%><!--#include file="funcoes.asp"-->