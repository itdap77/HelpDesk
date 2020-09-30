<%@ LANGUAGE=VBScript  %>
<!--#INCLUDE file="./library/Funciones.asp"-->
<%

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

If (Session("Login") <> 1) Then
	If vDataForm("lang") = "esp" Then
		Response.Redirect ("mensaje.asp?Msg=Usted no esta logueado en el sistema.&backtourl=default.asp")
	Else
		Response.Redirect ("mensaje.asp?Msg=You are not logged to the system.&backtourl=default.asp")
	End If
End If 

Set Rs_Lenguaje = Server.CreateObject("ADODB.Recordset")
Rs_Lenguaje.ActiveConnection = MM_HelpDesk_STRING
Rs_Lenguaje.Source = "{call dbo.SPU_Lenguaje_V(" + cstr(Session("IDidioma")) + ",'" + GetPageName() + "')}"
Rs_Lenguaje.Open()

%>
<html>
<head>
<title><% Call ReadLang(Rs_Lenguaje,69) %> - ITDap Worldwide Solutions</title>
<link rel="stylesheet" href="library/styles.css" type="text/css">
<link rel="stylesheet" href="library/FormArea.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="FormBoxHeader">
  <tr> 
    <td width="5" ><img alt="" src="images/gripgray.gif" width="10" height="13"/></td>
    <td><% Call ReadLang(Rs_Lenguaje,69) %></td>
    <td width="10" ><a href="" onclick="Javascript:window.close()";><img alt="" src="images/cerrar.gif" width="16" height="16" border="0"/></a></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="FormBoxBody">
  <tr>
    <td>
      <table width="100%" height="115" border="0" cellpadding="0" cellspacing="0" class="Texto8">
        <tr>
        <td><%=vDataForm ("Observaciones")%></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><div align="center">
    </div>
	</td>
  </tr>
</table>
<div align="center"> </div>
</body>
</html>
