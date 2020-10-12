<%@LANGUAGE="VBSCRIPT"%>
<!--#INCLUDE file="connections/helpdesk.asp"-->
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

set Rs_Ciudad = Server.CreateObject("ADODB.Recordset")
Rs_Ciudad.ActiveConnection = MM_HelpDesk_STRING
Rs_Ciudad.Source = "{call dbo.SPU_Ciudad_VE}"
Rs_Ciudad.CursorType = 0
Rs_Ciudad.CursorLocation = 2
Rs_Ciudad.LockType = 3
Rs_Ciudad.Open()
Rs_Ciudad_numRows = 0

If Request.Form ("pais") <> 0 Then
set Rs_IDCUsuario = Server.CreateObject("ADODB.Recordset")
Rs_IDCUsuario.ActiveConnection = MM_HelpDesk_STRING
Rs_IDCUsuario.Source = "{call dbo.SPU_IDCUsuario_L2(" + Replace(Request.Form("pais"),"'","''") + ")}"
Rs_IDCUsuario.Open()
End if
%>
<html>
<head>
<title>Selecion de Usuario</title>
<script  Languague="JavaScript">
function submitform()
{ if (document.frmseleccionusr.login.value != "") 
	 { window.opener.frmingreso.login.value = frmseleccionusr.login.value;
	   window.opener.frmingreso.idinventario.value = frmseleccionusr.idinventario.value;
	   window.close()
	   }
  else {
		frmseleccionusr.submit();}
		}


</script>
<link rel="stylesheet" href="library/FormArea.css" type="text/css">
<link rel="stylesheet" href="library/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF">

<p align="center"><b></b></p>

<div align="center">
  <center>
    <table width="100%" border="0" class="FormBoxHeader">
      <tr> 
        <td width="5"><b><img src="images/gripgray.gif" width="10" height="13"></b></td>
        <td><font size="2" face="Verdana">Seleccione la Ciudad y su correspondiente 
          Usuario</font></td>
      </tr>
    </table>
    <form name="frmseleccionusr" method="POST">
<%	If Request.Form ("pais") <> 0 Then %>
	<input type="hidden" name="idinventario" value="<%=Rs_IDCUsuario.Fields.Item("idinventario").Value%>">
	<% End if %>
      <table border="0" width="100%" cellspacing="0" cellpadding="0" class="FormBoxBody">
        <tr>
      <td width="100%">
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100%">
                </td>
          </tr>
        </table>
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100%">
                  <table border="0" width="100%" cellspacing="0" cellpadding="0">
                    <tr>
                  <td width="50%">
                        <table border="0" width="100%" cellspacing="0" cellpadding="0" class="FormTable">
                          <tr>
                        <td width="100%"><b><font size="2" face="Verdana">Seleccione
                          la Ciudad:</font></b></td>
                      </tr>
                      <tr>
                        <td width="100%">&nbsp;</td>
                      </tr>
                      <tr>
                        <td width="100%">&nbsp;</td>
                      </tr>
                      <tr>
                        <td width="100%">&nbsp;</td>
                      </tr>
                    </table>
                  </td>
                  <td width="50%">
<select size="4" name="pais" onclick="javascript:submitform();">
                        <%
While (NOT Rs_Ciudad.EOF)
%>
                        <option value="<%=(Rs_Ciudad.Fields.Item("IDCiudad").Value)%>"<%if (CStr(request.form("pais")) = CStr(Rs_Ciudad.Fields.Item("IDCiudad").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(Rs_Ciudad.Fields.Item("Ciudad").Value)%></option>
                        <%
  Rs_Ciudad.MoveNext()
Wend
If (Rs_Ciudad.CursorType > 0) Then
  Rs_Ciudad.MoveFirst
Else
  Rs_Ciudad.Requery
End If
%>

              </select></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100%">&nbsp;</td>
          </tr>
        </table>
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100%">
                  <table border="0" width="100%" cellspacing="0" cellpadding="0">
                    <tr>
                  <td width="44%">
                        <table border="0" width="100%" cellspacing="0" cellpadding="0" class="FormTable">
                          <tr>
                        <td width="100%"><b><font size="2" face="Verdana">Seleccione
                          el Usuario:</font></b></td>
                      </tr>
                      <tr>
                        <td width="100%">&nbsp;</td>
                      </tr>
                      <tr>
                        <td width="100%">&nbsp;</td>
                      </tr>
                      <tr>
                        <td width="100%">&nbsp;</td>
                      </tr>
                    </table>
                  </td>
                  <td width="56%"><select size="4" name="login">
<%
Response.Write (Request.Form("pais") )
If Request.Form("pais")<>0  Then

			While (NOT Rs_IDCUsuario.EOF)
			
		Response.Write (" <option value=""" )
		Response.Write (Rs_IDCUsuario.Fields.Item("login").Value & """> ")
		Response.Write (Rs_IDCUsuario.Fields.Item("login").Value )
		Response.Write ("</option>")
											
			  Rs_IDCUsuario.MoveNext()
			Wend
			If (Rs_IDCUsuario.CursorType > 0) Then
			 Rs_IDCUsuario.MoveFirst
			Else
			 Rs_IDCUsuario.Requery
			End If
End If
%>
              </select></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100%">&nbsp;</td>
          </tr>
        </table>
            <table border="0" width="100%" cellspacing="0" cellpadding="0" class="FormTable">
              <tr>
                <td width="100%"> 
                  <div align="center">
<input type="button" onclick="submitform();" class="FormButton" name="Button" value="Cerrar" >
                  </div>
                </td>
          </tr>
        </table>
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100%">&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>
  </center>
</div>

</body>

</html>
<%
Rs_Ciudad.Close()
If Request.Form ("pais") <> 0 Then
Rs_IDCUsuario.Close()
End if
%>
