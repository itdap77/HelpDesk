<html>

<head>

<!--- STYLE SHEETS --->
<link REL="STYLESHEET" TYPE="text/css" HREF="styles.css">

<!--- LIBRARIAS DE FUNCIONES --->
<script language="javascript" src="script.js"></script>

<title>Cambiar Contraseña</title>

</head>

<!--- VARIABLES DEL DOCUMENTO --->

<!--- FUNCIONES CLIENTE --->
<script language="javascript">
</script>

<body style="margin:5px;" link="#000000" vlink="#000000" alink="#000000">


<script language="javascript">
<!--
function Validate(oSesion)
{
	if (oSesion.PasswordNew.value != oSesion.PasswordNew2.value)
	{
		window.alert('La confirmación de la nueva contraseña no es correcta.');
		oSesion.PasswordNew2.focus();
		oSesion.PasswordNew2.select();
		return false;
	}
	return true;
}


//-->

</script>

<%
Dim vBackToURL

If Len(Request.QueryString("BackToURL")) = 0 Then
	vBackToURL = "library/password.asp"
Else
	vBackToURL = Request.QueryString("BackToURL")
End If
%>

  <center>
  <form method="POST" action="../vbslogin.asp" id="Sesion" name="Sesion" IsMainForm class="Form">
    <table border="0" cellspacing="0" cellpadding="0">
      <tr height="18" class="NavBoxHeader2">
	  	<td align="left" valign="top"><img src="../images/borde11.gif" WIDTH="8" HEIGHT="8"></td>
  		<td colspan="2" style="VERTICAL-ALIGN: middle"><b>Cambiar contraseña</b>
  		<input type="hidden" name="BackToURL" value="<%=vBackToURL%>">
  		<input type="hidden" name="Accion" value="CHANGE-PASSWORD"></td>
		<td align="left" valign="top"><img src="../images/borde12.gif" WIDTH="8" HEIGHT="8"></td>
	  </tr>
<%
select case Session("CambiarContrasenia")
case "1"	'Chanse OK
	Response.Write("	  <tr class=""LogInBox"">") & vbCrLf
	Response.Write("		<td>&nbsp;</td>") & vbCrLf
	Response.Write("		<td height=""23"" colspan=""2"" style=""vertical-align:middle;align:center;color:green"">La contraseña se cambio exitosamente") & vbCrLf
	Response.Write("	    </td>") & vbCrLf
	Response.Write("		<td>&nbsp;</td>") & vbCrLf
	Response.Write("      </tr>") & vbCrLf
case "0"	'Error
	Response.Write("	  <tr class=""LogInBox"">") & vbCrLf
	Response.Write("		<td>&nbsp;</td>") & vbCrLf
	Response.Write("		<td height=""23"" colspan=""2"" style=""vertical-align:middle;align:center;color:red""><img src=""../images/alert.gif"" align=""left"" border=""0"" WIDTH=""19"" HEIGHT=""19"">&nbsp;No se pudo cambiar la contraseña") & vbCrLf
	Response.Write("	    </td>") & vbCrLf
	Response.Write("		<td>&nbsp;</td>") & vbCrLf
	Response.Write("      </tr>") & vbCrLf
case else
	Response.Write("	  <tr class=""LogInBox"">") & vbCrLf
	Response.Write("		<td>&nbsp;</td>") & vbCrLf
	Response.Write("		<td colspan=""2"">&nbsp;</td>") & vbCrLf
	Response.Write("		<td>&nbsp;</td>") & vbCrLf
	Response.Write("      </tr>") & vbCrLf
end select
Session("CambiarContrasenia") = ""
%>
	  <tr class="LogInBox">
	    <td>&nbsp;</td>
        <td width="100">Usuario:</td>
	    <td width="100">
          <input type="text" name="Usuario" size="12" style="FONT-FAMILY: Tahoma, Verdana; FONT-SIZE: 10px" readonly value="<%=session("NombreUsuario")%>">
        </td>
		<td>&nbsp;</td>
	  </tr>
      <tr class="LogInBox">
		<td>&nbsp;</td>
        <td>Contraseña actual:</td>
        <td><input type="password" name="Password" size="12" style="FONT-FAMILY: Tahoma, Verdana; FONT-SIZE: 10px">
        </td>
	    <td>&nbsp;</td>
      </tr>
      <tr class="LogInBox">
		<td>&nbsp;</td>
        <td>Contraseña nueva:</td>
        <td><input type="password" name="PasswordNew" size="12" style="FONT-FAMILY: Tahoma, Verdana; FONT-SIZE: 10px">
        </td>
	    <td>&nbsp;</td>
      </tr>
      <tr class="LogInBox">
		<td>&nbsp;</td>
        <td>Confirmación:</td>
        <td><input type="password" name="PasswordNew2" size="12" style="FONT-FAMILY: Tahoma, Verdana; FONT-SIZE: 10px">
        &nbsp;&nbsp;
          <input type="image" src="../images/arrow_rg_gray.gif" align="middle" style="{cursor:hand;}" onClick="return(Validate(Sesion));" id="image2" name="image2" WIDTH="18" HEIGHT="18">
        </td>
	    <td>&nbsp;</td>
      </tr>
      <tr class="LogInBox">
		<td>&nbsp;</td>
        <td colspan="2" style="vertical-align:middle;text-align:center;color:darkblue;"><a href="javascript:window.close();">Cerrar ventana</a></td>
        </td>
	    <td>&nbsp;</td>
      </tr>
	  <tr class="LogInBox">
		<td style="VERTICAL-ALIGN: bottom"><img src="../images/borde21.gif" WIDTH="8" HEIGHT="8"></td>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
		<td style="VERTICAL-ALIGN: bottom"><img src="../images/borde22.gif" WIDTH="8" HEIGHT="8"></td>
	  </tr>
    </table>

<script language="javascript">
<!--

	Sesion.Password.focus();
	
-->
</script>

    
  </form>
  </center>
  
</body>

</html>
