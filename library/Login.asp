<script type="text/javascript" >
<!--

function OpenChangePassword()
{
window.open('library/password.asp', '','scrollbars=no,height=150,width=250,status=no,toolbar=no,menubar=no,location=no');

}

function Validate(oSesion)
{/*
	if (isNaN(parseInt(oSesion.Login.value)))
	{
		window.alert('No se indicó una identificación válida.');
		oSesion.Login.focus();
		oSesion.Login.select();
		return false;
	}
	return true;
*/
}

//-->

</script>
<%

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

 %>
<%
Dim vBackToURL

If Len(Request.QueryString("BackToURL")) = 0 Then
	vBackToURL = "main.asp?lang=eng"
Else
	vBackToURL = Request.QueryString("BackToURL")
End If
%>
<%
if request.querystring("IniciarSesion") <> 0 Then
	Session("Login") = request.querystring("iniciarsesion")
else
end if 


%>
  <center>
  <form method="post" action="vbsLogin.asp" id="Sesion" name="Sesion" IsMainForm="true" class="Form">
    <table border="0" Width="180px" cellspacing="0" cellpadding="0" align=center>
      <tr height="18" class="NavBoxHeader2"> 
        <td align="left" valign="top"><img alt="" src="images/borde11.gif" WIDTH="8" height="8" /></td>
        <td colspan="2" style="VERTICAL-ALIGN: middle">
        			<%	If session("IDidioma") = 1 or vDataForm("lang") = "esp" then
													Response.Write "<b>Inicio de sesión</b> " & vbCrLf
									Else
													Response.Write "<b>Start Session</b> " & vbCrLf
									End If	%>
        	
          <input type="hidden" name="BackToURL" value="<%=vBackToURL%>"/>
          <input type="hidden" name="Accion" value="LOGIN"/>
        </td>
        <td align="right" valign="top"><img alt="" src="images/borde12.gif" width="8px" height="8px" /></td>
      </tr>
      
      <tr class="LogInBox"> 
        <td colspan="4">		&nbsp;
        	
        	</td>
      </tr>
      
      <%
    
select case Session("Login")
		case 1 'logeado
			
			Response.Write("<tr class=""logInBox"">") & vbCrLf
			Response.Write("<td colspan=""4"" align=""Center""><b>")
			
			If session("IDidioma") = 1 or vDataForm("lang") = "esp" then
						    Response.Write "<b>Usted a ingresado con exito.</b> " & vbCrLf
						    vBackToURL = "main.asp?lang=esp"	
		    Else
						    Response.Write "<b>You entered sucessfully.</b> " & vbCrLf
						    vBackToURL = "main.asp?lang=eng"	
		    End If	
		    
			Response.Write ("</b><BR><BR></td>") & vbCrLf 	
			Response.Write("</tr>") & vbCrLf		
			Response.Write("<tr class=""logInBox"">")  & vbCrLf
			Response.Write("<td>&nbsp;</td>") & vbCrLf
			Response.Write("<td>&nbsp;</td>") & vbCrLf
			Response.Write("<td>&nbsp;</td>") & vbCrLf
			Response.Write("<td>&nbsp;</td>") & vbCrLf
			'Response.Write("<td colspan=""2"" style=""vertical-align:middle;text-align:center;color:darkblue;"">") & vbCrLf
			
			'If session("IDidioma") = 1 or vDataForm("lang") = "esp" then
			'	Response.Write("<a href=""JavaScript:document.forms['NavSesion'].submit();"">Cerrar sesión</a></td>") & vbCrLf
			'Else
			
          if (Request.ServerVariables("url") <> "/helpdesk/main.asp") Then
                Response.redirect ("main.asp?lang=" & vDataForm("lang"))
          End If	
          'Response.Write("<a href=""JavaScript:document.forms['NavSesion'].submit();"">Close session</a></td>") & vbCrLf
			'End If	
			
		'	Response.Write("<td>&nbsp;</td>") & vbCrLf-->
			Response.Write("</tr>") & vbCrLf
			
		case 2	'Bad UserID
			Response.Write("	  <tr class=""LogInBox"">") & vbCrLf
			Response.Write("		<td>&nbsp;</td>") & vbCrLf


			If session("IDidioma") = 1 or vDataForm("lang") = "esp" then
							Response.Write("		<td height=""23"" colspan=""2"" style=""vertical-align:middle;align:center;color:red""><img src=""images/alert.gif"" align=""left"" border=""0"" WIDTH=""19"" HEIGHT=""19""><b>La contraseña o el usuario ingresado es inválida</b>") & vbCrLf
			Else
							Response.Write("		<td height=""23"" colspan=""2"" style=""vertical-align:middle;align:center;color:red""><img src=""images/alert.gif"" align=""left"" border=""0"" WIDTH=""19"" HEIGHT=""19""><b>Invalid user or password</b>") & vbCrLf
			End If	
						
			Response.Write("	    </td>") & vbCrLf
			Response.Write("		<td>&nbsp;</td>") & vbCrLf
			Response.Write("      </tr>") & vbCrLf
			
			Response.Write "<tr class=""LogInBox"">" & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			
			If session("IDidioma") = 1 or vDataForm("lang") = "esp" then
				Response.Write "<td width=""70"">Usuario:</td>" & vbCrLf
			Else
				Response.Write "<td width=""70"">User:</td>" & vbCrLf
			End If	
			
			Response.Write "<td width=""100"">" & vbCrLf
			Response.Write "<input type=""text"" name=""usuario"" autofocus tabindex=""0"" size=""12"" style=""FONT-FAMILY: Tahoma, Verdana; FONT-SIZE: 10px"">" & vbCrLf
			Response.Write "</td>" & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			Response.Write "</tr>" & vbCrLf
			Response.Write "<tr class=""LogInBox"">"  & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			
			If session("IDidioma") = 1 or vDataForm("lang") = "esp" then
							Response.Write "<td>Contraseña:</td>" & vbCrLf
			Else
							Response.Write "<td>Password:</td>" & vbCrLf
			End If	
			
			Response.Write "<td>" & vbCrLf
			Response.Write "<input type=""password"" name=""Password"" size=""12"" style=""FONT-FAMILY: Tahoma, Verdana; FONT-SIZE: 10px"">" & vbCrLf
			Response.Write "&nbsp;&nbsp;" & vbCrLf
			Response.Write "<input type=""image"" src=""images/arrow_rg_gray.gif"" align=""middle"" style=""{cursor:hand;}"" onClick=""return(Validate(Sesion));"">" & vbCrLf
			Response.Write "</td>" & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			Response.Write "</tr>" & vbCrLf
		      
		case 3	'Bad Password
			Response.Write("	  <tr class=""LogInBox"">") & vbCrLf
			Response.Write("		<td>&nbsp;</td>") & vbCrLf
			
			
			If session("IDidioma") = 1 or vDataForm("lang") = "esp" then
							Response.Write("		<td height=""23"" colspan=""2"" style=""vertical-align:middle;align:center;color:red""><img src=""images/alert.gif"" align=""left"" border=""0"" WIDTH=""19"" HEIGHT=""19""><b>La contraseña o el usuario ingresado es inválida</b>") & vbCrLf
			Else
							Response.Write("		<td height=""23"" colspan=""2"" style=""vertical-align:middle;align:center;color:red""><img src=""images/alert.gif"" align=""left"" border=""0"" WIDTH=""19"" HEIGHT=""19""><b>Invalid username or password</b>") & vbCrLf
			End If	
			Response.Write("	    </td>") & vbCrLf
			Response.Write("		<td>&nbsp;</td>") & vbCrLf
			Response.Write("      </tr>") & vbCrLf
			
			Response.Write "<tr class=""LogInBox"">" & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			
			If session("IDidioma") = 1 or vDataForm("lang") = "esp" then
				Response.Write "<td width=""70"">Usuario:</td>" & vbCrLf
			Else
				Response.Write "<td width=""70"">User:</td>" & vbCrLf
			End If	
			
			Response.Write "<td width=""100"">" & vbCrLf
			Response.Write "<input type=""text"" name=""usuario""  tabindex=""0""  autofocus size=""12"" style=""FONT-FAMILY: Tahoma, Verdana; FONT-SIZE: 10px"">" & vbCrLf
			Response.Write "</td>" & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			Response.Write "</tr>" & vbCrLf
			Response.Write "<tr class=""LogInBox"">"  & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			
			If session("IDidioma") = 1 or vDataForm("lang") = "esp" then
							Response.Write "<td>Contraseña:</td>" & vbCrLf
			Else
							Response.Write "<td>Password:</td>" & vbCrLf
			End If	
			
			Response.Write "<td>" & vbCrLf
			Response.Write "<input type=""password"" name=""Password"" size=""12"" style=""FONT-FAMILY: Tahoma, Verdana; FONT-SIZE: 10px"">" & vbCrLf
			Response.Write "&nbsp;&nbsp;" & vbCrLf
			Response.Write "<input type=""image"" src=""images/arrow_rg_gray.gif"" align=""middle"" style=""{cursor:hand;}"" onClick=""return(Validate(Sesion));"">" & vbCrLf
			Response.Write "</td>" & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			Response.Write "</tr>" & vbCrLf
		
		case else	
			Response.Write("	  <tr class=""LogInBox"">") & vbCrLf
			Response.Write("		<td>&nbsp;</td>") & vbCrLf
			Response.Write("		<td colspan=""2"">&nbsp;</td>") & vbCrLf
			Response.Write("		<td>&nbsp;</td>") & vbCrLf
			Response.Write("      </tr>") & vbCrLf
			
			Response.Write "<tr class=""LogInBox"">" & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			
			If session("IDidioma") = 1 or vDataForm("lang") = "esp" then
				Response.Write "<td width=""70"">Usuario:</td>" & vbCrLf
			Else
				Response.Write "<td width=""70"">User:</td>" & vbCrLf
			End If	
			
			Response.Write "<td width=""100"">" & vbCrLf
			Response.Write "<input type=""text"" name=""usuario""  tabindex=""0"" autofocus size=""12"" style=""FONT-FAMILY: Tahoma, Verdana; FONT-SIZE: 10px"">" & vbCrLf
			Response.Write "</td>" & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			Response.Write "</tr>" & vbCrLf
			Response.Write "<tr class=""LogInBox"">"  & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			
			If session("IDidioma") = 1 or vDataForm("lang") = "esp" then
							Response.Write "<td>Contraseña:</td>" & vbCrLf
			Else
			
							Response.Write "<td>Password:</td>" & vbCrLf
			End If	
			
			Response.Write "<td>" & vbCrLf
			Response.Write "<input type=""password"" name=""Password"" size=""12"" style=""FONT-FAMILY: Tahoma, Verdana; FONT-SIZE: 10px"">" & vbCrLf
			Response.Write "&nbsp;&nbsp;" & vbCrLf
			Response.Write "<input type=""image"" src=""images/arrow_rg_gray.gif"" align=""middle"" style=""{cursor:hand;}"" onClick=""return(Validate(Sesion));"">" & vbCrLf
			Response.Write "</td>" & vbCrLf
			Response.Write "<td>&nbsp;</td>" & vbCrLf
			Response.Write "</tr>" & vbCrLf
end select

      
 

%>
      <tr class="LogInBox"> 
        <td style="VERTICAL-ALIGN: bottom" ><img src="images/borde21.gif" WIDTH="8" HEIGHT="8"></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td style="VERTICAL-ALIGN: bottom" align="right"><img src="images/borde22.gif" WIDTH="8" HEIGHT="8"></td>
      </tr>
    </table>
  </form>
  </center>