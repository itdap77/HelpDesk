<%@ LANGUAGE=VBScript %>
<!--#INCLUDE virtual="/itdap/helpdesk/library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If
%>
<!--- VARIABLES DEL DOCUMENTO --->
<script language="javascript">
{
	document.mtModulo = 'default';
	document.mtHelpSection = 'Inicio';
}



</script>


<table class="TopMenuArea" ID="PageBorder" border="0" cellspacing="0" cellpadding="0">
<tr><td class="Titulo">
   <% Call ReadLang(Rs_Lenguaje,30) %>
</td>
<td>

</td></tr>
</table>

<table class="ContentArea" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr><td>
 <table style="padding:5px; BORDER-BOTTOM: gray 1px solid; font-size: 13px;" border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr>
          <td width="800"> <img src="images/password.gif" align="left" border="0" hspace="10" WIDTH="39" HEIGHT="45">

           <% Call ReadLang(Rs_Lenguaje,26) %>

           </td>
  <td width="270" align="right">
  <!--#INCLUDE file="library/Login.asp"-->
  </td>
 </tr>
 </table>

</td></tr>
</table>

<!--#INCLUDE file="library/PageClose.asp"-->

</body>

</html>