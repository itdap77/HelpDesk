<%@ language=VBScript %> 

<%
response.Write "Session:" & Session("Login")

If (Session("Login") <> 1) Then
	Response.Redirect "default.asp"
End If

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 


%>

<script type="text/vbscript" language="vbscript" runat="server">

Dim DBConn, vIndex

Function GetTipoMsg
	Select case Request.QueryString("TipoMsg")
	Case 0
		GetTipoMsg="Aviso:"
	Case 1
		GetTipoMsg="Atención:"
	Case 2
		GetTipoMsg="Error:"
	Case else
		GetTipoMsg=""
	End select
End Function

</script>

<%


If Request.QueryString("BackToURL").Count <> 0 And Request.QueryString("Msg").Count = 0 And Request.QueryString("CodMsg").Count = 0 Then
	Response.Redirect Request.QueryString("BackToURL")
End If
%>



<!--- COMIENZO DE CONTENIDO --->
<table class="TopMenuArea" ID="PageBorder" border="0" cellspacing="0" cellpadding="0">
<tr><td class="Titulo">
  Mensaje
</td>
   
</tr>
</table>

<table class="ContentArea" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr><td>
    <center>



              <table border=0 class="mensaje">
                <tr height="18">
	  	      <td align="left" valign="top">&nbsp;</td>
  		<td style="VERTICAL-ALIGN: middle">&nbsp;</td>
		      <td align="right" valign="top">&nbsp;</td>
	  </tr>
	  <tr>
	    <td>&nbsp;</td>
        <td class="MsgBox">


			<b><%=GetTipoMsg%></b>
			<BR>
			<%If Request.QueryString("CodMsg").Count <> 0 Then%>
				<%=GetAppMsg(Request.QueryString("CodMsg"))%>
			<%Else%>
				<%=Request.QueryString("Msg")%>
			<%End If%>
			<BR>
			<%	
				If Request.QueryString("BackToURL").Count <> 0 Then
					Response.Write("<a href='" & Request.QueryString("BackToURL") & "'>Continuar</a>")
				Else
					Response.Write("<a href='./'>Continuar</a>")	
				End if
			%>
        
        </td>
		<td>&nbsp;</td>
	  </tr>
	  <tr>
		      <td style="VERTICAL-ALIGN: bottom;TEXT-ALIGN:left;">&nbsp;</td>
	    <td>&nbsp;</td>
		      <td style="VERTICAL-ALIGN: bottom;TEXT-ALIGN:right;">&nbsp;</td>
	  </tr>
    </table>

	</center>
</td></tr>
</table>


<!--#INCLUDE file="library/PageClose.asp"-->

