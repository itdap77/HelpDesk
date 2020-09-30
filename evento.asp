<%@ LANGUAGE=VBScript %> 

<!--#include file="library/Funciones.asp" -->

<%

If Not Len(request.Form) = 0 Then 
	Set vDataForm = Request.Form 
Else 
	Set vDataForm = Request.QueryString 
End If 

ValidSession()

set rs_estado = Server.CreateObject("ADODB.Recordset")
rs_estado.ActiveConnection = MM_HelpDesk_STRING
rs_estado.Source = "{call dbo.SPU_Estado_V}"
rs_estado.Open()


If vDataForm("IDEvento") <> 0 Then
	set Rs_Evento = Server.CreateObject("ADODB.Recordset")
	Rs_Evento.ActiveConnection = MM_HelpDesk_STRING
	Rs_Evento.Source = "{call dbo.SPU_Evento_V (" + vDataForm ("IDEvento") + ")}"
	Rs_Evento.Open()
	
End if

If vDataForm("IDTicket") = "" Then
	If vDataForm("lang") = "esp" Then
		Response.Redirect ("mensaje.asp?Msg=No se paso el ID de Ticket&backtourl=default.asp")
	Else
		Response.Redirect ("mensaje.asp?Msg=Ticket Id was not sent.&backtourl=default.asp")
	End If
End if
%>

<script type="text/javascript" >

function submitform() {
    
var validresult;
validresult = MM_validateForm('Descripcion', '', 'R');

if (validresult == 0) {    
                                        frmingreso.action = 'vbsevento.asp';
                                        frmingreso.submit();
                                        window.opener.location.reload();
                                }


}
</script>

<script type="text/javascript" src="library/script.js"></script>

<link rel="stylesheet" href="library/styles.css" type="text/css"/>
<link rel="stylesheet" href="library/FormArea.css" type="text/css"/>
<link rel="stylesheet" href="library/navarea.css" type="text/css"/>


<form name="frmingreso" method="post" action="">
  <input type="hidden" name="IdTicket" value="<%=vDataForm ("IdTicket") %>"/>
  <input type="hidden" name="Accion" value="<%=vDataForm("Accion") %>"/>
  <input type="hidden" name="IdEvento" value="<%=vDataForm ("IDEvento")%>"/>
  <input type="hidden" name="login" value="<%=Session("NombreUsuario")%>"/>
  
  <table width="400" border="0">
    <tr> 
            <td> 
     
        <table width="400" border="0" class="FormBoxHeader" cellpadding="0" cellspacing="0">
          <tr class="RowFormBoxBody"> 
            <td width="5"><img src="images/gripgray.gif" alt="" width="10" height="13" /></td>
            <td><%=vDataForm("title")%></td>
            <td width="30"> 
              <div align="center"><a href="#" onclick="JavaScript: submitform();"><img src="images/guardar.gif" alt="" width="16" height="16"  border="0"/></a></div>
            </td>
            <td align="right" width="30"> 
              <div align="center"><a href="#" onclick="JavaScript: window.close();"><img src="images/cerrar.gif" alt="" width="16" height="16" border="0"/></a></div>
            </td>
          </tr>
        </table>
                 <% 
                    Select Case vDataForm("Accion") 
							  	Case "S"    'Solucionar
							  			 	Response.Write ("<input type=""hidden"" name=""IDEstado"" value=""3""/>")
							  	Case "R"    'Reclamar
							  			 	Response.Write ("<input type=""hidden"" name=""IDEstado"" value=""5""/>") 
							  	Case "C"   'Cerrar
							  			 	Response.Write ("<input type=""hidden"" name=""IDEstado"" value=""4""/>")
							  	Case "D"   'Derivar
							  			 	Response.Write ("<input type=""hidden"" name=""IDEstado"" value=""2""/>")
							Case "I"    'check in
							  			 	Response.Write ("<input type=""hidden"" name=""IDEstado"" value=""6""/>")
							Case "O"    'check out
							  			 	Response.Write ("<input type=""hidden"" name=""IDEstado"" value=""7""/>")
							Case "H"    'on hold
							  			 	Response.Write ("<input type=""hidden"" name=""IDEstado"" value=""14""/>")
							  			 		 
					  End Select
                       'titulo para pasar la accion por mail
                      Response.Write ("<input type=""hidden"" name=""title"" value='")
                      Response.Write (vDataForm("title"))
                      Response.Write ("'/>")
                       
                      Response.Write ("<input type=""hidden"" name=""Problema"" value='")
                      Response.Write (vDataForm("Problema"))
                      Response.Write ("'/>")

								%>  
        <table width="100%" border="0" class="FormBoxBody">
            <tr> 
            
            	<% If vDataForm("Accion") = "D" Then
    				            set rs_uderiv = Server.CreateObject("ADODB.Recordset")
								rs_uderiv.ActiveConnection = MM_HelpDesk_STRING
								rs_uderiv.Source = "{call dbo.SPU_Usuarios_V_Deriv}"
								rs_uderiv.Open()
								
								
								%>
								<td><b>Asigned to</b></td>
                                <td > 
								 <select name="IDUsuarioDerivado" size="1" class="FormTable" >
			                      
			                      <% While (NOT rs_uderiv.EOF) %>
                                      <option value="<%=(rs_uderiv.Fields.Item("IDUsuario").Value)%>" ><%=(rs_uderiv.Fields.Item("login").Value)%></option>
                                      <%
							                  rs_uderiv.MoveNext()
							            Wend
							            %>
                                   </select>
                    <%
										
            		Else
                            Response.Write ("<td>&nbsp;</b></td><td >")
            				Response.Write ("<input type=""hidden"" name=""IDUsuarioDerivado"" value=""")
            				Response.Write (vDataForm("IDUsuarioDeriv"))
            				Response.Write (""">")
            				
            				
            		 End If
            	%>
            
               
            </td>
          </tr>
          <tr> 
            <td> <b>Event description</b></td>
            <td> 
              <%
							  If vDataForm("IDEvento") <> 0 Then
					
										Response.Write ("<textarea name=""Descripcion"" class=""FormTable"" cols=""50"" rows=""3"" value=""")
										Response.Write (Rs_evento.Fields.Item("Descripcion").value)
										Response.Write ("""></textarea>")
				              Else              
						      	        Response.Write ("<textarea name=""Descripcion"" class=""FormTable"" cols=""50"" rows=""3"" value="""">" & vDataForm("title") & "</textarea>")
							  End if
              %>
            </td>
          </tr>
          <tr> 
            <td> <b>Notes</b></td>
            <td rowspan="2"> 
              <%
							  If vDataForm("IDEvento") <> 0 Then
										Response.Write ("<textarea name=""Observaciones"" rows=""3"" cols=""50"" class=""FormTable"">")
										Response.Write (Rs_evento.Fields.Item("Observaciones").value)
										Response.Write ("</textarea>")
				              Else              
						                Response.Write ("<textarea name=""Observaciones"" rows=""3"" cols=""50"" class=""FormTable""></textarea>")
							  End if
              %>
            </td>
          </tr>
          <tr> 
            <td> 
              &nbsp;
            </td>
          </tr>
        </table>
            </td>
          </tr>
        </table>
      </form>

<%
rs_estado.Close()
%>