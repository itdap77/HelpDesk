<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
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

set RsTipoProblema = Server.CreateObject("ADODB.Recordset")
RsTipoProblema.ActiveConnection = MM_HelpDesk_STRING
RsTipoProblema.Source = "{call dbo.SPU_Problema_Cat1}"
RsTipoProblema.Open()

If  Request.Form ("Categoria1") <> 0 Then
	set RsCategoria1 = Server.CreateObject("ADODB.Recordset")
	RsCategoria1.ActiveConnection = MM_HelpDesk_STRING
	RsCategoria1.Source = "{call dbo.SPU_Problema_Cat2(" + Replace(Request.Form("categoria1"), "'", "''") + ")}"
	RsCategoria1.Open()
End if

If  Request.Form ("Categoria2") <> 0 Then
	set RsCategoria2 = Server.CreateObject("ADODB.Recordset")
	RsCategoria2.ActiveConnection = MM_HelpDesk_STRING
	RsCategoria2.Source = "{call dbo.SPU_Problema_Cat3(" + Replace(Request.Form("categoria2"), "'", "''") + ")}"
	RsCategoria2.Open()
End if

if len(request.Form("idpc")) <> 0 then
	set Rs_PC = Server.CreateObject("ADODB.Recordset")
	Rs_PC.ActiveConnection = MM_HelpDesk_STRING
	Rs_PC.Source = "{call dbo.SPU_IDUsuarioXLogin_L('" & Request.Form("login") & "')}"
	Rs_PC.Open()
else
	set Rs_PC = Server.CreateObject("ADODB.Recordset")
	Rs_PC.ActiveConnection = MM_HelpDesk_STRING
	Rs_PC.Source = "{call dbo.SPU_IDUsuarioXLogin_L('" & session("nombreusuario") & "')}"
	Rs_PC.Open()
end if

%>

<script type="text/vbscript" language="VBScript" RUNAT="Server">
dim oRs, vDataForm, ST

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 
</script>

<script language="JavaScript" type="text/javascript">

//document.Modulo = 'frmingreso.asp';
//document.HelpSection = 'Ingreso';

function seldat(URL,DATOS) 
{
window.open(URL,"_blank",DATOS);
}

function reload()
{
frmingreso.submit();
}

function querySt(ji) {
	hu = window.location.search.substring(1);
	gy = hu.split("&");
	for (i=0;i<gy.length;i++) {
	ft = gy[i].split("=");
	if (ft[0] == ji) {
	return ft[1];
}
}
}

function submitfrm()
{
var language = querySt("lang");

if (frmingreso.categoria1.value == 0)
{ 
	if (language=='esp'){
		alert ('El campo de tipo de problema es obligatorio.!!');
	}else {alert ('Type of problem field is required.!!');}
}
else { if (frmingreso.categoria2.value == 0)
		{ if (language=='esp'){
		alert ('El campo de categoria es obligatorio.!!');
	}else {alert ('Category field is required.!!');}
			
	
		}
		else {if (frmingreso.categoria3.value == 0)
				{if (language=='esp'){
					alert ('El campo de detalle es obligatorio.!!');
				}else {alert ('Detail field is required.!!');}
				}		
			else {	
				if (frmingreso.observaciones.value == 0) 
					{if (language=='esp'){
					alert ('El campo observaciones es obligatorio.!!');
				}else {alert ('Description field is required.!!');}
					}
						
				else {
					frmingreso.action = "vbsingresotickets.asp";
					frmingreso.submit();	
					}
				}
			}
		}
}
</script>
    
      
          <table width="100%" border="0" class="TopMenuArea" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo"><% Call ReadLang(Rs_Lenguaje,1) %></td>
            </tr>
			</table>
      <form name="frmingreso" method="post" action="">
          <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" class="ContentArea">
            <tr>
              <td>
                <table width="100%" border="0"  class="FormBoxHeader">
                  <tr>
                    <td width="5"><img alt="" src="images/gripgray.gif" width="10" height="13"/></td>
                    <td><% Call ReadLang(Rs_Lenguaje,2) %></td>
                  </tr>
                </table>
	       
	       				<input name="idinventario" id="idinventario" value="NULL" type="hidden" />
		                <%
						 If (len(Request.Form("login")) <> 0) Then 
						    Response.Write (" <input type=""hidden"" name=""login"" id=""login"" class=""formtable"" readonly value="""  )
							Response.Write (Request.Form("login"))
							Response.Write (""" class=""FormTable"" onblur=""MM_validateForm('login','','R')""> ")
						 Else
							Response.Write ("<input type=""hidden"" name=""login"" id=""login"" class=""formtable""  readonly value=""" )
							Response.Write  Session("NombreUsuario")
							Response.Write (""" class=""FormTable"" onblur=""MM_validateForm('login','','R')""> ")
						End If
					%>
							  <div align="center"> 
                                <%	
                                If  ((Request.Form("otrouser"))="1") Then 
								Response.Write ("<input type=""hidden"" value=""1"" name=""otrouser"" readonly onclick=""javascript:seldat('listauser.asp','width=495, height=290, scrollbars=no');"" checked >")
									Response.Write ("<input type=""hidden"" value=""0"" name=""otrouser"" readonly> ")
								Else
									Response.Write ("<input type=""hidden"" value=""1"" name=""otrouser"" readonly onclick=""javascript:seldat('listauser.asp','width=495, height=290, scrollbars=no');""  >")
									Response.Write ("<input type=""hidden"" value=""0"" name=""otrouser"" checked readonly> ")
								End If
					
								%>
                 </div> 
                         
              <table border="0" width="100%" cellspacing="0" cellpadding="0" class="FormBoxBody" height="300">
                 <tr> 
                   <td colspan="3" > 
                      <%
						 If (len(Request.Form("idpc")) = 0) Then 
						    Response.Write (" <input type=""hidden"" name=""idpc"" id=""idpc"" class=""formtable"" readonly value="""  )
							Response.Write (Rs_Pc.fields.item("idpc").value)
							Response.Write (""" class=""FormTable"" onblur=""MM_validateForm('login','','R')""> ")
						 Else
							Response.Write ("<input type=""hidden"" name=""idpc"" id=""idpc"" class=""formtable""  readonly value=""" )
							Response.Write  Request.Form("idpc") 'Session("idpc")
							Response.Write (""" class=""FormTable"" onblur=""MM_validateForm('login','','R')""> ")
						End If
					%>
                    <!--a href="#" onClick="javascript:seldat('listapc.asp','width=495, height=290, scrollbars=no');">Seleccione 
                    la PC.</a>*/ --> 
                                <input type="hidden" name="lang" value='<%=vDataForm("lang") %>' />    
                  </td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td >&nbsp;</td>
                  <td>&nbsp; </td>
                  <td >&nbsp;</td>
                </tr>
                <tr> 
                  <td>&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,3) %></td>
                  <td > 
                    
                    <select name="categoria1" size="1" onchange="javascript: reload();" class="FormTable" onblur="MM_validateForm('categoria1','','R')">
                      <option  value="0">----</option>
                      
                      <% While (NOT RsTipoProblema.EOF) 
                      			If vDataForm("lang") = "esp" or session("ididioma") = 1 Then
                      	%>
                      						<option value="<%=(RsTipoProblema.Fields.Item("IDProblema").Value)%>" <% if (CStr(request.form("categoria1")) = CStr(RsTipoProblema.Fields.Item("IDProblema").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(RsTipoProblema.Fields.Item("DetalleProblema").Value) %></option>
                      			<% Else %>
                      	
                      						<option value="<%=(RsTipoProblema.Fields.Item("IDProblema").Value)%>" <% if (CStr(request.form("categoria1")) = CStr(RsTipoProblema.Fields.Item("IDProblema").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(RsTipoProblema.Fields.Item("ProblemDetail").Value) %></option>
                      
                      			<% End If %>
                      
                        <%
							  RsTipoProblema.MoveNext()
								Wend
						%>
                    </select>
                    
                  </td>
                  <td width="10%" ><% Call ReadLang(Rs_Lenguaje,4) %></td>
                  <td > 
                    <select name="categoria2" size="1" onChange="javascript: reload();" class="FormTable">
                      <option  value="0">----</option>
                      <% If  Request.Form ("categoria1") <> 0 Then
							While (NOT RsCategoria1.EOF)
										If vDataForm("lang") = "esp" or session("ididioma") = 1 Then
					 %>
						<option value="<%=(RsCategoria1.Fields.Item("IDProblema").Value)%>" <%if (CStr(request.form("categoria2")) = CStr(RsCategoria1.Fields.Item("IDProblema").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(RsCategoria1.Fields.Item("DetalleProblema").Value)%></option>
						
					<% Else %>
					
					  <option value="<%=(RsCategoria1.Fields.Item("IDProblema").Value)%>" <%if (CStr(request.form("categoria2")) = CStr(RsCategoria1.Fields.Item("IDProblema").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(RsCategoria1.Fields.Item("ProblemDetail").Value)%></option>
					  
					  <% End If %>

                    <%
					  RsCategoria1.MoveNext()
						Wend
			     	End If
	 				 %>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td >&nbsp;</td>
                  <td >&nbsp;</td>
                  <td >&nbsp;</td>
                </tr>
                <tr> 
                  <td>&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,5) %></td>
                  <td colspan="3" > 
                    <select name="categoria3" size="1" class="FormTable">
                      <option value="0">----</option>
                      <%
							If Request.Form("Categoria2")<>0  Then
										
										While (NOT RsCategoria2.EOF)
															If vDataForm("lang") = "esp" or session("ididioma") = 1 Then
																	Response.Write (" <option value=""" )
																	Response.Write (RsCategoria2.Fields.Item("IDProblema").Value & """> ")
																	Response.Write (RsCategoria2.Fields.Item("DetalleProblema").Value )
																	Response.Write ("</option>")
															Else
																	Response.Write (" <option value=""" )
																	Response.Write (RsCategoria2.Fields.Item("IDProblema").Value & """> ")
																	Response.Write (RsCategoria2.Fields.Item("ProblemDetail").Value )
																	Response.Write ("</option>")
															End If
																				
										  RsCategoria2.MoveNext()
										Wend
										
								End If
							%>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td >&nbsp;</td>
                  <td >&nbsp;</td>
                  <td >&nbsp;</td>
                </tr>
                <tr> 
                  <td>&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,6) %></td>
                  <td colspan="3"> 
                    <%
											 If    (Len(Request.Form("observaciones"))=0) Then 
				  							    Response.Write ("<textarea name=""observaciones"" maxlength=""500"" rows=""9"" cols=""85"" class=""FormTable"" onKeyDown=""CountLeft(this.form.observaciones,this.form.left,500);""  onKeyUp=""CountLeft(this.form.observaciones,this.form.left,500);""></textarea>")
											 Else
														Response.Write ("<textarea name=""observaciones"" maxlength=""500"" rows=""9"" cols=""85"" class=""FormTable"" onKeyDown=""CountLeft(this.form.observaciones,this.form.left,500);""  onKeyUp=""CountLeft(this.form.observaciones,this.form.left,500);"">")
														Response.Write Request.Form("observaciones") 
														Response.Write (" </textarea> ")
											 End If
										%>
                  </td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td >&nbsp;</td>
                  <td >&nbsp;</td>
                  <td ><input readonly type="text" name="left" size="3" maxlength="3" value="500" class="FormTable">  <% Call ReadLang(Rs_Lenguaje,34) %></td>
                </tr>
                <tr> 
                  <td colspan="4"> 
                    <div align="center">
                    	<% 	Response.Write ("<input type=""button"" name=""Button"" value=""")
		                    	Call ReadLang(Rs_Lenguaje,35) 
		                    	Response.Write (""" class=""formbutton"" onclick=""javascript: submitfrm();"">")
                    	%>
                    	<% 	Response.Write ("<input type=""reset"" name=""Limpiar"" value=""")
		                    	Call ReadLang(Rs_Lenguaje,36) 
		                    	Response.Write (""" class=""formbutton"" onclick=""javascript: submitfrm();"">")
                    	%>
                      
                      
                    </div>
                  </td>
                </tr>
              </table>
                    
                 
                  <div align="left"></div>
                </td>
              </tr>
            </table>
   </form>
		      
        
<!--#INCLUDE file="library/PageClose.asp"-->

<%
RsTipoProblema.Close()
If  Request.Form ("Categoria1") <> 0 Then
rscategoria1.Close
end if
If  Request.Form ("Categoria2") <> 0 Then
rscategoria2.Close
end if
%>