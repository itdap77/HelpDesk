<%@ language=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") <> 1) Then
	If vDataForm("lang") = "esp" Then
		Response.Redirect ("mensaje.asp?Msg=Usted no esta logueado en el sistema.&backtourl=default.asp")
	Else
		Response.Redirect ("mensaje.asp?Msg=You are not logged to the system.&backtourl=default.asp")
	End If
End If 

'ValidUserAction "ABM", "ABM"

If Not Len(vDataForm) = 0 Then 
	Set vDataForm = vDataForm 
Else 
	Set vDataForm = request.QueryString 
End If 



' Categorias

 If vDataForm("categoria_a") <> 0 and vDataForm ("categoria_a") <> 1 Then
 
 	
	
	'  set RS_Problema = Server.CreateObject("ADODB.Recordset")
	'        RS_Problema.ActiveConnection = MM_HelpDesk_STRING
	'        RS_Problema.Source = "{call dbo.SPU_Problema_CatXId(" & vDataForm("categoria_a") & ")}"
	'        RS_Problema.Open()
    '    							
	'			       
    '        if vDataForm ("Problema_a") <> 0 Then
    '        	
	'            set RS_Problema2 = Server.CreateObject("ADODB.Recordset")
	'            RS_Problema2.ActiveConnection = MM_HelpDesk_STRING
	'            RS_Problema2.Source = "{call dbo.SPU_Problema_Id(" & vDataForm("Problema_a") & ")}"
	'            RS_Problema2.Open()
	'           
    '        End if
            
            
	Select case vDataForm ("categoria_a")
					
		case 2
			set RS_Problema = Server.CreateObject("ADODB.Recordset")
			RS_Problema.ActiveConnection = MM_HelpDesk_STRING
			RS_Problema.Source = "{call dbo.SPU_Problema_Cat2('%')}"
			RS_Problema.Open()
			
			'set RS_Problema2 = Server.CreateObject("ADODB.Recordset")
		    'RS_Problema2.ActiveConnection = MM_HelpDesk_STRING
			'RS_Problema2.Source = "{call dbo.SPU_Problema_Cat2('%')}"
			'RS_Problema2.Open()	
				
		case 3
			set RS_Problema = Server.CreateObject("ADODB.Recordset")
			RS_Problema.ActiveConnection = MM_HelpDesk_STRING
			RS_Problema.Source = "{call dbo.SPU_Problema_Cat3('%')}"
			RS_Problema.Open()
			
			'set RS_Problema2 = Server.CreateObject("ADODB.Recordset")
			'RS_Problema2.ActiveConnection = MM_HelpDesk_STRING
			'RS_Problema2.Source = "{call dbo.SPU_Problema_Cat3('%')}"
	'	RS_Problema2.Open()
			
	end select

End if 


'Recordset modificacion
if vDataForm("categoria_m") <> 0 then

           
	        set RS_Problema_m = Server.CreateObject("ADODB.Recordset")
	        RS_Problema_m.ActiveConnection = MM_HelpDesk_STRING
	        RS_Problema_m.Source = "{call dbo.SPU_Problema_CatXId(" & vDataForm("categoria_m") & ")}"
	        RS_Problema_m.Open()
        							
				       
            if vDataForm ("idproblema_m") <> 0 Then
            	
	            set RS_Problema_Id = Server.CreateObject("ADODB.Recordset")
	            RS_Problema_Id.ActiveConnection = MM_HelpDesk_STRING
	            RS_Problema_Id.Source = "{call dbo.SPU_Problema_Id(" & vDataForm("idproblema_m") & ")}"
	            RS_Problema_Id.Open()
	           
            End if
	        
End if



'Recordset para la baja


if vDataForm ("categoriab") <> 0  then 'baja
	
	set RS_Problema_b = Server.CreateObject("ADODB.Recordset")
	RS_Problema_b.ActiveConnection = MM_HelpDesk_STRING
	RS_Problema_b.Source = "{call dbo.SPU_Problema_CatXIdE(" & vDataForm("categoriab") & ")}"
	RS_Problema_b.Open()
			
	set RS_Problema2 = Server.CreateObject("ADODB.Recordset")
	RS_Problema2.ActiveConnection = MM_HelpDesk_STRING
	RS_Problema2.Source = "{call dbo.SPU_Problema_CatXIdE(" & vDataForm("categoriab") & ")}"
	RS_Problema2.Open()
	
	
end if

	set Rs_Categoria = Server.CreateObject("ADODB.Recordset")
	Rs_Categoria.ActiveConnection = MM_HelpDesk_STRING
	Rs_Categoria.Source = "{call dbo.SPU_Problema_Cat}"
	Rs_Categoria.Open()


	set Rs_Problema3 = Server.CreateObject("ADODB.Recordset")
	Rs_Problema3.ActiveConnection = MM_HelpDesk_STRING
	Rs_Problema3.Source = "{call dbo.SPU_Problema_VEE}"
	Rs_Problema3.Open()


'Fin de baja
%>

<script type="text/javascript" >
    var lang = getQuerystring('lang');
    var validresult;

function submitfrmalta() {
    
    var langtext_a,alerttext_a;

    validresult = MM_validateForm('problemdetail_a', '', 'R', 'detalleproblema_a', '', 'R');

    if (lang == 'esp') {
        langtext_a = 'Usted esta seguro que quiere agregar el problema ingresado ?';
        alerttext_a = 'Los campos Nivel y problema a agregar son obligatorios.';
    }
    else {
        langtext_a = 'Are you sure that you want to insert the problem ?';
        alerttext_a = 'The level field is required.';
    }

    if (validresult == 0) {
        
                if (altaproblema.categoria_a != 0)
                    alert(alerttext_a);
                else {
                    if (confirm(langtext_a)) {
                        altaproblema.action = 'vbsproblema.asp';
                        altaproblema.submit();
                    }
                }
    }
}


function submitfrmbaja() {

    var langtext, alerttext;
    
        
    if (lang== 'esp') {
        langtext = 'Usted esta seguro que quiere borrar el problema seleccionado ?';
        alerttext = 'Los campos Nivel y problema a borrar son obligatorios.';
    }
    else {
        langtext = 'Are you sure that you want to delete the selected problem ?';
        alerttext = 'The level and problem fields are required for deletion.';
    }

    if ((bajaproblema.categoriab.value == 0) || (bajaproblema.problemab.value == 0))
        alert(alerttext);
    else {
        if (confirm(langtext)) {
            bajaproblema.action = 'vbsproblema.asp';
            bajaproblema.submit();
        }
    }
}

function submitfrmmodif() {

    var langtext_m, alerttext_m, validresult;

    validresult = MM_validateForm('detalleproblema_m', '', 'R', 'problemdetail_m', '', 'R');
    
    if (lang == 'esp') {
        langtext_m = 'Usted esta seguro que quiere modificar el problema seleccionado ?';
        alerttext_m = 'Los campos son obligatorios.';
    }
    else {
        langtext_m = 'Are you sure that you want to modify the selected problem ?';
        alerttext_m = 'The fields are required.';
    }

    if (validresult == 0) {
        if ((modificacionproblema.categoria_m.value == 0) || (modificacionproblema.idproblema_m.value == 0))
            alert(alerttext_m);
        else {
            if (confirm(langtext_m)) {
                modificacionproblema.action = 'vbsproblema.asp';
                modificacionproblema.submit();
            }
        }
    }
}


function reloada()
{
altaproblema.submit();
}

function reloadb()
{
bajaproblema.submit();
}

function reloadm() {
modificacionproblema.idproblema_m.value = '0';
modificacionproblema.submit();
}

function reloadmodproblema()
{

modificacionproblema.submit();
}

</script>

      <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td class="Titulo"><% Call ReadLang(Rs_Lenguaje,163)  %></td>
          <td> </td>
        </tr>
      </table>
      
          <table class="ContentArea" border="0" cellspacing="0" cellpadding="0">
           <tr>
         		 <td colspan="4"> 
         		 
         		 <!-------------------------- Alta problema ----------------------------------->
            <form name="altaproblema" method="post" action="">
            <input type="hidden" name="accion_problema" value="A"/>
            <input type="hidden" name="lang" value='<%=vDataForm("lang")%>'/>

              <table width="100%" border="0" class="FormBoxHeader">
                <tr> 
                  <td width="10"><img alt="" src="images/gripgray.gif" width="10" height="13"/></td>
                  <td><% Call ReadLang(Rs_Lenguaje,150)  %></td>
                </tr>
              </table>
              
              <table  border="0" class="FormBoxBody" width="100%" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="30%" colspan="4">&nbsp;</td>
                 </tr>
                <tr> 
                  <td  >&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,153)  %></td>
                  <td > 
                  	<% If len(vDataForm ("detalleproblema_a")) <> 0 Then
                  	
                  	        Response.Write ("<input type=""text"" name=""detalleproblema_a"" class=""FormTable"" size=""20"" maxlength=""40"" value=")
                  	        Response.Write vDataForm("detalleproblema_a") 
                  	        Response.Write (">")

                  	        Else
                  	        Response.Write ("<input type=""text"" name=""detalleproblema_a"" class=""FormTable"" size=""20"" maxlength=""40"">" )
                  	        End If
                  	%>
                    
                  </td>
                  <td >&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,165)  %></td>
                  <td>
                         <input type="text" name="problemdetail_a" size="20" maxlength="40" class="FormTable" value="<%  If len(vDataForm("problemdetail_a")) <> 0 then Response.Write vDataForm("problemdetail_a") End if %>" />
                         
                  	</td>
                </tr>
                <tr> 
                  <td >&nbsp;&nbsp;&nbsp; <% Call ReadLang(Rs_Lenguaje,154)  %> 
                  </td>
                  <td > 
                    <select name="categoria_a" class="FormTable" onchange="reloada();">
		                      <option value="0">----</option>
		                      <%
			                        Rs_Categoria.MoveFirst()
									While (NOT Rs_Categoria.EOF)
							  %>
		                              <option value="<%=(Rs_Categoria.Fields.Item("Categoria").Value)%>"
		                              <% 
													        if (CStr(vDataForm("categoria_a")) = CStr(Rs_Categoria.Fields.Item("categoria").Value)) then 
														        Response.Write("SELECTED") 
													        Else 
														        Response.Write("")
													        End If
		                              %>><%Response.write (Rs_Categoria.Fields.Item("Categoria").Value)%></option>
		                              <%
		                                  Rs_Categoria.MoveNext()
								          Wend
								        %>
								
                    </select>
                  </td>
                   <td  align="center"><% Call ReadLang(Rs_Lenguaje,156)  %></td>
                  <td>
                  <input type="text" name="estimacion_a" size="6" class="FormTable" maxlength="3" onBlur="MM_validateForm('estimacion','','NisNum');return document.MM_returnValue" value="<%  If len(vDataForm("estimacion_a")) <> 0 then Response.Write vDataForm("estimacion_a") End if %>"  />
                 <font face="Verdana" size="1"><% Call ReadLang(Rs_Lenguaje,164)  %></font>
                  </td>
                </tr>
                <tr>
                <% If vDataForm("categoria_a") <> 1 and vDataForm("categoria_a") <> 0 then %>
                
                    <td>&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,155)  %></td>
                      <td colspan="3">
                      
                              <select name="problema_a" id="problema_a" class="FormTable">
                                        <% 
								            if vDataForm ("categoria_a") <> 0 and vDataForm ("categoria_a") <> 1 Then
			                            Rs_Problema.MoveFirst()
							            While NOT RS_Problema.EOF
							            %>
                                      <option value="<%=(RS_Problema.Fields.Item("IDProblema").Value)%>"> 
                                    <%if vDataForm("lang") = "esp" then
                                            Response.Write (RS_Problema.Fields.Item("detalleproblema").Value)
                                          Else
                                            Response.Write (RS_Problema.Fields.Item("problemdetail").Value)
                                          End If
				                    %>
                                          </option>
		                                     <%
						              RS_Problema.MoveNext()
						            Wend
					            end if												
						             %>
                                </select>
                    
                    <% Else 
                            Response.Write ("<td colspan=""4"">&nbsp;</td>")
                            Response.Write ("<input type=""hidden"" name=""problema_a"" value=""0"" />")
                            End If
                    %>
                    
                    </td>
                    </tr>
                
                
                
                  <tr><td colspan="4"></td></tr>
                  <tr>
                  
                           <%If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if is the demo account %>
                                 <td colspan="4" align="center">   
                         <% 	Response.Write ("<input type=""button"" name=""ingresar_a"" value=""")
                	                        Call ReadLang(Rs_Lenguaje,159) 
                	                        Response.Write (""" class=""formbutton"" onclick=""JavaScript:submitfrmalta();""/>")
        	                        %>&nbsp;
        	                        <% 	Response.Write ("<input type=""reset"" name=""limpiar_a"" value=""")
                	                        Call ReadLang(Rs_Lenguaje,160) 
                	                        Response.Write (""" class=""formbutton"" />")
        	                        %>
                  </td>
                                   <% Else %>
                                    <td width="20%" colspan="4" style="color:red;text-align:center;font-weight:bold;">
                                    <div align="center">
                                        <% Call ReadLang(Rs_Lenguaje,207) %>
                                    </div></td>
                                  <%End if %>    
                        

                </tr>
              </table>
              <table width="100%" border="0">
                <tr>
                  <td>&nbsp;</td>
                </tr>
              </table>
            </form>
            
            <!-------------------------- Fin Alta Problema ----------------------------------->
            
            <!--------------------------------- BORRADO PROBLEMA ------------------------------------------------>
            
            <form name="bajaproblema" action="" method="post">
                   <input type="hidden" name="accion_problema" value="B"/>
                   <table border="0" class="FormBoxHeader">
                    <tr> 
                      <td width="10"><img alt="" src="images/gripgray.gif" width="10" height="13"/></td>
                      <td><% Call ReadLang(Rs_Lenguaje,151)  %></td>
                    </tr>
                  </table>
                   <table width="100%" border="0" class="FormBoxBody" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="4" >&nbsp;</td>
                </tr>
                <tr> 
                  <td width="20%" >&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,154)  %></td>
                  <td > 
                    <select name="categoriab" class="FormTable" onchange="reloadb();">
                      <option value="0">----</option>
                                    <%
                                        Rs_Categoria.MoveFirst()
                                        
										While (NOT Rs_Categoria.EOF)
										    Response.Write "<option value=""" & Rs_Categoria.Fields.Item("Categoria").Value & """"
					                      
											If (CStr(vDataForm("categoriab")) = CStr(Rs_Categoria.Fields.Item("categoria").Value)) then 
												Response.Write(" SELECTED>") 
											Else 
												Response.Write(">")
											End If
					                      
					                        Response.write (Rs_Categoria.Fields.Item("Categoria").Value) & "</option>"
					                      
										    Rs_Categoria.MoveNext()
										Wend
										%>
                    </select>
                  </td>
                  <td width="20%" > <% Call ReadLang(Rs_Lenguaje,170)  %> </td>
                  <td > 
                    <select name="problemab" class="FormTable">
                      <option value="0">----</option>
                      <% If vDataForm ("categoriab") <> 0 Then
							        RS_Problema_b.MoveFirst()
									While NOT RS_Problema_b.EOF
									        Response.Write "<option value=""" & RS_Problema_b.Fields.Item("IDProblema").Value & """>" 
                                            if vDataForm("lang") = "esp" then                  
                                                Response.Write ( RS_Problema_b.Fields.Item("DetalleProblema").Value)
               		                        Else
                                                Response.Write ( RS_Problema_b.Fields.Item("ProblemDetail").Value)
                                            End If
                                            Response.Write "</option>"
								    RS_Problema_b.MoveNext()
								Wend
							End if												
						%>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td colspan="4">&nbsp;</td>
                </tr>
                <tr> 
                       <%If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if is the demo account %>
                  <td colspan="4"> 
                    <div align="center"> 
                                    <% 	Response.Write ("<input type=""button"" name=""Eliminar"" value=""")
                	                        Response.Write "Eliminar"
                	                        Response.Write (""" class=""formbutton"" onclick=""javascript: submitfrmbaja();""/>")
        	                        %>
                    </div>
                     <% Else %>
                                    <td width="20%" colspan="4" style="color:red;text-align:center;font-weight:bold;">
                                    <div align="center">
                                        <% Call ReadLang(Rs_Lenguaje,207) %>
                                    </div></td>
                                  <%End if %>    
                        


                  </td>
                </tr>
              </table>
                   <table width="100%" border="0">
                <tr>
                  <td>&nbsp;</td>
                </tr>
              </table>
            </form>
            
 <!--------------------------- MODIFICACION PROBLEMA --------------------------------------------->
            
            <form name="modificacionproblema" action="" method="post">
            <input type="hidden" name="accion_problema" value="M"/>
            <input type="hidden" name="lang" value='<%=vDataForm("lang")%>'/>
              <table width="100%" border="0" class="FormBoxHeader">
                <tr> 
                  <td width="5"><img src="images/gripgray.gif" width="10" height="13" alt="" /></td>
                  <td><% Call ReadLang(Rs_Lenguaje,152)  %></td>
                </tr>
              </table>
              <table width="100%" border="0" class="FormBoxBody" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr> 
                  <td >&nbsp;&nbsp;&nbsp; <% Call ReadLang(Rs_Lenguaje,154)  %></td>
                  <td> 
                  <!------------------- Categorias --------------->
                    <select name="categoria_m" class="FormTable" onchange="Javascript:reloadm();">
                      <option value="0">----</option>
								            
								       <%
                                        Rs_Categoria.MoveFirst()
                                        
										While (NOT Rs_Categoria.EOF)
										    Response.Write "<option value=""" & Rs_Categoria.Fields.Item("Categoria").Value & """"
					                      
											If (CStr(vDataForm("categoria_m")) = CStr(Rs_Categoria.Fields.Item("categoria").Value)) then 
												Response.Write(" SELECTED>") 
											Else 
												Response.Write(">")
											End If
					                      
					                        Response.write (Rs_Categoria.Fields.Item("Categoria").Value) & "</option>"
					                      
										    Rs_Categoria.MoveNext()
										Wend
										%>   
								            
								           
                    </select>
                    <!------------------- FIN Categorias --------------->
                  </td>
                </tr>
                <tr> 
                  <td width="35%">&nbsp;&nbsp;&nbsp; <% Call ReadLang(Rs_Lenguaje,155)  %></td>
                  <td > 
                  <!------------------- Top-categoria --------------->
                    <select name="idproblema_m" id="idproblema_m" class="FormTable" onchange="JavaScript:reloadmodproblema();">
                      <option value="0">----</option>
					
                      <% If vDataForm ("categoria_m") <> 0 Then
							        RS_Problema_m.MoveFirst()
									While NOT RS_Problema_m.EOF
									        Response.Write "<option   value=""" & RS_Problema_m.Fields.Item("IDProblema").Value & """" 
									        
									        If (CStr(vDataForm("idproblema_m")) = CStr(RS_Problema_m.Fields.Item("IDProblema").Value)) then 
												Response.Write(" SELECTED>") 
												If vDataForm ("categoria_m") = 1 then
                                                      if vDataForm("lang") = "esp" then                    
												            Response.Write ( RS_Problema_m.Fields.Item("detalleproblema").Value)
                                                       Else
                                                             Response.Write ( RS_Problema_m.Fields.Item("ProblemDetail").Value)
                                                       End If
												Else        		                                
												    if vDataForm("lang") = "esp" then                  
												        Response.Write ( RS_Problema_m.Fields.Item("topproblema").Value)       
                                                    else
                          					            Response.Write ( RS_Problema_m.Fields.Item("topproblema_eng").Value)       
                                                    end if
												End If
               		                        Else
               		                            Response.Write(">")
               		                            If vDataForm ("categoria_m") = 1 then
                                                    if vDataForm("lang") = "esp" then                  
												        Response.Write ( RS_Problema_m.Fields.Item("detalleproblema").Value)       
                                                    Else
                                                             Response.Write ( RS_Problema_m.Fields.Item("problemdetail").Value)
                                                    End If
												Else    
                                                    if vDataForm("lang") = "esp" then                  
												        Response.Write ( RS_Problema_m.Fields.Item("topproblema").Value)       
                                                    else
                          					            Response.Write ( RS_Problema_m.Fields.Item("topproblema_eng").Value)       
                                                    end if
												End If
               		                         	
											End If
               		                        Response.Write "</option>"
								    RS_Problema_m.MoveNext()
								Wend
								
														
							End if												
						%>
				
								
                    </select>
                    <!------------------- FIN top-categorias --------------->
                  </td>
                </tr>
                <tr> 
                  <td> 
                    <p align="left">&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,153)  %></p>
                  </td>
                  <td> 
                  <!------------------- Nuevo Nombre de problema --------------->
                  
                   <% 
                   If vDataForm ("categoria_m") <> 0 Then
                           If (vDataForm ("idproblema_m") <> 0)  Then
									        Response.Write("<input name=""detalleproblema_m"" width=""40"" size=""40""  class=""FormTable"" maxlength=""40"" value=""")
									        Response.Write(RS_Problema_id.Fields.Item("DetalleProblema").Value) 
									        Response.Write (""">")
        											
		  				        Else
		  					                Response.Write("<input name=""detalleproblema_m"" type=""text"" width=""40"" size=""40"" class=""FormTable"" maxlength=""40"" value=""")
		  					                Response.Write (""">")
		  				        End if	
		  		    Else
		  		    Response.Write("<input name=""detalleproblema_m"" type=""text"" width=""40"" size=""40"" class=""FormTable"" maxlength=""40"" value=""")
		  					                Response.Write (""">")		
		  			End If	
									 %>
					<!------------------- FIN Nuevo Nombre de problema --------------->
                  </td>
                </tr>
                 <tr>
                 <td><!------------------- New name of the problem --------------->
                 &nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,165)  %></td>
                 <td>
                
                   <% 
                   If vDataForm ("categoria_m") <> 0 Then
                           If (vDataForm ("idproblema_m") <> 0)  Then
									        Response.Write("<input name=""problemdetail_m"" width=""40"" size=""40""  class=""FormTable"" maxlength=""40"" value=""")
									        Response.Write(RS_Problema_id.Fields.Item("problemdetail").Value) 
									        Response.Write (""">")
        											
		  				        Else
		  					                Response.Write("<input name=""problemdetail_m"" type=""text"" width=""40"" size=""40"" class=""FormTable"" maxlength=""40"" value=""")
		  					                Response.Write (""">")
		  				        End if	
		  		    Else
		  		                Response.Write("<input name=""problemdetail_m"" type=""text"" width=""40"" size=""40"" class=""FormTable"" maxlength=""40"" value=""")
		  					    Response.Write (""">")		
		  			End If	
									 %>
					<!------------------- FIN New name of the problem --------------->
                 </td></tr>
                <tr> 
                  <td>&nbsp;&nbsp;&nbsp; <% Call ReadLang(Rs_Lenguaje,156)  %></td>
                  <td> 
                    <input type="text" name="Estimacion" size="6" class="FormTable" maxlength="3" value="<% 
                    If vDataForm ("idproblema_m") then 
                    response.Write RS_Problema_Id.fields.item("estimacion").value
                    End If
                    %>"/>
                    <font face="Verdana" size="1"><% Call ReadLang(Rs_Lenguaje,164)  %></font> </td>
                </tr>
                <tr> 
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr> 
                    <%If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if is the demo account %>
                  <td colspan="2"> 
                    <div align="center">&nbsp; 

                          <% 	Response.Write ("<input type=""button"" name=""Modificar"" value=""")
                	                        Call ReadLang(Rs_Lenguaje,162) 
                	                        Response.Write (""" class=""formbutton"" onclick=""javascript: submitfrmmodif();""/>")
        	                        %>&nbsp;
        	                        <% 	Response.Write ("<input type=""reset"" name=""Limpiar"" value=""")
                	                        Call ReadLang(Rs_Lenguaje,160) 
                	                        Response.Write (""" class=""formbutton"" />")
        	                        %>
        	       <% Else %>
                                    <td width="20%" colspan="4" style="color:red;text-align:center;font-weight:bold;">
                                    <div align="center">
                                        <% Call ReadLang(Rs_Lenguaje,207) %>
                                    </div></td>
                                  <%End if %>    
                        
                    </div>
                  </td>
                </tr>
              </table>
            </form>
               </td>
                  </tr>
                </table>
   

<!--#INCLUDE file="library/PageClose.asp"-->
<%
if vDataForm ("categoria") <> 0 and vDataForm ("categoria") <> 1 Then
	RS_Problema.Close()
End if

%>