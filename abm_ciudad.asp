<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

ValidSession()

'ValidUserAction "ABM", "ABM"

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

If  vDataForm("IdCiudad") <> 0 Then
	set Rs_IDCPais_L = Server.CreateObject("ADODB.Recordset")
	Rs_IDCPais_L.ActiveConnection = MM_HelpDesk_STRING
	Rs_IDCPais_L.Source = "{call dbo.SPU_IDCPais_L(" + Replace(vDataForm("IdCiudad"), "'", "''") + ")}"
	Rs_IDCPais_L.Open()
End if

dim cambio


set RsPais = Server.CreateObject("ADODB.Recordset")
RsPais.ActiveConnection = MM_HelpDesk_STRING
RsPais.Source = "{call dbo.SPU_Pais_VE(" & 0 & "," & 1 & ")}"
RsPais.Open()


set Rs_Ciudad_V = Server.CreateObject("ADODB.Recordset")
Rs_Ciudad_V.ActiveConnection = MM_HelpDesk_STRING
Rs_Ciudad_V.Source = "{call dbo.SPU_Ciudad_VE}"
Rs_Ciudad_V.Open()

%>

<script type="text/javascript" src="" language="JavaScript">
 
var lang = getQuerystring("lang");

function submitfrmalta()
{
var validresult;
validresult = MM_validateForm('newciudadalta', '', 'R', 'sigla', '', 'R');

if (validresult == 0) {    
                                var paisalta = document.getElementById('paisalta');

                                if (paisalta.value != 0)
                                    {
                                        altaciudad.action = 'vbsciudad.asp' ;
                                        altaciudad.submit();
                                     }
                                }
}

function reload(cambio)
{
modificacionciudad.submit();
}

function submitfrmbaja() 
{
var ciudadesbaja = document.getElementById('ciudadesbaja');


    if (lang == 'eng')
    {
        if (ciudadesbaja.value != 0)
            {
                if (confirm("Are you sure that you want to delete the selected item ?")) 
                {
                    bajaciudad.action = "vbsciudad.asp";
                    bajaciudad.submit();
                 }
            } else{
                    alert("You need to choose a city to be deleted.")
                    }   
    }
    else 
    {
            if (ciudadesbaja.value != 0)
            {
                    if (confirm('Esta seguro que desea borrar el elemento seleccionado ?')) 
                    {
                        bajaciudad.action = 'vbsciudad.asp';
                        bajaciudad.submit();
                    }       
             } else{
                    alert('Nesecita seleccionar una ciudad para ser borrada.')
                    }
    }

}

function submitfrmmodif()
{
var IdCiudad = document.getElementById('IdCiudad');
var paismod =  document.getElementById('paismod');
MM_validateForm('ciudadmod', '', 'R');

    if (lang == 'eng')
    {
        if (IdCiudad.value != 0 && paismod.value != 0)
            {
                if (confirm('Are you sure that you want to modify the selected item ?')) 
                {
                    modificacionciudad.action = 'vbsciudad.asp' ;
                    modificacionciudad.submit();
                 }
            }    
    }
    else 
    {
            if (IdCiudad.value != 0 && paismod.value != 0)
            {
                    if (confirm('Esta seguro que desea modificar el elemento seleccionado ?')) 
                    {
                        modificacionciudad.action = 'vbsciudad.asp' ;
                        modificacionciudad.submit();
                    }       
             } 
    }
    

}

</script>
          <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo"><% Call ReadLang(Rs_Lenguaje,134)  %> </td>
            </tr>
          </table>
          
          <table class="ContentArea" cellspacing="0" cellpadding="0" >
            <tr> 
                  <td width="100%" height="160px;"> 
                        
                            <form name="altaciudad" method="post" action="">
                                <input name="accion" type="hidden" value="A"/>
                                
                                <input type="hidden" name="telefonoalta" size="20" class="FormTable" maxlength="30"/>
                                <input type="hidden" name="internoalta" size="20" class="FormTable" maxlength="10" />
                                
                              <table  class="FormBoxHeader">
                                <tr>
                                  <td width="5" height="16"><img src="images/gripgray.gif" width="10" height="13"></td> 
                                  <td> 
                                    <% Call ReadLang(Rs_Lenguaje,119)  %> 
                                  </td>
                                </tr>
                              </table>
                              
                              <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormBoxBody">
                                <tr>
                                  <td width="25%" colspan="4">&nbsp;</td>
                                  
                                </tr>
                                <tr> 
	                                  <td> 
		                                    &nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,121)  %>
		                                  
	                                  </td>
	                                  <td><select size="1" name="paisalta" class="FormTable">
                              		        	    <option selected value="0">----</option>
					                                    <%
														    While (NOT RsPais.EOF)
														    %>
					                                    <option value="<%=(RsPais.Fields.Item("IDPais").Value)%>"><%=(RsPais.Fields.Item("Pais").Value)%></option>
					                                     <%
															    RsPais.MoveNext()
															    Wend
															    If (RsPais.CursorType > 0) Then
															      RsPais.MoveFirst
															    Else
															      RsPais.Requery
															    End If
													    %>
                              	    </select>
						                         
						                        </td>
	                                  <td><% Call ReadLang(Rs_Lenguaje,120)  %> (*)</td>
	                                  <td><input type="text" name="newciudadalta" size="20" class="FormTable" maxlength="30"/> 
						                          
						                        </td>
                                </tr>
                                <tr>
                                  <td colspan="4">&nbsp;
                                  </td>
                                  
                                </tr>
							    <tr>	
								    <td >&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,122)  %> (*)</td>
                                    <td><input type="text" name="sigla" size="4" class="FormTable" maxlength="3"/></td>
                                     <td>&nbsp;</td>
                                     <td>&nbsp;</td>
							    </tr>
                                <tr>
                                  <td colspan="4">&nbsp;</td>
                                </tr>
                                <tr>
                                  
                                    
                                    <%If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if is the demo account %>
                                    <td colspan="4">
                                        <div align="center">
                                        <% 	
                                        Response.Write ("<input type=""button"" name=""Ingresar"" value=""")
                	                        Call ReadLang(Rs_Lenguaje,123) 
                	                        Response.Write (""" class=""formbutton"" onclick=""javascript: submitfrmalta();"">")
        	                        %>&nbsp;
        	                        <% 	Response.Write ("<input type=""reset"" name=""Limpiar2"" value=""")
                	                        Call ReadLang(Rs_Lenguaje,124) 
                	                        Response.Write (""" class=""formbutton"" >")
        	                        %></div></td>
                                        <%Else %>
                                        
                                            <td width="20%" colspan="4" style="color:red;text-align:center;font-weight:bold;">
                                                <div align="center">
                                            <% Call ReadLang(Rs_Lenguaje,205) %>
                                            </div> </td>
            	                    <%End If %>
                                </tr>
                              </table>
                            </form>
                            
                 </td>
            </tr>
            <tr>
            	<td>
            	        <form name="bajaciudad" action="" method="post">
                        <input name="accion" type="hidden" value="B" />
                        
              
                          <table class="FormBoxHeader">
                            <tr>
                              <td width="5"><img alt="" src="images/gripgray.gif" width="10" height="13"/></td> 
                              <td><p><b><% Call ReadLang(Rs_Lenguaje,125)  %> </b></p>
                              </td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormBoxBody">
                            <tr>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td width="33%">&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,126)  %> </td>
                              <td width="33%"> 
                                <select name="ciudadesbaja" class="FormTable">
                                          <Option selected value="0">----</option>
                                          <%
						  	              While (NOT Rs_Ciudad_V.EOF)
							              %>
                                          <option value="<%=(Rs_Ciudad_V.Fields.Item("IDCiudad").Value)%>"><%=(Rs_Ciudad_V.Fields.Item("Ciudad").Value)%></option>
                                          <%
							              Rs_Ciudad_V.MoveNext()
							            Wend
							            If (Rs_Ciudad_V.CursorType > 0) Then
							              Rs_Ciudad_V.MoveFirst
							            Else
							              Rs_Ciudad_V.Requery
							            End If
							            %>
                                </select>
                              </td>
                              
                                   <%If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if is the demo account %>                                                                              <td style="text-align:center;">                               
                                     <% 	Response.Write ("<input type=""button"" name=""Eliminar"" value=""")
                	                        Call ReadLang(Rs_Lenguaje,127) 
                	                        Response.Write (""" class=""formbutton"" onclick=""javascript: submitfrmbaja();"">")
        	                        %>
                                       </td>
                            
                                        <tr><td colspan="3">&nbsp;</td></tr>
                                <%Else %>
                              <td>&nbsp;</td>
                                <tr>
                              <td width="20%" colspan="4" style="color:red;text-align:center;font-weight:bold;">
                                          <% Call ReadLang(Rs_Lenguaje,205) %>
                                     </td>
        	                      <%End if %>
                            
                              </tr>
                          </table>
                          
                          
                        </form>
                        
                   
                        <form name="modificacionciudad" action="" method="post">
                                <input name="accion" type="hidden" value="M"/>
                                
                                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td>&nbsp;</td>
                                    </tr>
                                  </table>
                                  
                                  <table class="FormBoxHeader">
                                    <tr>
                                      <td width="5" height="16"><img alt="" src="images/gripgray.gif" width="10" height="13"/></td>
                                      <td><p><b><% Call ReadLang(Rs_Lenguaje,128)  %> </b></p>
                                      </td>
                                    </tr>
                                  </table>
                                  
						          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormBoxBody">
                                        <tr>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          <td><input type="hidden" name="lang" value="<%=vDataForm("lang") %>" /></td>
                                        </tr>
                                        <tr>
                                          <td> &nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,129)  %> </td>
                                          <td><select size="1" name="IdCiudad" class="FormTable" onchange="javascript:reload();">
                                                        <option value="0">----</option>
                                                        <%
				                                        While (NOT Rs_Ciudad_V.EOF)
				                                        %>
                                                        <option value="<% =Rs_Ciudad_V.Fields.Item("IDCiudad").Value %>"
				                                        <% if (CStr(request.form("IdCiudad")) = CStr(Rs_Ciudad_V.Fields.Item("IDCiudad").Value)) then Response.Write("SELECTED") : Response.Write("") %>>
                                                        <% =Rs_Ciudad_V.Fields.Item("Ciudad").Value %>
                                                        </option>
                                                        <%               
				                                          Rs_Ciudad_V.MoveNext()
				                                        Wend
				                                        %>
                                                </select>
                                          </td>
                                          <td><% Call ReadLang(Rs_Lenguaje,130)  %> </td>
                                          <td><%
					                             If    (Len(Request.Form("ciudadmod"))=0) Then 
  				 	                                    Response.Write (" <input type=""text"" name=""ciudadmod"" class=""formtable"" maxlength=""30"" value=""")
							                            If  Request.Form ("IdCiudad") <> 0 Then
								                             Response.Write (Rs_IDCPais_L.fields.item("ciudad").value ) 
							                            End if
  				 	                                    Response.Write (""">" )
					                             Else
							                            Response.Write ("<input type=""text"" name=""ciudadmod"" class=""formtable""  maxlength=""30"" value=""")
							                            If  Request.Form ("IdCiudad") <> 0 Then
								                             Response.Write (Rs_IDCPais_L.fields.item("ciudad").value ) 
							                            Else																			
	 							                             'Response.Write Request.Form ("ciudadmod") 
							                            End if
  				 	                                    Response.Write (""">" )
					                            End If
				                             %>
                                          </td>
                                        </tr>
                                        <tr>
                                          <td>&nbsp;</td>
                                          <td><%
					                             If    (Len(Request.Form("telefonomod"))=0) Then 
  				 	                                    Response.Write (" <input type=""hidden"" name=""telefonomod"" class=""formtable"" maxlength=""30"" value="""  )
							                            If  Request.Form ("IdCiudad") <> 0 Then
								                             Response.Write (Rs_IDCPais_L.fields.item("telefono").value ) 
							                            End if
  				 	                                    Response.Write (""">" )
					                             Else
							                            Response.Write ("<input type=""hidden"" name=""telefonomod"" class=""formtable""  maxlength=""30"" value=""" )
							                            If  Request.Form ("IdCiudad") <> 0 Then
								                             Response.Write (Rs_IDCPais_L.fields.item("telefono").value ) 
							                            Else																			
	 							                             Response.Write Request.Form ("telefonomod") 
							                            End if
  				 	                                    Response.Write (""">" )
					                            End If
				                              %>
                                          </td>
                                          <td>&nbsp;</td>
                                          <td><%
					                             If    (Len(Request.Form("internomod"))=0) Then 
  				 	                                    Response.Write (" <input type=""hidden"" name=""internomod"" class=""formtable"" maxlength=""10"" value=""")
							                            If  Request.Form ("IdCiudad") <> 0 Then
								                             Response.Write (Rs_IDCPais_L.fields.item("interno").value ) 
							                            End if
  				 	                                    Response.Write (""">" )
					                             Else
							                            Response.Write ("<input type=""hidden"" name=""internomod"" class=""formtable""  maxlength=""10"" value=""")
							                            If  Request.Form ("IdCiudad") <> 0 Then
								                             Response.Write (Rs_IDCPais_L.fields.item("interno").value ) 
							                            Else																			
	 							                             Response.Write Request.Form ("IdCiudad") 
							                            End if
  				 	                                    Response.Write (""">" )
					                            End If
				                              %>
                                          </td>
                                        </tr>
                                        <tr>
                                          <td>&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,131)  %> </td>
                                          <td>
                                          
                                            <select size="1" name="paismod" class="FormTable">
                                                <option value="0">----</option>
                                                        <% 
		                                                If  Request.Form ("IdCiudad") <> 0 Then
			                                                While (NOT RsPais.EOF)
				                                                Response.Write  ("<option value=""" )
				                                                Response.Write  (RsPais.Fields.Item("IDPais").Value)
				                                                Response.Write  (""" ")
				                                                if (CStr(Rs_IDCPais_L.Fields.Item("IDPais").Value) = CStr(RsPais.Fields.Item("IDPais").Value)) then 
				                                                  Response.Write("SELECTED") 
				                                                else 
				                                                  Response.Write("")
				                                                end if
				                                                Response.Write  (">")
				                                                Response.Write  (RsPais.Fields.Item("Pais").Value)
				                                                Response.Write  ("</option>")

				                                                RsPais.MoveNext()
			                                                Wend
		                                                Else
			                                                While (NOT RsPais.EOF)
				                                                Response.Write  ("<option value=""" )
				                                                Response.Write  (RsPais.Fields.Item("IDPais").Value)
				                                                Response.Write  (""">")
				                                                Response.Write  (RsPais.Fields.Item("Pais").Value)
				                                                Response.Write  ("</option>")
				                                                RsPais.MoveNext()
			                                                Wend
                                                				
                                                									
		                                                End If
		                                                %>
                                            </select>
                                          </td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                        </tr>
                                        <tr>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                        </tr>
                                        <tr>
                                             
                                                    <%If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if is the demo account %>
                                                            <td colspan="4">
                                                            <div align="center">                                                            
                                                     <% 	Response.Write ("<input type=""button"" name=""Enviar"" value=""")
                	                                        Call ReadLang(Rs_Lenguaje,132) 
                	                                        Response.Write (""" class=""formbutton"" onclick=""javascript: submitfrmmodif();"">")
        	                                        %>
        	                                        &nbsp;
        	                                        <% 	Response.Write ("<input type=""reset"" name=""Limpiar"" value=""")
                	                                        Call ReadLang(Rs_Lenguaje,133) 
                	                                        Response.Write (""" class=""formbutton"" >")
        	                                        %>
                                                                
                                                            </div>    </td>
                                                    <%Else %>
                                                             <td width="20%" colspan="6" style="color:red;text-align:center;font-weight:bold;">
                                                                                                       <div align="center"><% Call ReadLang(Rs_Lenguaje,205) %>
                                                                 </div></td>

    										        <%End If %>            	
                                              
                                          </div>
                                          </tr>
              </table>
                        </form>
                    
                </td>
                </tr>
                </table>  

<!--#INCLUDE file="library/PageClose.asp"-->

<%
if  Request.Form ("ciudadesmod") <> 0 Then
Rs_IDCPais_L.Close()
end if
Rs_Ciudad_V.Close()
RsPais.Close()
%>