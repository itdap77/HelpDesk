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

set Rs_Pais = Server.CreateObject("ADODB.Recordset")
Rs_Pais.ActiveConnection = MM_HelpDesk_STRING
Rs_Pais.Source = "{call dbo.SPU_Pais_VE(" & 0 & "," & 1 & ")}"
Rs_Pais.Open()

If len(vDataForm("idpais")) <> 0 then
        set Rs_PaisMod = Server.CreateObject("ADODB.Recordset")
        Rs_PaisMod.ActiveConnection = MM_HelpDesk_STRING
        Rs_PaisMod.Source = "{call dbo.[SPU_Pais_V_ID](" & 0 & "," & 1 & "," & vDataForm("idpais") & ")}"
        Rs_PaisMod.Open()
End If

%>


<script  type="text/javascript">
    var lang = getQuerystring('lang');
    var borrarpais, seleccountry,modifpais;
    var validresult;

    if (lang == 'esp')
    { borrarpais = 'Esta seguro de borrar el pais seleccionado?'; seleccountry = 'Tiene que seleccionar un pais para poder ser borrado.'; modifpais='Esta seguro que desea modificar el nombre del pais ?' } else
    { borrarpais = 'Are you sure that you want to delete the selected country ?'; seleccountry = 'You have to select a country to be able to deleted.';modifpais = 'Are you sure that you want to rename the selected country ?'; }
    
        function submitfrmalta()
        {
            
            validresult = MM_validateForm('alta_pais', '', 'R', 'sigla', '', 'R');
                        
            if (validresult == 0) {
                                        altapais.accion.value = "A";
                                        altapais.action = "vbspais.asp" ;
                                        altapais.submit();
                                        }
        }

        function submitfrmbaja()
        {
            var paisbaja = document.getElementById('paises_baja');

            if (paisbaja.value != 0) {

                if (confirm(borrarpais)) {
                    bajapais.accion.value = "B";
                    bajapais.action = "vbspais.asp";
                    bajapais.submit();
                }
            } else alert(seleccountry);

        }

        function submitfrmmodif() {
            validresult = MM_validateForm('sigla_modificacion', '', 'R', 'paismod', '', 'R', 'idpais', '', 'R');
            
        if (validresult == 0) {
            if (confirm(modifpais)) {
                modificacionpais.accion.value = "M";
                modificacionpais.action = "vbspais.asp";
                modificacionpais.submit();
            } 
        }
    }
    function reload() {
        modificacionpais.submit();
    }
</script>

      <div align="center"> 
        <center>
          <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo"><% Call ReadLang(Rs_Lenguaje,135)  %> </td>
            </tr>
          </table>
          
          <table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
            <tr> 
              <td width="100%"> 
                
                        
                        <form name="altapais" method="post" action="">
                            <input type="hidden" value="" name="accion"/>
                              <table width="100%" border="0" class="FormBoxHeader">
                                <tr>
                                  <td width="5"><img alt="" src="images/gripgray.gif" width="10" height="13"/></td>
                                  <td> 
                                    <p><% Call ReadLang(Rs_Lenguaje,136)  %> </p>
                                  </td>
                                </tr>
                              </table>
                          
                               <table width="100%" border="0" class="FormBoxBody" cellpadding="2" cellspacing="0">
                                    <tr> 
                                      <td colspan="4">&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      <td width="30%">&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,139)  %> </td>
                                      <td width="30%"> 
                                        <input type="text" name="alta_pais" class="FormTable" size="20" maxlength="20"/>
                                      </td>
                                      <td> &nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,140)  %></td>
                                      <td>
                                        <input type="text" name="sigla" size="4" maxlength="2" class="FormTable"/>
                                      </td>
                                    </tr>
                                    <tr> 
                                      <td>&nbsp;</td>
                                      <td colspan="3">&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      
                                
                                          <%If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if is the demo account %>
                                                        <td colspan="4"> 
                                                      <div align="center"> 
                                                  <% 	Response.Write ("<input type=""button"" name=""Ingresar"" value=""")
                                                            Call ReadLang(Rs_Lenguaje,141) 
                                                            Response.Write (""" class=""formbutton"" onclick=""javascript: submitfrmalta();""/>")
                                                    %>
                                                    &nbsp;
                                                    <% 	Response.Write ("<input type=""reset"" name=""Limpiar2"" value=""")
                                                            Call ReadLang(Rs_Lenguaje,142) 
                                                            Response.Write (""" class=""formbutton"" />")
                                                    %>
                                                          </div></td>

                                            <%eLSE %>
                                                <div align="center"> 
                                               <td width="20%" colspan="4" style="color:red;text-align:center;font-weight:bold;">
                                                             <% Call ReadLang(Rs_Lenguaje,206) %>
                                        </div>
                                                                 </td>
                                        
                                            <%End if %>
                                                
                                      </td>
                                    </tr>
                                  </table>
                          
                              <table width="100%" border="0">
                                <tr> 
                                  <td>&nbsp; </td>
                                </tr>
                              </table>
                        </form>
                        
                        <form name="bajapais" action="" method="post">
                                <input type="hidden" value="0" name="accion"/>
                                  <table width="100%" border="0" class="FormBoxHeader">
                                    <tr>
                                      <td width="5"><img alt="" src="images/gripgray.gif" width="10" height="13"/></td>
                                      <td> 
                                        <p><% Call ReadLang(Rs_Lenguaje,137)  %> </p>
                                      </td>
                                    </tr>
                                  </table>
                                  
                                  <table width="100%" border="0" class="FormBoxBody" cellpadding="2" cellspacing="0">
                            <tr> 
                              <td colspan="2">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td width="33%">&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,143)  %> </td>
                              <td width="33%"> 
                                <select name="paises_baja" class="FormTable">
                                  <option value="0" selected >----</option>
                                  <%
                                    While (NOT Rs_Pais.EOF)
                                    %>
                                      <option value="<%=(Rs_Pais.Fields.Item("idpais").Value)%>"><%=(Rs_Pais.Fields.Item("Pais").Value)%></option>
                                      <%
                                          Rs_Pais.MoveNext()
                                        Wend
                                        If (Rs_Pais.CursorType > 0) Then
                                          Rs_Pais.MoveFirst
                                        Else
                                          Rs_Pais.Requery
                                        End If
                                        %>
                                </select>
                              </td>
                           
                              <%If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if is the demo account %>
                              <td colspan="4"> 
                              <% 	Response.Write ("<input type=""button"" name=""Eliminar"" value=""")
                                        Call ReadLang(Rs_Lenguaje,144) 
                                        Response.Write (""" class=""formbutton"" onclick=""javascript: submitfrmbaja();""/>")
                                %>
                              </td>
                                 <%Else %>
                                <td colspan="2"></td>
                                
                                      
                            <tr><td width="20%" colspan="4" style="color:red;text-align:center;font-weight:bold;">
                                    <% Call ReadLang(Rs_Lenguaje,206) %>
                                    </td>
                                  <%End if %></tr>
                           
                          </table>
                          
                                  <table width="100%" border="0" height="2%">
                                    <tr> 
                                      <td>&nbsp; </td>
                                    </tr>
                                  </table>
                        </form>
                        
                        <form name="modificacionpais" action="" method="post">
                        <input type="hidden" value="0" name="accion"/>
                          <table width="100%" border="0" class="FormBoxHeader">
                            <tr>
                              <td width="5"> 
                                <p><img alt="" src="images/gripgray.gif" width="10" height="13"></p>
                                </td>
                              <td> 
                                <p><% Call ReadLang(Rs_Lenguaje,138)  %> </p>
                              </td>
                            </tr>
                          </table>
                          
                  <table width="100%" border="0" class="FormBoxBody" cellpadding="2" cellspacing="0">
                    <tr> 
                      <td width="40%">&nbsp;</td>
                      <td width="70%">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,145)  %> </td>
                      <td width="70%"> 
                   
                        <select  name="idpais" class="FormTable" onchange="javascript: reload();">
                                        <option value="0">----</option>
                                                        
                                <% 
							While (NOT Rs_Pais.EOF)
							%>
						        <option value="<%=(Rs_Pais.Fields.Item("IdPais").Value)%>"
							<%	if (CStr(request.form("idpais")) = CStr(Rs_Pais.Fields.Item("IdPais").Value)) then 
							Response.Write("SELECTED") 
							Else 
							Response.Write("")
							End If
							%>><%Response.Write(Rs_Pais.Fields.Item("Pais").Value )%>
						        </option>
						    <%
							Rs_Pais.MoveNext()
							Wend
										
							%>                           
				                                    
                           </select>
                        
                      </td>
                    </tr>
                    <tr> 
                      <td>&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,146)  %> </td>
                      <td width="70%"> 
                        
                                        <%
                            
                                                If  Request.Form ("idpais")  = 0 Then
                                                        Response.Write (" <input type=""text"" name=""paismod"" class=""formtable"" maxlength=""30"" value=""")
                                                        Response.Write (""">" )
                                                 Else
                                                        Response.Write (" <input type=""text"" name=""paismod"" class=""formtable"" maxlength=""30"" value=""")
                                                        Response.Write (Rs_PaisMod.Fields.Item("pais").value)
                                                        Response.Write (""">" )
                                                 End If
                                                
					                            
					                            
				                             %>
                
                      </td>
                    </tr>
                    <tr> 
                      <td>&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,147)  %> </td>
                      <td width="70%"> 
                        
                        
                        <%
                        
                        If  Request.Form ("idpais")  = 0 Then
                                Response.Write (" <input type=""text"" name=""sigla_modificacion"" class=""formtable"" maxlength=""2"" value=""")
                                Response.Write (""">" )
                            Else
                                Response.Write (" <input type=""text"" name=""sigla_modificacion"" class=""formtable"" maxlength=""2"" value=""")
                                Response.Write (Rs_PaisMod.Fields.Item("sigla").value)
                                Response.Write (""">" )
                            End If
                                            
				                             %>
                   
                      </td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td width="70%">&nbsp;</td>
                    </tr>
                    <tr> 
                      

                              <%If lcase(Session("NombreUsuario")) <> "demo" Then 'Run only if is the demo account %>
                              <td colspan="2"> 
                                                          <div align="center"> 
                              <% 	Response.Write ("<input type=""button"" name=""Ingresar"" value=""")
                                        Call ReadLang(Rs_Lenguaje,148) 
                                        Response.Write (""" class=""formbutton"" onclick=""javascript: submitfrmmodif();""/>")
                            %>
                            &nbsp;
                            <% 	Response.Write ("<input type=""reset"" name=""Limpiar"" value=""")
                                    Call ReadLang(Rs_Lenguaje,149) 
                                    Response.Write (""" class=""formbutton"" />")
                            %>
                        </div>
                                  </td>
                            <% Else %>
                                    <td width="20%" colspan="4" style="color:red;text-align:center;font-weight:bold;">
                                    <div align="center">
                                        <% Call ReadLang(Rs_Lenguaje,206) %>
                                    </div></td>
                                  <%End if %>    
                        
                      
                    </tr>
                  </table>
                          </form>
                         
              </td>
              </tr>        
          </table>
        </center>
      </div>
      <!--#INCLUDE file="library/PageClose.asp"-->
<%
Rs_Pais.Close()
%>