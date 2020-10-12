<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
      
<%

ValidSession()

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

dim ST, ST2,cont

Set Rs_detalle = Server.CreateObject("ADODB.Recordset")
Rs_detalle.ActiveConnection = MM_HelpDesk_STRING
ST = "{call dbo.SPU_TicketDetalle_V ("  + vDataForm("IDTicket")  +  ")}"
Rs_detalle.Source =  ST
Rs_detalle.Open()

If Rs_detalle.EOF then
    If vDataForm("lang") = "esp" or session("ididioma") = 1 Then
	    Response.redirect ("mensaje.asp?tipomsg=1&msg=No se encontro ningun evento.&BackToURL=vertickets.asp")
    Else
        Response.redirect ("mensaje.asp?tipomsg=1&msg=No event found.&BackToURL=vertickets.asp")
    End if
End If

    Set Rs_evento = Server.CreateObject("ADODB.Recordset")
    Rs_evento.ActiveConnection = MM_HelpDesk_STRING
    ST2 = "{call dbo.SPU_TicketEvento_V ("  + vDataForm ("IDTicket")  +  ")}"
    Rs_evento.Source =  ST2
    Rs_evento.Open()

If Rs_evento.EOF then
    if session("idioma") = 1 then
	Response.redirect ("mensaje.asp?tipomsg=1&msg=No se encontro ningun evento&BackToURL=vertickets.asp")
	Else
	Response.redirect ("mensaje.asp?tipomsg=1&msg=No event found.&BackToURL=vertickets.asp")
	End If
End if

	while not rs_evento.EOF 
		cont = cont + 1
		rs_evento.MoveNext
	wend
		rs_evento.MoveFirst
             

set Rs_Problema_ID = Server.CreateObject("ADODB.Recordset")
Rs_Problema_ID.ActiveConnection = MM_HelpDesk_STRING
Texto = Rs_detalle ("IDProblema")

Resp1 = mid(Texto,1,instr(Texto,",")-1)
Rs_Problema_ID.Source = "{call dbo.SPU_Problema_ID2(" + Resp1 + ")}"
Rs_Problema_ID.Open()
ResultProblema = Rs_Problema_ID.Fields.Item("DetalleProblema").Value 
Rs_Problema_ID.close()

texto = mid (texto,instr(Texto,",")+1,len(Texto))
Resp2 = mid(Texto,1,instr(Texto,",")-1)

Rs_Problema_ID.Source = "{call dbo.SPU_Problema_ID2(" + Resp2 + ")}"
Rs_Problema_ID.Open()
ResultProblema = ResultProblema + " - " + Rs_Problema_ID.Fields.Item("DetalleProblema").Value 
Rs_Problema_ID.Close()

resp3 = mid (texto,instr(Texto,",")+1,len(Texto))

Rs_Problema_ID.Source = "{call dbo.SPU_Problema_ID2(" + Resp3 + ")}"
Rs_Problema_ID.Open()
ResultProblema = ResultProblema + " - " + Rs_Problema_ID.Fields.Item("DetalleProblema").Value 
Rs_Problema_ID.close()


%>


<script type="text/javascript">
	document.detalletickets = 'detalletickets.asp';
	document.HelpSection = 'detalletickets';

function openwin2 ()
{
	frmobs.submit();
	window.open('detalleobs.asp', '_blank', 'width=300,height=300,resizable=no,scrollbars=no,toolbar=no,menubar=no,location=no,directories=no,status=no,titlebar=no');
}
</script>

      <table border="0" class="TopMenuArea" cellpadding="0" cellspacing="0">
        <tr> 
          <td class="titulo"><% Call ReadLang(Rs_Lenguaje,70) %></td>
          
				
          
        </tr>
      </table>
            <table width="100%" border="0" class="ContentArea" height="100%" cellspacing="0" cellpadding="0">
        <tr> 
                
          <td> 
     <table class="TopActionMenuTable">
     <tr>
     	<% if Rs_detalle.Fields.Item("Estado").value <> "Cerrado" Then
     	           
			'Reclamar para todos los usuarios
					Response.Write ("<td>")'test
                    Response.Write ("<table style=""border:0px solid;padding:0 auto 0 auto;margin:0;text-align:center;width:100%""><tr>") 'test		

				    ' solo para administradores y solucionadores
                    If (Session("IDGrupo") = 4 or Session("IDGrupo") = 1) Then
                    
                      'Check In
     	                    Response.write("<a href=""#""  ID=""mnuecento"" onclick=""javascript: openwin('evento.asp?IDTicket=")
					        Response.write(vDataForm("IDTicket"))
					        Response.Write("&Title=")
					        call ReadLang(Rs_Lenguaje,183)
					        Response.Write("&Accion=I")
                            Response.Write("&Problema=" & ResultProblema)                            
					        Response.Write("&LoginDeriv=" + Rs_evento.Fields.Item("loginderiv").value)
					        Response.Write("&IDUsuarioDeriv=" + cstr(Rs_evento.Fields.Item("IDUsuarioDeriv").value))
					        Response.write("','_blank','width=400,height=150,resizable=no,scrollbars=no,toolbar=no,menubar=no,location=no,directories=no,status=no,titlebar=no');"">")
					        Response.write("<td width=""80"" class=""TopActionMenu"">")
					        call ReadLang(Rs_Lenguaje,183)
				            Response.Write("</td></a>")
				    
				            'Check Out
     	                    Response.write("<a href=""#""  ID=""mnuecento"" onclick=""javascript: openwin('evento.asp?IDTicket=")
					        Response.write(vDataForm("IDTicket"))
					        Response.Write("&Title=")
					        call ReadLang(Rs_Lenguaje,184)
					        Response.Write("&Accion=O")
                            Response.Write("&Problema=" & ResultProblema)
					        Response.Write("&LoginDeriv=" + Rs_evento.Fields.Item("loginderiv").value)
					        Response.Write("&IDUsuarioDeriv=" + cstr(Rs_evento.Fields.Item("IDUsuarioDeriv").value))
					        Response.write("','_blank','width=400,height=150,resizable=no,scrollbars=no,toolbar=no,menubar=no,location=no,directories=no,status=no,titlebar=no');"">")
					        Response.write("<td width=""80"" class=""TopActionMenu"">")
					        call ReadLang(Rs_Lenguaje,184)
				            Response.Write("</td></a>")

                            Response.write("<a href=""#""  ID=""mnuecento"" onclick=""javascript: openwin('evento.asp?IDTicket=")
					        Response.write(vDataForm("IDTicket"))
					        Response.Write("&Title=")
					        call ReadLang(Rs_Lenguaje,72)
					        Response.Write("&Accion=S")
                            Response.Write("&Problema=" & ResultProblema)
					        Response.Write("&LoginDeriv=" + Rs_evento.Fields.Item("loginderiv").value)
					        Response.Write("&IDUsuarioDeriv=" + cstr(Rs_evento.Fields.Item("IDUsuarioDeriv").value))
					        Response.write("','_blank','width=400,height=150,resizable=no,scrollbars=no,toolbar=no,menubar=no,location=no,directories=no,status=no,titlebar=no');"">")
					        Response.write("<td width=""80"" class=""TopActionMenu"">")
					        call ReadLang(Rs_Lenguaje,72)
				            Response.Write("</td></a>")


					End if 

										'Assign solo para administradores
				    If (Session("IDGrupo") <> 10) Then
				    
				            Response.write("<a href=""#""  ID=""mnuecento"" onclick=""javascript: openwin('evento.asp?IDTicket=")
					        Response.write(vDataForm("IDTicket"))
					        Response.Write("&Title=")
					        call ReadLang(Rs_Lenguaje,71)
					        Response.Write("&Accion=D")
                            Response.Write("&Problema=" & ResultProblema)
					        Response.write("','_blank','width=400,height=150,resizable=no,scrollbars=no,toolbar=no,menubar=no,location=no,directories=no,status=no,titlebar=no');"">")
					        Response.write("<td width=""80"" class=""TopActionMenu"">")
					        call ReadLang(Rs_Lenguaje,71)
				            Response.Write("</td></a>")

				    End If
				
					

					Response.write("<a href=""#""  ID=""mnuecento"" onclick=""javascript: openwin('evento.asp?IDTicket=")
			        Response.write(vDataForm("IDTicket"))
			        Response.Write("&Title=")
			        call ReadLang(Rs_Lenguaje,73)
			        Response.Write("&Accion=R")
                    Response.Write("&Problema=" & ResultProblema)
			        Response.Write("&LoginDeriv=" + Rs_evento.Fields.Item("loginderiv").value)
			        Response.Write("&IDUsuarioDeriv=" + cstr(Rs_evento.Fields.Item("IDUsuarioDeriv").value))
			        Response.write("','_blank','width=400,height=150,resizable=no,scrollbars=no,toolbar=no,menubar=no,location=no,directories=no,status=no,titlebar=no');"">")
			        Response.write("<td width=""80"" class=""TopActionMenu"">")
			        call ReadLang(Rs_Lenguaje,73)
		            Response.Write("</td></a>")
	
	                'On hold
					Response.write("<a href=""#""  ID=""mnuecento"" onclick=""javascript: openwin('evento.asp?IDTicket=")
			        Response.write(vDataForm("IDTicket"))
			        Response.Write("&Title=")
			        call ReadLang(Rs_Lenguaje,195)
			        Response.Write("&Accion=H")
                    Response.Write("&Problema=" & ResultProblema)
			        Response.Write("&LoginDeriv=" + Rs_evento.Fields.Item("loginderiv").value)
			        Response.Write("&IDUsuarioDeriv=" + cstr(Rs_evento.Fields.Item("IDUsuarioDeriv").value))
			        Response.write("','_blank','width=400,height=150,resizable=no,scrollbars=no,toolbar=no,menubar=no,location=no,directories=no,status=no,titlebar=no');"">")
			        Response.write("<td width=""80"" class=""TopActionMenu"">")
			        call ReadLang(Rs_Lenguaje,195)
		            Response.Write("</td></a>")
		            
             
                    'Cerrar
					Response.write("<a href=""#""  ID=""mnuecento"" onclick=""javascript: openwin('evento.asp?IDTicket=")
			        Response.write(vDataForm("IDTicket"))
			        Response.Write("&Title=")
			        call ReadLang(Rs_Lenguaje,74)
			        Response.Write("&Accion=C")
                    Response.Write("&Problema=" & ResultProblema)
			        Response.Write("&LoginDeriv=" + Rs_evento.Fields.Item("loginderiv").value)
			        Response.Write("&IDUsuarioDeriv=" + cstr(Rs_evento.Fields.Item("IDUsuarioDeriv").value))
			        Response.write("','_blank','width=400,height=150,resizable=no,scrollbars=no,toolbar=no,menubar=no,location=no,directories=no,status=no,titlebar=no');"">")
			        Response.write("<td width=""80"" class=""TopActionMenu"">")
			        call ReadLang(Rs_Lenguaje,74)
		            Response.Write("</td></a>")
                    
                    If (Session("IDGrupo") = 10) Then
                            
                            Response.write ("<td class=""TopActionMenuEmpty"">&nbsp;</td>")
                     End If
                    
                     Response.Write("</tr></table></td>") 'test   
				 End If 
				 %>
     </tr>
     </table>

              <!-- Header table -->
            <table width="100%" border="0" class="FormBoxHeader"> 
              <tr> 
                <td width="10px"><img alt="" src="images/gripgray.gif" width="10" height="13"/></td>
                  <td width="115px"><b><% Call ReadLang(Rs_Lenguaje,197) %></b></td>
                  <td> <% =vDataForm("IDTicket") %></td>
                   
                <td width="80px" colspan="2"> &nbsp; 
                  <%   if vDataForm("lang") = "esp" then
                      Response.Write (Rs_evento.Fields.Item("Estado").value )
                      else
                     Response.Write (Rs_evento.Fields.Item("State").value )
                      end if %>
                </td>
              </tr>
            </table>


              <!-- Main table -->
            <table width="100%" border="0" name="header" class="FormBoxBody" cellpadding="0" cellspacing="0">
              <tr> 
                <td > 
                  <div align="left"><b><% Call ReadLang(Rs_Lenguaje,192) %></b></div>
                </td>
                <td colspan="5"> &nbsp; 
                  <% =Rs_detalle.Fields.Item("login").value %>
                </td>
               
                                     
              </tr>
                
 
              <tr> 
                <td > <b><% Call ReadLang(Rs_Lenguaje,190) %></b></td>
                <td colspan="5"> &nbsp; 
                  <% =Rs_detalle.Fields.Item("FechaTicket").value %>
                </td>
                
              
              </tr>
              <tr> 
                <td><b><% Call ReadLang(Rs_Lenguaje,191) %></b></td>
                <td colspan="5"> &nbsp;
                  <%Response.Write (ResultProblema)%>
                </td>
              
               </tr>

                <tr>
                    <td>
                        <div align="left"><b><% Call ReadLang(Rs_Lenguaje,194) %></b></div> <!-- Observaciones -->
                    </td>
                    <td colspan="5">
                        &nbsp;
                        <% =Rs_detalle.Fields.Item("observaciones").value %>
                      
                    </td>
              </tr>
            </table>
            

            <table border="0" name="body" class="FormBoxBody" cellspacing="0" cellpadding="0">
              <tr class="RowFormBoxBody" >
                  <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% Call ReadLang(Rs_Lenguaje,185) %></td>
                <td><% Call ReadLang(Rs_Lenguaje,186) %></td>
                <td><% Call ReadLang(Rs_Lenguaje,187) %></td>
                <td><% Call ReadLang(Rs_Lenguaje,188) %></td>
                <td colspan="3" > 
                  <div align="center"><% Call ReadLang(Rs_Lenguaje,189) %></div>
                </td></tr>
               <% While not rs_evento.EOF %>         
              <tr <%if flag = 1 then
						Response.Write ("Class=""FormBoxListOddRow""")
						flag = 0
					  else
						Response.Write ("Class=""FormBoxListEvenRow""")
						flag = 1
					  end if%>>
                	<td> &nbsp;
                  <% =Rs_evento.Fields.Item("FechaEvento").value %>
                </td>
                <td> 
                  <% 
                      if vDataForm("lang") = "esp" then
                      Response.Write (Rs_evento.Fields.Item("Estado").value )
                      else
                     Response.Write (Rs_evento.Fields.Item("State").value )
                      end if
                      
                      %>
                </td>
			
                <td> 
                  <% =Rs_evento.Fields.Item("loginderiv").value %>
                </td>
                <td> 
                  <% =Rs_evento.Fields.Item("Descripcion").value %>
                </td>
                <td colspan="3" align="center"> 
                  <% 
					If (Rs_evento.Fields.Item("Observaciones").value <> "") then
   						Response.Write ("<a href=""javascript:openwin('detalleobs.asp?Observaciones=")
						Response.Write (Rs_evento.Fields.Item("Observaciones").value) 
						Response.Write ("','_blank','width=400,height=160,resizable=no,scrollbars=no,toolbar=no,menubar=no,location=no,directories=no,status=no,titlebar=no');"">")
						Response.Write ("<img src=""images/lupa.gif"" width=""22"" border =""0"" height=""17"" ")
						Response.Write ("></a>")
					Else
						Response.Write ("&nbsp;")
					End if
                %>
                </td>
                 </tr>
              	<%  
				rs_evento.MoveNext ()
				Wend 
				%>
            </table>

          </td>
        </tr>
      </table>

<!--#include file="library/pageClose.asp" -->
<% 

Rs_detalle.Close ()
Rs_evento.Close ()
%>