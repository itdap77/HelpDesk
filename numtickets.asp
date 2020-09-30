<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="library/pageopen.asp" -->

<%
Validsession()

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_HelpDesk_STRING
rs.Source = "{call dbo.SPU_Ticket_Max}"
rs.Open()

%>
      <table width="640px" border="0" class="Titulo">
        <tr> 
          <td><% Call ReadLang(Rs_Lenguaje,56) %></td>
        </tr>
      </table>
      <table class="ContentArea" width="100%" align="center" cellpadding="0" cellspacing="0" height="100%">
        <tr> 
          <td valign="middle" height="5%">&nbsp;</td>
        </tr>
        <tr> 
          <td valign="middle"> 
            <div align="center"><font size="3"> </font> 
              <table width="50%" border="0" cellpadding="0" cellspacing="0" align="center" class="LogInBox">
                <tr> 
                  <td><img  alt="" src="images/borde11.gif" width="8" height="8"/></td>
                  <td width="100%">&nbsp;</td>
                  <td><img  alt="" src="images/borde12.gif" width="8" height="8" valign="top"/></td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td> 
                    <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <% Call ReadLang(Rs_Lenguaje,57) %> 
                      </div>
                  </td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td>
                    <div class="numticket" align="center"> 
                      <font face="Verdana, Arial, Helvetica, sans-serif" size="2" ><b>000-000<%=(rs.Fields.Item("maximo").Value)%></b></font> <br>
                      </div>
                      <center><img alt="" src="images/barras.jpg" width="167" height="63" align="center" alt="" /></center>
                    
                  </td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td height="1%" style="VERTICAL-ALIGN: bottom"><img src="images/borde21.gif" v-align="bottom" width="8" height="8"  alt="" /></td>
                  <td>&nbsp;</td>
                  <td style="VERTICAL-ALIGN: bottom"><img src="images/borde22.gif" valign="bottom" width="8" height="8"  alt="" /></td>
                </tr>
              </table>
              <% 
                
                Response.Write ("<br><a href=detalletickets.asp?Accion=P&lang=")
                Response.Write (vDataForm("lang") )
                Response.Write ("&IDTicket=")
                Response.Write (rs.Fields.Item("maximo").Value)
                Response.Write (" class=""AccionBlack"">")
                Call ReadLang(Rs_Lenguaje,196) 
                Response.Write ("</a>")  
                %>
              
            </div>
          </td>
        </tr>
        <tr>
          <td valign="middle" height="33%">&nbsp;</td>
        </tr>
      </table>

<%
rs.Close()
%>
<!--#include file="library/pageclose.asp" -->