<%@LANGUAGE="VBSCRIPT"%>
<!--#INCLUDE file="connections/helpdesk.asp"-->
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

set Rs_Ciudad = Server.CreateObject("ADODB.Recordset")
Rs_Ciudad.ActiveConnection = MM_HelpDesk_STRING
Rs_Ciudad.Source = "{call dbo.SPU_Ciudad_VE}"
Rs_Ciudad.Open()

If Request.Form ("pais") <> 0 Then
set Rs_Pais = Server.CreateObject("ADODB.Recordset")
Rs_Pais.ActiveConnection = MM_HelpDesk_STRING
Rs_Pais.Source = "{call dbo.SPU_IDPCiudad_L("+Replace(Request.Form("pais"),"'","''")+")}"
Rs_Pais.Open()
End if
%>
<html>
<head>
<title>Selecion de PC</title>
<script  Languague="JavaScript">
function submitform()
{ if (document.frmseleccionopc.idpc.value != "") 
	 { window.opener.frmingreso.idpc.value = frmseleccionopc.idpc.value;
	   window.close()
	   }
  else {frmseleccionopc.submit();}

}


</script>
<link rel="stylesheet" href="library/FormArea.css" type="text/css">
</head>
<body bgcolor="#FFFFFF">
<p align="center"><b></b></p>
<div align="center"> 
  <center>
    <form name="frmseleccionopc" method="POST">
      <table width="100%" border="0" class="FormBoxHeader">
        <tr> 
          <td width="5"><img src="images/gripgray.gif" width="10" height="13"></td>
          <td><font size="2" face="Verdana">Seleccione la Ciudad y su Correspondiente 
            PC</font></td>
        </tr>
      </table>
      <table border="0" cellspacing="0" cellpadding="0" class="FormBoxBody" width="100%">
        <tr> 
        <td width="100%"> 
          <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="100%">&nbsp;</td>
            </tr>
          </table>
          <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="100%"> 
                <table border="0" width="100%" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="47%"> 
                      <table border="0" width="100%" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="100%"><font color="#000000" size="2" face="Verdana"><b> 
                            </b></font></td>
                        </tr>
                        <tr> 
                          <td width="100%"><font color="#000000" size="2" face="Verdana"><b>Seleccione 
                            la Ciudad:</b></font></td>
                        </tr>
                        <tr> 
                          <td width="100%">&nbsp;</td>
                        </tr>
                        <tr> 
                          <td width="100%">&nbsp;</td>
                        </tr>
                      </table>
                    </td>
                    <td width="53%"> 
                      <select size="4" name="pais" onclick="javascript:submitform();">
                        <%
						While (NOT Rs_Ciudad.EOF)
						%>
                        <option value="<%=(Rs_Ciudad.Fields.Item("IDCiudad").Value)%>"<%if (CStr(request.form("pais")) = CStr(Rs_Ciudad.Fields.Item("IDCiudad").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(Rs_Ciudad.Fields.Item("Ciudad").Value)%></option>
                        <%
						  Rs_Ciudad.MoveNext()
						Wend
						If (Rs_Ciudad.CursorType > 0) Then
						  Rs_Ciudad.MoveFirst
						Else
						  Rs_Ciudad.Requery
						End If
						%>
					  </select>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
          <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="100%">&nbsp;</td>
            </tr>
          </table>
          <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="100%"> 
                <table border="0" width="100%" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="38%"> 
                        <table border="0" width="100%" cellspacing="0" cellpadding="0" class="FormTable">
                          <tr> 
                          <td width="100%"><font color="#000000" size="2" face="Verdana"> 
                            </font></td>
                        </tr>
                        <tr> 
                          <td width="100%"><font color="#000000" size="2" face="Verdana"><b>Seleccione 
                            la PC:</b></font></td>
                        </tr>
                        <tr> 
                          <td width="100%">&nbsp;</td>
                        </tr>
                        <tr> 
                          <td width="100%">&nbsp;</td>
                        </tr>
                      </table>
                    </td>
                    <td width="62%"> 
                      <select size="4" name="idpc">
						<%
						Response.Write (Request.Form("pais") )
						If Request.Form("pais")<>0  Then

							While (NOT Rs_Pais.EOF)
									
								Response.Write (" <option value=""" )
								Response.Write (Rs_Pais.Fields.Item("IDPc").Value & """> ")
								Response.Write (Rs_Pais.Fields.Item("IDPc").Value )
								Response.Write ("</option>")
																	
							    Rs_Pais.MoveNext()
							Wend
							If (Rs_Pais.CursorType > 0) Then
							 Rs_Pais.MoveFirst
							Else
							 Rs_Pais.Requery
							End If
						End If
						%>
                      </select>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
          <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="100%">&nbsp;</td>
            </tr>
          </table>
          <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="100%">
                  <div align="center">&nbsp; 
                    <input type="button" value="Cerrar" name="B1" onclick="submitform();" class="FormButton">
                  </div>
                </td>
            </tr>
          </table>
          <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="100%">&nbsp;</td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
</form>
  </center>
</div>
</body>
</html>
<%
Rs_Ciudad.Close()
If Request.Form ("pais") <> 0 Then
Rs_Pais.Close()
End if
%>

