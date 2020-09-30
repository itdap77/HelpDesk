<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If
ValidUserAction "ABM", "ABM"

set Rs_Parte = Server.CreateObject("ADODB.Recordset")
Rs_Parte.ActiveConnection = MM_HelpDesk_STRING
Rs_Parte.Source = "{call dbo.SPU_Parte_VE2}"
Rs_Parte.Open()

%>

<script languague="javascript" language="JavaScript">
function submitfrmalta()
{
	if (altaparte.alta_parte.value == "")
	{
		alert ('El campo nombre es obligatorio.!!');
	}
	else
	{
		altaparte.action = "vbsparte.asp" ;
		altaparte.submit();
	}
}

function submitfrmbaja()
{
if (confirm("Esta Seguro de borrar la parte seleccionada?"))
{
	bajaparte.action = "vbsparte.asp" ;
	bajaparte.submit();
}
}

function submitfrmmodif()
{
modificacionparte.action = "vbsparte.asp" ;
modificacionparte.submit();

}
</script>

      <div align="center"> 
        <center>
          <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td class="Titulo">ABM de Partes del Inventario</td>
            </tr>
          </table>
          <table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
            <tr> 
              <td width="100%"> 
                
                        
                <table width="100%" border="0" cellspacing="0" cellpadding="0" height="8%">
                  <tr>
                            
                    <td>&nbsp;</td>
                          </tr>
                        </table>
                        <form name="altaparte" method="POST" action="">
                        <input name="accion" type="hidden" value="A">                          
                  <table width="100%" class="FormBoxHeader">
                    <tr> 
                              <td>
<p>Alta</p>
                              </td>
                            </tr>
                          </table>
                          
                  <table width="100%" border="0" class="FormBoxBody" cellpadding="2" cellspacing="0">
                    <tr> 
                      <td colspan="2">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;Nombre de la nueva 
                        Parte(*):</td>
                      <td> 
                        <input type="text" name="alta_parte" class="FormTable" size="20" maxlength="20">
                      </td>
                    </tr>
                    <tr> 
                      <td>&nbsp;&nbsp;&nbsp;&nbsp;Sigla(*):</td>
                      <td> 
                        <input type="text" name="sigla" size="2" class="FormTable" maxlength="2">
                      </td>
                    </tr>
                    <tr>
                      <td colspan="2">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="2"> 
                        <div align="center"> 
                          <input type="button" value="Agregar" name="Agregar" class="FormButton" onClick="javascript: submitfrmalta();">
                          <input type="reset" name="Limpiar2" value="Limpiar" class="FormButton">
                        </div>
                      </td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="2%">
                    <tr> 
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                </form>
                        <form name="bajaparte" action="" method="POST">
                        <input name="accion" type="hidden" value="B">
                          
                  <table width="100%" class="FormBoxHeader">
                    <tr> 
                              <td> 
                                <p>Baja</p>
                                </td>
                            </tr>
                          </table>
                          
                  <table width="100%" border="0" class="FormBoxBody" cellpadding="2" cellspacing="0">
                    <tr> 
                      <td colspan="3">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;Seleccione la Parte:</td>
                      <td colspan="2"> 
                        <select name="partes_baja" class="FormTable">
                          <Option selected value="----">----</option>
                          <%While (NOT Rs_Parte.EOF)%>
                          <option value="<%=(Rs_Parte.Fields.Item("Parte").Value)%>"><%=(Rs_Parte.Fields.Item("Parte").Value)%></option>
                          <%Rs_Parte.MoveNext()
								Wend
								If (Rs_Parte.CursorType > 0) Then
								  Rs_Parte.MoveFirst
								Else
								  Rs_Parte.Requery
								End If
								%>
                        </select>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="3">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="3"> 
                        <div align="center"> 
                          <input type="button" value="Eliminar" name="Eliminar" class="FormButton" onClick="javascript: submitfrmbaja();">
                        </div>
                      </td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="2%">
                    <tr> 
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                </form>
                        <form name="modificacionparte" action="" method="POST">
                          <input name="accion" type="hidden" value="M">
                  <table width="100%" class="FormBoxHeader">
                    <tr> 
                              <td> 
                                
                        <p>Modificacion</p>
                                </td>
                            </tr>
                          </table>
                          
                  <table width="100%" border="0" class="FormBoxBody" cellpadding="2" cellspacing="0">
                    <tr>
                      <td width="33%">&nbsp;</td>
                      <td width="70%">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;Seleccione la Parte:</td>
                      <td width="70%"> 
                        <select name="partes_modificacion" class="FormTable">
                          <Option selected value="----">----</option>
                          <%While (NOT Rs_Parte.EOF)%>
                          <option value="<%=(Rs_Parte.Fields.Item("IDParte").Value)%>"><%=(Rs_Parte.Fields.Item("Parte").Value)%></option>
                          <%Rs_Parte.MoveNext()
								Wend
								If (Rs_Parte.CursorType > 0) Then
								  Rs_Parte.MoveFirst
								Else
								  Rs_Parte.Requery
								End If
								%>
                        </select>
                      </td>
                    </tr>
                    <tr> 
                      <td>&nbsp;&nbsp;&nbsp;&nbsp;Nombre de la nueva Parte:</td>
                      <td width="70%"> 
                        <input type="text" name="modificacion_parte" class="FormTable" size="20" maxlength="20">
                      </td>
                    </tr>
                    <tr> 
                      <td>&nbsp;&nbsp;&nbsp;&nbsp;Sigla(*):</td>
                      <td width="70%"> 
                        <input type="text" name="sigla_modificacion" size="2" class="FormTable" maxlength="2">
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="2">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="2"> 
                        <div align="center"> 
                          <input type="button" value="Modificar" name="Modificar" class="FormButton" onClick="javascript: submitfrmmodif();">
                          <input type="reset" name="Limpiar" value="Limpiar" class="FormButton">
                        </div>
                      </td>
                    </tr>
                  </table>
                          </form>
                      
          </table>
        </center>
      </div>
      <!--#INCLUDE file="library/PageClose.asp"-->
<%Rs_Parte.Close()%>