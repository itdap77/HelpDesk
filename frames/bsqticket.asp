<%@LANGUAGE="VBSCRIPT"%>
<!--#INCLUDE file="../connections/HelpDesk.asp" -->
<!--#INCLUDE file="../library/Funciones.asp"-->
<script src="../library/Script.js"></script>

<%

Dim vDataForm,total',vPageNumber
vDataForm = NULL
If Not Len(Request.Form) = 0 Then
		Set vDataForm = Request.Form
Else
		Set vDataForm = Request.QueryString
End If

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.ActiveConnection = MM_HelpDesk_STRING
oRs.Source = "{call dbo.SPU_Productos_F('" + Replace(vDataForm("busqueda") , "'", "''") + "')}"
oRS.CursorType = 0
oRS.CursorLocation = 2
oRS.LockType = 3
oRs.Open()
oRs_numRows = 0

'Varialbe de Repeticion de datos
Repeat1__index = 0
oRs_numRows = oRs_numRows + vPageSize


If Not oRs.Eof and Not oRs.BOF then
  Total = 0
  While (Not oRS.EOF)
   Total = Total + 1
   oRS.MoveNext
  Wend
oRs.MoveFirst
End If
%>

<!--#include file="../library/paginas.asp" -->
<script Languaje="JavaScript">

	function buscar()
	{
		document.frmbuscar.submit();
	}

</script>
<html>
<head>
<link rel="stylesheet" href="../library/FormArea.css" type="text/css">
<link rel="stylesheet" href="mtpopup.css" type="text/css">
</head>
<body BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#990000" VLINK="#990000" ALINK="#990000">
<form name="frmbuscar" id="frmbuscar"method="post" action="bsqticket.asp" >
  <input type="hidden" name="DataField" value="<%=Request.Form("DataField")%>">
<input type="hidden" name="FrameName" value="<%=Request.Form("FrameName")%>">
  	<table border="0" cellspacing="0" cellpadding="0" class="PopupTable">
    <tr><td> 
        <table border="0" cellspacing="0" cellpadding="0" class="PopupTitle">
          <tr> 
            <td> BuscarTicket</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td> 
        <table border="0" cellspacing="0" cellpadding="0" class="PopupAction">
          <tr> 
            <td> Producto a buscar:&nbsp; 
              <input type="text" size="20" name="busqueda" value="<%Response.Write(vDataForm("busqueda"))%>" class="PopupInputText">
            </td>
            <td> <img src="arrow_rg_gray.gif" align="middle" style="{cursor:hand;}" onClick="buscar()" width="18" height="18"> 
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td> <!-- ABRO TABLA POPUPS -->
        <table border="0" cellspacing="0" cellpadding="0" class="PopupResults"> 
        
	  <%
		dim vHaveResults
		vHaveResults = True
		
		If Not oRS.State = 1 Then
			vHaveResults = False
		Else
			If oRS.EOF Then    
				vHaveResults = False
			End If
		End If
		
		If vHaveResults = False Then
			Response.Write(" <tr><td style=""color:red;"">") & vbCrLf
			Response.Write("  No se encontraron registros") & vbCrLf
			Response.Write(" </td></tr>") & vbCrLf
		Else
		  
		  If Not oRs.EOF OR Not oRs.BOF Then  'Encabezado de datos de lista
			  Response.Write ("<tr><td>") & vbCrLf
	  		  Response.Write ("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""FormboxHeader"" > ") & vbCrLf
			  Response.Write ("<tr>") & vbCrLf
			  Response.Write ("<td width=""45"">Nro.</td>") & vbCrLf
			  Response.Write ("<td width=""155"">Producto</td>") & vbCrLf
			  Response.Write ("<td width=""100"">Marca</td>") & vbCrLf
			  Response.Write ("<td width=""100"">Descripción</td>") & vbCrLf
			  Response.Write ("</tr>") & vbCrLf
			  Response.Write ("</table>")
		  End IF	'CIERRO IF DE Encabezado de datos de lista
			
			'TABLA BODY DE LISTA
			  Response.Write ("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""Formboxbody"" > ") & vbCrLf
    	  	While ((vPageSize <> 0) AND (NOT oRs.EOF))  'Area de datos de lista
				  Response.Write ("<tr ")	'COLORES DE CELDAS PAR O IMPAR 
				  dim flag
					  if flag = 1 then
						Response.Write ("Class=""FormBoxListEvenRow""")
						flag = 0
					  else
						Response.Write ("Class=""FormBoxListOddRow""")
						flag = 1
					  end if
					Response.Write(">") & vbCrLf
					Response.Write("	<td width=""45""><a href=""../frmproductos.asp?idproducto=""" & oRs.Fields.Item("IdProducto").value & """>" & oRs.Fields.Item("IdProducto").value & "</td>" & vbCrLf)
					Response.Write("	<td width=""155"">" & Mid(oRs.Fields.Item("Producto").value, 1, 28) & "</td>") & vbCrLf
					Response.Write("	<td width=""100"">" & Mid(oRs.Fields.Item("Marca").value, 1, 25) & "</td>") & vbCrLf
					Response.Write("	<td width=""100"">" & Mid(oRs.Fields.Item("Descripcion").value, 1, 20) & "</td>") & vbCrLf
					Response.Write ("</tr>") & vbCrLf
					
					Repeat1__index=Repeat1__index+1
					 vPageSize=vPageSize-1
					 oRs.MoveNext()
			 Wend 'CIERRO Area de datos de lista
				 
			   If (vPageSize > 1) or (vPageSize < 4) then
				   While (vPageSize <> 0)	
						Response.Write ("<tr ")	 'COLORES DE CELDAS PAR O IMPAR 
						if flag = 1 then
							Response.Write ("Class=""FormBoxListEvenRow""")
							flag = 0
						else
							Response.Write ("Class=""FormBoxListOddRow""")
							flag = 1
						end if
						Response.Write(">") & vbCrLf			
						Response.Write("<td colspan=4>&nbsp;</td>" & vbCrLf & "</tr>" & vbCrLf )
						vPageSize=vPageSize-1
				   Wend
			  End If
	
			Response.Write ("</table>") 'Cierro tabla BODY
			Response.Write ("</tr></td>") & vbCrLf
End If 'CIERRO IF DE VHAVERESULTS

%>

 </table>

 <table BORDER="0" CELLSPACING="0" CELLPADDING="0" class="PopupFooter">

          <tr> 
            <td> &nbsp;&nbsp; 
            
              <% 
			'	If vPageNumber > 1 Then
					Response.Write("<A HREF=""" & MM_movePrev & """>Anterior</a>")
			'	Else
			'		Response.Write("Anterior")
			 '	End If
			  %>
              &nbsp;|&nbsp;
              
              <%
			'	If (vPageNumber > 0  and vPageNumber < Total )  Then
					Response.Write("<A HREF=""" & MM_moveNext & """>Siguiente</a>")
			'	Else
			 '		Response.Write("Siguiente")
			'	End If
			  %>&nbsp; | &nbsp; <font color="#000000">Registros Encontrados: <%=Total%> </font> &nbsp; |  
            </td>
           <td align="right"> <!--<span onClick="frmbuscar.CloseFrame();" alt="Cerrar búsqueda">Cancelar</span>&nbsp;&nbsp; -->
           

            </td>
          </tr>
        </table>
   </table>
    </form>
<% oRs.Close()%>
<script>
<!--

var oDataField
var oFrameName

frmbuscar.busqueda.select();

function Initnombre()
{
	frmbuscar.DataField.value = oDataField;
	frmbuscar.FrameName.value = oFrameName;
		
	frmbuscar.busqueda.focus();
	frmbuscar.busqueda.select();
}

//-->
</script>
</body>
</html>
