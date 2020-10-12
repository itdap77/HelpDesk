<%@ LANGUAGE=VBScript %>

<!--#INCLUDE file="library/Funciones.asp"-->
<!--#INCLUDE file="connections/helpdesk.asp"-->



<script type="text/vbscript" language="VBScript" runat="Server">
ValidSession ()

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

dim ST, Cadena, IDTicket, Str, IdFinal


   Set oSQLConn = Server.CreateObject("ADODB.Connection")
oSQLConn.ConnectionString = MM_HelpDesk_STRING
oSQLConn.Open()



set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_HelpDesk_STRING

set Rs2 = Server.CreateObject("ADODB.Recordset")
Rs2.ActiveConnection = MM_HelpDesk_STRING

Str = "{call dbo.SPU_Login_ID2 ('" + replace(vDataForm("login"), "'", "''") + "')}"

Rs.Source = Str
'Rs.Open Str, oSQLConn
    Rs.Open()


Cadena = "'" & replace(request.form("categoria1"), "'", "''") & "," & replace(request.form("categoria2"), "'", "''") & "," & replace(request.form("categoria3"), "'", "''") & "'"

ST = "{call dbo.SPU_Ticket_A (" + _
			"NULL" + "," + _ 
			Cadena + "," + _
			replace(Session("IDU"), "'", "''") + ",'" + _		
			replace(vDataForm("login"), "'", "''") + "','" + _	
			replace(vDataForm("Observaciones"), "'", "''") + "','" + _	
			replace(Rs.Fields.Item("IDUsuario").value, "'", "''") + "','" + _
			replace(vDataForm("IDPc"), "'", "''") + "'," + _
			replace(vDataForm("IDInventario"), "'", "''") + _
			")}"
			
Rs2.Source = ST
Rs2.Open()


' ----------------------------- language section -----------------------------
Set Rs_Lenguaje_Mail = Server.CreateObject("ADODB.Recordset")
Rs_Lenguaje_Mail.ActiveConnection = MM_HelpDesk_STRING
Rs_Lenguaje_Mail.Source = "{call dbo.SPU_Lenguaje_V(" + cstr(Session("IDidioma")) + ",'" + GetPageName() + "')}"
Rs_Lenguaje_Mail.Open()
' ----------------------------- End language section -----------------------------

'Comienzo de recordset para mandar mail

Set Rs_Usuarios = Server.CreateObject("ADODB.Recordset")
Rs_Usuarios.ActiveConnection = MM_HelpDesk_STRING
Rs_Usuarios.Source = "{call dbo.SPU_Usuarios_L_Full(" + Replace(Session("IDU"), "'", "''") + ")}"
Rs_Usuarios.Open()

Set RsHd= Server.CreateObject("ADODB.Recordset")
RsHd.ActiveConnection = MM_HelpDesk_STRING
RsHd.Source = "{call dbo.SPU_HelpDesk_X}"
RsHd.Open()

Set RsMax= Server.CreateObject("ADODB.Recordset")
RsMax.ActiveConnection = MM_HelpDesk_STRING
RsMax.Source = "{call dbo.SPU_Ticket_Max}"
RsMax.Open()

Set RsProblema1= Server.CreateObject("ADODB.Recordset")
RsProblema1.ActiveConnection = MM_HelpDesk_STRING
RsProblema1.Source = "{call dbo.SPU_Problema_ID(" + Replace(request.form("categoria1"), "'", "''") + ")}"
RsProblema1.Open()

Set RsProblema2= Server.CreateObject("ADODB.Recordset")
RsProblema2.ActiveConnection = MM_HelpDesk_STRING
RsProblema2.Source = "{call dbo.SPU_Problema_ID(" + Replace(request.form("categoria2"), "'", "''") + ")}"
RsProblema2.Open()

Set RsProblema3= Server.CreateObject("ADODB.Recordset")
RsProblema3.ActiveConnection = MM_HelpDesk_STRING
RsProblema3.Source = "{call dbo.SPU_Problema_ID(" + Replace(request.form("categoria3"), "'", "''") + ")}"
RsProblema3.Open()


' ----------------------------- Email sent section -----------------------------
Dim From,FromAddress,Subject,HTMLBody,HTMLBody2,HTMLBody3,HTMLBody4,ToName,ToAddress,Attach,CC,CCAddress,BBC, BBCAddress,ToName2,ToAddress2,From2,From2Address

' Mail al helpdesk
From = replace(Rs_Usuarios.Fields.Item("Apellido").Value, "'", "''") & " " & replace(Rs_Usuarios.Fields.Item("Nombre").Value, "'", "''")
FromAddress = "helpdesk@itdap.com"
ToName = "HelpDesk | ITDap Worldwide Solutions"
ToAddress = "helpdesk@itdap.com"
ReplyTo = replace(Rs_Usuarios.Fields.Item("Apellido").Value, "'", "''") & " " & replace(Rs_Usuarios.Fields.Item("Nombre").Value, "'", "''")
ReplyToAddress = replace(Rs_Usuarios.Fields.Item("Mail").Value, "'", "''")


' Mail al usuario
From2 = "HelpDesk | ITDap Worldwide Solutions"
From2Address = "helpdesk@itdap.com"
ToName2 = replace(Rs_Usuarios.Fields.Item("Apellido").Value, "'", "''") & " " & replace(Rs_Usuarios.Fields.Item("Nombre").Value, "'", "''")
ToAddress2 = replace(Rs_Usuarios.Fields.Item("Mail").Value, "'", "''")
ReplyTo2 = "HelpDesk - ITDap Worldwide Solutions"
ReplyToAddress2 = "helpdesk@itdap.com"

' Mail common
CC= ""
CCAddress=""
BCC=""
BCCAddress=""
Subject = ReadLang(Rs_Lenguaje_Mail,178) & " " & RsMax.fields.item("maximo").value 
Subject2 = RsProblema1.Fields.Item("DetalleProblema").Value & " - " & RsProblema2.Fields.Item("DetalleProblema").Value  & " - " & RsProblema3.Fields.Item("DetalleProblema").Value  
Subject = Subject & " | " & left(Subject2,50)

HTMLBody2= "<html><body style=""font: 11px Verdana; size:10px;"">" 
HTMLBody3= "<b>" & ReadLang(Rs_Lenguaje_Mail,181) & "</b><br>&nbsp;&nbsp;&nbsp;" & vDataForm("observaciones")
HTMLBody4=  "</body></html>" 
HTMLBody = HTMLBody2 & "<br>" & HTMLBody3  & "<br>" & HTMLBody4

'HTMLBody = "test"
Attach = "''"

'Mail al HelpDesk
mailresult = SendMail (From,FromAddress,ToName,ToAddress,CC,CCAddress,BCC, BCCAddress,ReplyTo,ReplyToAddress,Subject,HTMLBody,Attach)

'Mail al Usuario
mailresult = SendMail (From2,From2Address,ToName2,ToAddress2,CC,CCAddress,BCC, BCCAddress,ReplyTo2,ReplyToAddress2,Subject,HTMLBody,Attach) 	


' ----------------------------- End Email sent section -----------------------------


redir = "numtickets.asp?lang=" & vDataForm("lang")
Response.Redirect (redir)

Rs.Close()
Rs2.close()

</script>