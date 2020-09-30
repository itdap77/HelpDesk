<% @ LANGUAGE=VBScript %> 

<!--#INCLUDE file="library/Funciones.asp"-->

<script language="VBScript" type="text/vbscript" runat="Server">

ValidSession ()

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

' ----------------------------- language section -----------------------------
Set Rs_Lenguaje_Mail = Server.CreateObject("ADODB.Recordset")
Rs_Lenguaje_Mail.ActiveConnection = MM_HelpDesk_STRING
Rs_Lenguaje_Mail.Source = "{call dbo.SPU_Lenguaje_V(" + cstr(Session("IDidioma")) + ",'" + GetPageName() + "')}"
Rs_Lenguaje_Mail.Open()
' ----------------------------- End language section -----------------------------

Select Case vDataForm("Accion")
	Case "D"        'Derivar
			set Rs_Evento_D = Server.CreateObject("ADODB.Recordset")
			Rs_Evento_D.ActiveConnection = MM_HelpDesk_STRING
			dim Str_Der
			Str_Der = "{call dbo.SPU_Evento_D(" + _
			Replace(vDataForm("IDTicket"), "'", "''") + ",'" +_ 
			Replace(vDataForm("Descripcion"), "'", "''") + "','" + _
			Replace(vDataForm("Observaciones"), "'", "''") + "','" + _
			Replace(vDataForm("login"), "'", "''") + "'," + _
			Replace(vDataForm("IDEstado"), "'", "''") + "," + _
			Replace(vDataForm("IDUsuarioDerivado"), "'", "''") + "," + _
			"Null" +  _
			")}"
			Rs_Evento_D.Source = Str_Der
			Rs_Evento_D.Open()
	        Set Rs_Evento_D = Nothing
	
	Case Else
            
            set Rs_Evento = Server.CreateObject("ADODB.Recordset")
			Rs_Evento.ActiveConnection = MM_HelpDesk_STRING
			dim Str_Evento
			Str_Evento= "{call dbo.SPU_Evento_A(" + _
			Replace(vDataForm("IDTicket"), "'", "''") + ",'" +_ 
			Replace(vDataForm("Descripcion"), "'", "''") + "','" + _
			Replace(vDataForm("Observaciones"), "'", "''") + "','" + _
			Replace(vDataForm("login"), "'", "''") + "'," + _
			Replace(vDataForm("IDEstado"), "'", "''") + "," + _
			Replace(vDataForm("IDUsuarioDerivado"), "'", "''") + "," + _
			"Null" + _
			")}"
			Rs_Evento.Source = Str_Evento
			Rs_Evento.Open()
			Set Rs_Evento = Nothing
			
End Select

if err.number <> 0 then
Response.Write err.Description
response.End
end if

If vDataForm("Accion") <> "I" and  vDataForm("Accion") <> "O" Then
' ----------------------------- Email sent section -----------------------------
        Dim From,FromAddress,Subject,HTMLBody,HTMLBody2,HTMLBody3,HTMLBody4,ToName,ToAddress,Attach,CC,CCAddress,BBC, BBCAddress,ToName2,ToAddress2,From2,From2Address

			        Set Rs_Usuarios= Server.CreateObject("ADODB.Recordset")
			        Rs_Usuarios.ActiveConnection = MM_HelpDesk_STRING
			        Rs_Usuarios.Source = "{call dbo.SPU_Usuarios_L_Full(" + Replace(Session("IDU"), "'", "''") + ")}"
			        Rs_Usuarios.Open()
			
			        Set RsHd= Server.CreateObject("ADODB.Recordset")
			        RsHd.ActiveConnection = MM_HelpDesk_STRING
			        RsHd.Source = "{call dbo.SPU_HelpDesk_X}"
			        RsHd.Open()
	
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

        Subject = ReadLang(Rs_Lenguaje_Mail,198) & " " & vDataForm("IDTicket") 
        Subject2 = vDataForm("Problema")
        Subject = Subject & " | " & left(Subject2,50)

        HTMLBody = "Descripcion: " & vDataForm("Descripcion") & "<br>" & "Observaciones: " & vDataForm("Observaciones") 
        HTMLBody2= "<html><body style=""font: 11px Verdana; size:10px;"">" 
        HTMLBody3= "<b>" & ReadLang(Rs_Lenguaje_Mail,200) & " </b>" & vDataForm("title") & "<br><b>" & ReadLang(Rs_Lenguaje_Mail,199) & "</b> " & vDataForm("Descripcion") & "<b>" & ReadLang(Rs_Lenguaje_Mail,181) & "</b><br>&nbsp;&nbsp;&nbsp;" & vDataForm("Observaciones")
        HTMLBody4=  "</body></html>" 
        HTMLBody = HTMLBody2 & "<br>" & HTMLBody3  & "<br>" & HTMLBody4

        Attach = "''"

        'Mail al HelpDesk
        mailresult = SendMail (From,FromAddress,ToName,ToAddress,CC,CCAddress,BCC, BCCAddress,ReplyTo,ReplyToAddress,Subject,HTMLBody,Attach)

        'Mail al Usuario
        mailresult = SendMail (From2,From2Address,ToName2,ToAddress2,CC,CCAddress,BCC, BCCAddress,ReplyTo2,ReplyToAddress2,Subject,HTMLBody,Attach) 	


        ' ----------------------------- End Email sent section -----------------------------
End If  'envia mail excepto si es check in o check out

set Rs_Usuarios = Nothing
set RsHd = Nothing



</script>
<%Response.Write("<script>window.close();</script>") %>