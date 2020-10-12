<%@LANGUAGE="VBSCRIPT"%> 
<!--#INCLUDE file="library/PageOpen.asp"-->

<%
ValidSession()

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 

dim Login__User, ID_User
ID_User = Session("IDU")
Login__User = Session("NombreUsuario")

set Rs_Estado_V = Server.CreateObject("ADODB.Recordset")
Rs_Estado_V.ActiveConnection = MM_HelpDesk_STRING
Rs_Estado_V.Source = "{call dbo.SPU_Estado_V}"
Rs_Estado_V.Open()
Rs_Estado_V_numRows = 0

set Rs_Problema_ID = Server.CreateObject("ADODB.Recordset")
Rs_Problema_ID.ActiveConnection = MM_HelpDesk_STRING


If (Session("Login") = 1) Then
	Select Case Request.QueryString("Accion")
	
				 Case "A"			' Asignados 
				 		If Session("IDGrupo")= 1 or Session("IDgrupo")=4 Then
						set Rs_Ticket_Vs = Server.CreateObject("ADODB.Recordset")
						Rs_Ticket_Vs.ActiveConnection = MM_HelpDesk_STRING
						Rs_Ticket_Vs.Source = "{call dbo.SPU_Ticket_VA(" + Replace(ID_User, "'", "''") + ")}"
						Rs_Ticket_Vs.Open()
						Rs_Ticket_Vs_numRows = 0
					Else
					If vDataForm("lang") = "esp" or session("ididioma") = 1 Then
						Response.Redirect "Mensaje.asp?Msg=No tiene permisos para ver los casos asignados.&BackToURL=vertickets.asp?Accion=P"
						Else
						Response.Redirect "Mensaje.asp?Msg=You dont have permission to see the asigned cases.&BackToURL=vertickets.asp?Accion=P"
						End If
						
					End If
					
				 Case "P"		' Tickets Propios
						set Rs_Ticket_Vs = Server.CreateObject("ADODB.Recordset")
						Rs_Ticket_Vs.ActiveConnection = MM_HelpDesk_STRING
						Rs_Ticket_Vs.Source = "{call dbo.SPU_Ticket_VP('" + Replace(ID_User, "'", "''") + "')}"
						Rs_Ticket_Vs.Open()
						Rs_Ticket_Vs_numRows = 0
						
				 Case "T"	' Todos los tickets - Solo administradores 
				 		    If Session("IDGrupo") = 1 Then
						 		    set Rs_Ticket_Vs = Server.CreateObject("ADODB.Recordset")
								    Rs_Ticket_Vs.ActiveConnection = MM_HelpDesk_STRING
								    Rs_Ticket_Vs.Source = "{call dbo.SPU_Ticket_V}"
								    Rs_Ticket_Vs.Open()
								    Rs_Ticket_Vs_numRows = 0
						Else
						If vDataForm("lang") = "esp" or session("ididioma") = 1 Then
                            Response.Redirect "Mensaje.asp?Msg=No tiene permisos para ver todos los casos.&BackToURL=vertickets.asp?Accion=P"
                        Else
                            Response.Redirect "Mensaje.asp?Msg=You dont have permission to list all cases.&BackToURL=vertickets.asp?Accion=P"    
                        End If
				End if
	End Select

Else
        If vDataForm("lang") = "esp" or session("ididioma") = 1 Then
		    Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
		Else
		    Response.Redirect "Mensaje.asp?Msg=You are not logged into the system.&BackToURL=default.asp"
		End If
End If

Dim Repeat1__numRows, Total, Repeat1__index
Repeat1__numRows = 25
Repeat1__numRows2 = Repeat1__numRows  
Repeat1__index = 0
Rs_Ticket_Vs_numRows = Rs_Ticket_Vs_numRows + Repeat1__numRows

If Not Rs_Ticket_Vs.Eof and Not Rs_Ticket_Vs.BOF then
  Total = 0
  While (Not Rs_Ticket_Vs.EOF)
   Total = Total + 1
   Rs_Ticket_Vs.MoveNext
  Wend
Rs_Ticket_Vs.MoveFirst
End If

Rs_Ticket_Vs_total = Rs_Ticket_Vs.RecordCount

If (Rs_Ticket_Vs_numRows < 0) Then
  Rs_Ticket_Vs_numRows = Rs_Ticket_Vs_total
Elseif (Rs_Ticket_Vs_numRows = 0) Then
  Rs_Ticket_Vs_numRows = 1
End If

Rs_Ticket_Vs_first = 1
Rs_Ticket_Vs_last  = Rs_Ticket_Vs_first + Rs_Ticket_Vs_numRows - 1

If (Rs_Ticket_Vs_total <> -1) Then
  If (Rs_Ticket_Vs_first > Rs_Ticket_Vs_total) Then Rs_Ticket_Vs_first = Rs_Ticket_Vs_total
  If (Rs_Ticket_Vs_last > Rs_Ticket_Vs_total) Then Rs_Ticket_Vs_last = Rs_Ticket_Vs_total
  If (Rs_Ticket_Vs_numRows > Rs_Ticket_Vs_total) Then Rs_Ticket_Vs_numRows = Rs_Ticket_Vs_total
End If

Set MM_rs    = Rs_Ticket_Vs
MM_rsCount   = Rs_Ticket_Vs_total
MM_size      = Rs_Ticket_Vs_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If

if (Not MM_paramIsDefined And MM_rsCount <> 0) then
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record
End If

If (MM_rsCount = -1) Then
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If

Rs_Ticket_Vs_first = MM_offset + 1
Rs_Ticket_Vs_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (Rs_Ticket_Vs_first > MM_rsCount) Then Rs_Ticket_Vs_first = MM_rsCount
  If (Rs_Ticket_Vs_last > MM_rsCount) Then Rs_Ticket_Vs_last = MM_rsCount
End If

MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)

MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)


%>
      
   
      <script type="text/javascript" language="JavaScript">
function verticket(URL,DATOS) {window.open(URL,"_blank",DATOS)};

</script>
<%
'Response.AddHeader "Refresh", "240"
%>
      <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="Titulo">
            <% Select Case request.QueryString("Accion") 
            				Case "A"
            				 				Call ReadLang(Rs_Lenguaje,58)   
            				Case "P" 
            						 		Call ReadLang(Rs_Lenguaje,67)  
            				Case "T"
            								Call ReadLang(Rs_Lenguaje,68)  
            				End Select
            				%>
          </td>
        </tr>
      </table>
      <table class="ContentArea" border="0" width="100%" cellspacing="0" cellpadding="0" height="100%">
        <tr> 
          <td width="100%"> 

            <table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td width="100%"> 
                  <div align="center"> 
                    <center>
                      <table width="100%" border="0" class="FormBoxHeader">
                        <tr> 
                          <td width="40"> 
                            <div align="center"><b><font color="#FFFFFF"><% Call ReadLang(Rs_Lenguaje,59) %>  
                              </font></b></div>
                          </td>
                          <td width="160"> 
                            <div align="center"><b><font color="#FFFFFF"><% Call ReadLang(Rs_Lenguaje,60) %>  
                              </font></b></div>
                          </td>
                          <td width="320"> 
                            <div align="center"><% Call ReadLang(Rs_Lenguaje,61) %> </div>
                          </td>
                          <td width=""> 
                            <div align="center"><b><font color="#FFFFFF"><% Call ReadLang(Rs_Lenguaje,62) %> </font></b></div>
                          </td>
                          <td> 
                            <div align="center"><% Call ReadLang(Rs_Lenguaje,63) %> </div>
                          </td>
                        </tr>
                      </table>
                      <table border="0" width="75%" class="FormBoxBody" align="center" cellpadding="0" cellspacing="0">
                        <% 
						dim flag
						While ((Repeat1__numRows <> 0) AND (NOT Rs_Ticket_Vs.EOF)) %>
						 <!-- <a href="detalletickets.asp?IDTicket=<% 'Response.Write ( Rs_Ticket_Vs.Fields.Item("IDTicket").Value & "&" & "lang=" & vDataForm("lang") )%>">-->
                            <%
                            
						     Response.Write ("<tr  onclick=""document.location = 'detalletickets.asp?IDTicket=")
						     Response.Write (Rs_Ticket_Vs.Fields.Item("IDTicket").Value)
						     Response.Write ("&lang=")
						     Response.Write (vDataForm("lang"))
						     Response.Write ("';"" ")
						     
							  If flag = 1 then
													Response.Write ("Class=""FormBoxListOddRow""")
													Response.Write (" onMouseOver=this.className='FormBoxListHighlight' onMouseOut=this.className='FormBoxListOddRow'>")
													flag=0
							  Else
													Response.Write ("Class=""FormBoxListEvenRow""")
													Response.Write (" onMouseOver=this.className='FormBoxListHighlight' onMouseOut=this.className='FormBoxListEvenRow'>")
													flag=1
							  End if
							  'Response.Write (">")
							  
							  
							%> 
							<!--</a> -->
							
                          <td width="40" align="center"><%=(Rs_Ticket_Vs.Fields.Item("IDTicket").Value)%> </td>
                          <td width="160" align="left" >&nbsp;&nbsp;&nbsp;&nbsp;<%=(Rs_Ticket_Vs.Fields.Item("FechaTicket").Value)%> </td>
                          <td width="320"> 
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%
													Resp = mid(mid(Rs_Ticket_Vs ("IDProblema"),instr(Rs_Ticket_Vs ("IDProblema"),",")+1,len(Rs_Ticket_Vs ("IDProblema"))),instr(mid(Rs_Ticket_Vs ("IDProblema"),instr(Rs_Ticket_Vs ("IDProblema"),",")+1,len(Rs_Ticket_Vs ("IDProblema"))),",")+1,len(mid(Rs_Ticket_Vs ("IDProblema"),instr(Rs_Ticket_Vs ("IDProblema"),",")+1,len(Rs_Ticket_Vs ("IDProblema")))))
													
													Rs_Problema_ID.Source = "{call dbo.SPU_Problema_ID(" + Resp + ")}"
													Rs_Problema_ID.Open()
													if not Rs_Problema_ID.EOF Then
															If vDataForm("lang") = "esp" or session("ididioma") = 1 Then
																Response.Write (Rs_Problema_ID.Fields.Item("DetalleProblema").Value)
															Else
																Response.Write (Rs_Problema_ID.Fields.Item("ProblemDetail").Value)
															End If
													End If
													%>
                          </td>
                          <td align="center" width=""> 
                            <%
														Rs_Problema_ID.Close()
															Rs_Estado_V.MoveFirst 
															While (NOT Rs_Estado_V.EOF)
																if  (Rs_Estado_V.Fields.Item("IDEstado").Value = Rs_Ticket_Vs.Fields.Item("IDEstado").Value) then
																	If vDataForm("lang") = "esp" Then
																			Response.Write (Rs_Estado_V.Fields.Item("Estado").Value)
																	Else
																			Response.Write (Rs_Estado_V.Fields.Item("state").Value)
																	End If
																	
																end if
														  		Rs_Estado_V.MoveNext()
															Wend
														%>
                          </td>
                          
                          <td width="10px" align="center" >  

															<% 
															If (Rs_Ticket_Vs.Fields.Item("Observaciones").value <> "") then
										   						    Response.Write ("<a href=""javascript:openwin('detalleobs.asp?Observaciones=")
																	Response.Write (Rs_Ticket_Vs.Fields.Item("Observaciones").value) 
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
												  Repeat1__index=Repeat1__index+1
												  Repeat1__numRows=Repeat1__numRows-1
												  Rs_Ticket_Vs.MoveNext()
												Wend
												%>
                      </table>
                    </center>
                  </div>
                </td>
              </tr>
            </table>
            <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="100%" align="center"> 
                <table width="100%" >
                    <tr>
                        <td >
                  <p><b>&nbsp;
									<a class="AccionBlack" href="<%=MM_movePrev%>">&lt;&lt; <% Call ReadLang(Rs_Lenguaje,64) %></a></b>
				 </p>
				 </td>
				 <td style="text-align:center;width:430px;"><font size="2" ><b><% Call ReadLang(Rs_Lenguaje,173) %></b>
                 
                 <% 
                 if MM_offset = 0 Then
                    pagenumber = 1
                 Else
                    pagenumber = (MM_offset / MM_size) + 1
                 End If
                 Response.Write (pagenumber & " / " & roundUp(Total/MM_size))
                 
                 %>
                 
                 </font></td>
				 <td ><p><b><a class="AccionBlack" href="<%=MM_moveNext%>"> <% Call ReadLang(Rs_Lenguaje,65) %>   &gt;&gt;</a></b></p></td>
				 </tr>
				 <tr><td colspan="3" style="text-align:center;">
				 <font size="2" ><b><% Call ReadLang(Rs_Lenguaje,172) %></b>  <%Response.Write (MM_size & " / " & Total) %></font>
				 </td></tr>
				 </table>
                  
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <!--#INCLUDE file="library/PageClose.asp"-->

<%
Rs_Ticket_Vs.Close()
Rs_Estado_V.Close()
%>