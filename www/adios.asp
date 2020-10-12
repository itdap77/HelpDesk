<%@ LANGUAGE=VBScript %>
<html>
<head>
<title>ITDap Help Desk</title>

<!--- VARIABLES DEL DOCUMENTO --->

<link rel="STYLESHEET" type="text/css" href="library/Styles.css" />
<link rel="STYLESHEET" type="text/css" href="library/NavArea.css"/>
<link rel="STYLESHEET" type="text/css" href="library/MenuArea.css"/>
<link rel="STYLESHEET" type="text/css" href="library/FormArea.css"/>
<link rel="STYLESHEET" type="text/css" href="library/dap.css"/>

<%

If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 



 %>
 
</head>
<body >
	
<table width="780px" align="left" height="100%" border="0" cellspacing="0" cellpadding="0" width="100%" leftmargin="0px">

<tr> 
    <td  align="center"> 
      
    <div align="center">
      <table class="MainTopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0" >
        <tr>
          <td class="Titulo"> 
            <div align="center">Help Desk</div>
          </td>
          <td> </td>
        </tr>
      </table>
 <table class="defaultContent">
 <tr><td>
      <table id="adioscontent">
		   		<tr><td height="20px">&nbsp;</td></tr>
      		<tr>
							<td height="7%"><center><img alt="www.itdap.com" id="img1"  hspace="0" src="images/logo completo 200x54.gif" align="center" border="0" /></center></td>
					</tr>
					<tr><td height="20px">&nbsp;</td></tr>
			    <tr> 
	             <td ><div align="center"><font face="verdana" size="5">
	             <b><% If vDataForm("lang") = "esp" Then 
	                            Response.Write "Sesion Cerrada"
	                        Else
	                            Response.Write "Session Closed"
	                            End If
	                         %>
	                         
	             </b>
	             </font>
	             <br/>
	              <font face="verdana" size="2">
	              <p style="width:80%;">
	              <% If vDataForm("lang") = "esp" Then 
	                            Response.Write "Muchas Gracias por ingresar al Help Desk, a la brevedad nuestro Soporte T&eacute;cnico se pondrá en contacto con usted para asistirlo."
	                        Else
	                            Response.Write "Thank you for enter to the helpdesk, a technician will contact you soon."
	                            End If
	                         %>
	                                   <br/><br/></p></font>
	              </div>
	          	 </td>
	        </tr>
	      <tr><td ><br /><hr width="95%" size="1px"></td>
	      </tr>
	      
	          <tr><td height="30px"><br/><center><font face="verdana" size="2">
	             <% If vDataForm("lang") = "esp" Then 
	                            Response.Write "Ingrese su usuario y contrasena nuevamente para ingresar.."
	                        Else
	                            Response.Write "Enter your user name and password to enter again."
	                            End If
	                         %>
	          </font></center></br></br></td></tr>
	               
	    <tr>
				<td width="100%" align="center" valign="middle" colspan="2">      
				<!--#INCLUDE file="library/login.asp"-->         
				</td>
																               
		</tr>
						<tr><td height="30px" colspan="2"></td></tr>
																    
						<tr>
							<td height="100%">
							</td>
						</tr>
	        <tr >
	          <td valign="bottom"> 
	            <!--#INCLUDE file="library/footernav_2.asp"-->
	          </td>
	        </tr>
	        
      </table>
      </td>
      </tr>
   </table>    
      </div>
    </td>
  </tr>
</body>
</html>