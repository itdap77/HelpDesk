<%If Not Len(request.Form) = 0 Then 
	Set vDataForm = request.Form 
Else 
	Set vDataForm = request.QueryString 
End If 
 %>
<!--- Footer  ---->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td colspan="2" class="FooterNotes" style="height:20; border-bottom: gray 1px solid;"> 
     &nbsp;</td>
</tr>
<tr>
 <td colspan="2" class="FooterNotes" >
 <%
 If vDataForm("lang") = "esp" then 
        Response.Write ("Diseño y desarrollo by ITDap Worldwide Solutions")
Else 
        Response.write ("Powered by ITDap Worldwide Solutions")
 End If
  %>
  
 </td>
</tr>
<tr>
   <td class="FooterNotes2" >
    <%
 If vDataForm("lang") = "esp" then 
        Response.Write ("Resolución mínima 800x600 - Optimizado para Microsoft Explorer 5.0 o superior")
Else 
        Response.write ("Minimum resolution 800x600 - Optimized for Microsoft internet Explorer 5.0 or superior")
 End If
  %>
    
   </td>
    <td class="FooterNotes3" >
    <%
 If vDataForm("lang") = "esp" then 
        Response.Write ("Copyright © 2013 ITDap Worldwide Solutions. Todos los derechos reservados.")
Else 
        Response.write ("Copyright © 2013 ITDap Worldwide Solutions. All rights reserved")
 End If
  %>
      
    </td>
</tr>
</table>
<!-- Fin Pie de página -->
