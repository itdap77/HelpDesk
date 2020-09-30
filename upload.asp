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

response.addHeader "Cache-Control","no-store" 
response.addHeader "Pragma","no-cache" 
response.Expires = 0 
response.addHeader "Cache-Control" ,"no-cache, must-revalidate" 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Upload Imagén</title>

<script LANGUAGE="JavaScript">
extArray = new Array(".jpg", ".gif", ".doc", ".txt", ".zip", ".rar");
function LimitAttach(form, file) 
{
allowSubmit = false;
if (!file) return;

while (file.indexOf("\\") != -1)

file = file.slice(file.indexOf("\\") + 1);
ext = file.slice(file.indexOf(".")).toLowerCase();

for (var i = 0; i < extArray.length; i++) 
{
if (extArray[i] == ext) { allowSubmit = true; break; }
}
if (allowSubmit) form.submit();
else
alert("Se permiten únicamente archivos con la extención: " 
+ (extArray.join(" ")) + "\nPor favor, seleccione otro archivo "
+ "e intente de nuevo.");
return false;
}
</script>

<STYLE type=text/css>
.presenta {
BORDER-RIGHT: #6699CC 1px solid; BORDER-TOP: #6699CC 1px solid; FONT: 8.5pt , verdana; BORDER-LEFT: #6699CC 1px solid; COLOR: #003399; BORDER-BOTTOM: #6699CC 1px solid; LETTER-SPACING: 0px; BACKGROUND-COLOR: #FCFCF3
}
.presenta2 {
BORDER-RIGHT: #FF6600 1px solid; BORDER-TOP: #FF6600 1px solid; FONT: 8.5pt , verdana; BORDER-LEFT: #FF6600 1px solid; COLOR: #000000; BORDER-BOTTOM: #FF6600 1px solid; LETTER-SPACING: 0px; BACKGROUND-COLOR: #C0C0C0
}
</STYLE>

</head>
<body text="#000000" topmargin="0" leftmargin="0" bgcolor="#F3F3F3">

<table width="100%" cellspacing="4" border="0" style="border:1px solid #C0C0C0; border-collapse: collapse; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1" bordercolor="#111111" bgcolor="#EEF3F9">
<tr>
<td height="23">
<%
Func = Request("Func")

If Len(Func) = 0 then
Func = 1
End If

Select Case Func
case 1
%>

<FORM ENCTYPE=multipart/form-data ACTION=upload.asp?func=2 METHOD=POST ID=form1 NAME=form1>
      <p><br>
        <B> <FONT FACE=Arial style="font-size: 9pt">Instrucciones:</FONT><FONT FACE=Verdana Size=1><br>
        </FONT> </B> <FONT FACE=Verdana Size=1><BR>
        </FONT> <FONT FACE=Arial style="font-size: 8pt" color="#606060">1) Seleccione 
        el archivo que corresponde al documento a ingresar.<BR>
        2) El archivo debe estar en formato .JPG, .GIF, .DOC <br>
        3) El archivo no debe superar los 300 KB de tamaño.<br>
        4) Pulsa el botón <B>Subir Archivo</B></FONT><FONT FACE=Verdana Size=1><BR>
        <BR>
        <FONT COLOR=Red>Nota: dependiendo del tamaño de los archivos a subir, 
        la transferencia puede tardar varios minutos.</FONT><BR>
        <BR>
        <BR>
        </FONT> 
        <INPUT NAME=File1 SIZE=30 TYPE=File class=presenta>
        <BR>
        <br>
        &nbsp;
        <INPUT TYPE=button VALUE='Subir Archivo' class=presenta2 onclick="JavaScript:LimitAttach(this.form,this.form.File1.value);">
        <BR>
        <BR>
        <%
Case 2
ForWriting = 2
adLongVarChar = 201
lngNumberUploaded = 0

noBytes = Request.TotalBytes 
binData = Request.BinaryRead (noBytes)

Set RST = CreateObject("ADODB.Recordset")
LenBinary = LenB(binData)

MaxBytes = 300000

If noBytes > MaxBytes then  'Chequea Bytes a subir
	Lenbinary = 0
	KBLimit = 1
End IF

if LenBinary > 0 then
RST.Fields.Append "myBinary", adLongVarChar, LenBinary
RST.Open
RST.AddNew
RST("myBinary").AppendChunk BinData
RST.Update
strDataWhole = RST("myBinary")
End If

strBoundry = Request.ServerVariables ("HTTP_CONTENT_TYPE")
lngBoundryPos = instr(1,strBoundry,"boundary=") + 8 
strBoundry = "--" & right(strBoundry,len(strBoundry)-lngBoundryPos)

lngCurrentBegin = instr(1,strDataWhole,strBoundry)
lngCurrentEnd = instr(lngCurrentBegin + 1,strDataWhole,strBoundry) - 1

do while lngCurrentEnd > 0
strData = mid(strDataWhole,lngCurrentBegin, lngCurrentEnd - lngCurrentBegin)
strDataWhole = replace(strDataWhole,strData,"")

lngBeginFileName = instr(1,strdata,"filename=") + 10
lngEndFileName = instr(lngBeginFileName,strData,chr(34)) 

if lngBeginFileName = lngEndFileName and lngNumberUploaded = 0 then
Response.Write "<br><FONT FACE=Verdana SIZE=4>Ha ocurrido un error</FONT><BR><BR>"
Response.Write "<FONT FACE=Verdana SIZE=2><B>Explicación:</B><BR>"
Response.Write "1) Deberías seleccionar al menos 1 archivo.<BR><br>"
Response.Write "<B>Solución:</B><BR>"
Response.Write "Retrocede pulsando el botón de <B>Retroceder</B> e inténtalo de nuevo.</FONT><BR><BR>"
Response.Write "<INPUT TYPE=Button onclick=history.go(-1) value='<< Retroceder' ID='button'1 NAME='button'1 Class=presenta2>"
Response.End 
end if

if lngBeginFileName <> lngEndFileName then
strFilename = mid(strData,lngBeginFileName,lngEndFileName - lngBeginFileName)
tmpLng = instr(1,strFilename,"\")

Do While tmpLng > 0
PrevPos = tmpLng
tmpLng = instr(PrevPos + 1,strFilename,"\")
Loop

FileName = right(strFilename,len(strFileName) - PrevPos)

lngCT = instr(1,strData,"Content-Type:")

If lngCT > 0 then
lngBeginPos = instr(lngCT,strData,chr(13) & chr(10)) + 4
Else
lngBeginPos = lngEndFileName
End If

lngEndPos = len(strData) 

lngDataLenth = lngEndPos - lngBeginPos

strFileData = mid(strData,lngBeginPos,lngDataLenth)

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(server.mappath(".") & "\Upload\" & FileName, ForWriting, True)
f.Write strFileData
set f = nothing
set fso = nothing

lngNumberUploaded = lngNumberUploaded + 1

End If

lngCurrentBegin = instr(1,strDataWhole,strBoundry)
lngCurrentEnd = instr(lngCurrentBegin + 1,strDataWhole,strBoundry) - 1
loop

Response.Write "<br>" 

if KBLimit <> 1 then
Response.Write "<FONT FACE=Verdana SIZE=4>El archivo fue registrado con éxito</FONT><BR><BR>"
Response.Write "<FONT FACE=Verdana SIZE=2>" & lngNumberUploaded & " archivo subidos al servidor.</FONT><BR><BR>"
Else
	Response.write ("<FONT FACE=Verdana SIZE=3>El Archivo supero el limite de: ") & left(MaxBytes,3) & " KBytes.</FONT><BR><BR>"
End If
Response.Write "<INPUT TYPE=Button onclick='document.location=" & chr(34) & "upload.asp" & chr(34) & "' VALUE='<< Registrar otro documento' ID='button'1 NAME='button'1 Class=presenta2>"
End Select 
%>
      </p>
      </form>
    </td>
</tr>
</table>
</body>

</html>
