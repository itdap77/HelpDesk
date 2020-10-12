function FormatDate(vFecha, vFormato)
dim vDia, vMes, vAnio

if isdate(vFecha) then
    vDia = day(cdate(vFecha))
    if vDia < 10 then
	   vDia = "0" & vDia
	end if
	vMes = month(cdate(vFecha))
    if vMes < 10 then
	   vMes = "0" & vMes
	end if
	vAnio = year(cdate(vFecha))
	
    select case vFormato	
    case "DD/MM/YYYY"
		FormatDate = cstr(vDia) & "/" & cstr(vMes)  & "/" & cstr(vAnio)
    case "MM/DD/YYYY"
		FormatDate = cstr(vMes) & "/" & cstr(vDia)  & "/" & cstr(vAnio)
    case "YYYY/MM/DD"
		FormatDate = cstr(vAnio) & "/" & cstr(vMes) & "/" & cstr(vDia)
	case "MM/YYYY"
		FormatDate = cstr(vMes) & "/" & cstr(vAnio)
	case "DD/MM"
		FormatDate = cstr(vDia) & "/" & cstr(vMes)
    case "YYYYMMDD"
		FormatDate = cstr(vAnio) & cstr(vMes) & cstr(vDia) 
	case "YYYY"
		FormatDate = cstr(vAnio)
	case "MM"
		FormatDate = cstr(vMes) 	
	case "DD"
		FormatDate = cstr(vDia) 
	case else
		FormatDate = ""
	end select	
	
else
		FormatDate = ""
end if
End function

function ValidateDate(oFecha)
    if ucase(trim(ofecha.value)) = "H" then
	   oFecha.value = date
	else
		if not isdate(ofecha.value) then
			ofecha.value = ""
		end if
	end if   
End function
