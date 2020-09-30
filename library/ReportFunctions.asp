<SCRIPT RUNAT=SERVER Language="VBScript">
'*****************************************************************
' Scripting Library 
' Manejo de interfaces para ASP.
'
' Copyright 2002. Todos los derechos reservados.
'*****************************************************************


'*****************************************************************

'ENUM de tipos de datos
Const mtString = 0
Const mtNumeric = 1
Const mtDate = 2
Const mtBoolean = 3

'ENUM de tipos de condiciones
Const mtEqual = 0
Const mtBetween = 1



Function GetFormulaFromRequest(vData, vReportName)

Dim vPosLast, vPosNext, vValue, vField, vPosValue, vWhere, strNullValue, vCondition

    'Ejecuta los reemplazos en la info del form
    vData = Replace(vData, "%2F", "/")
    
    strNullValue = "NULL"
    
    vPosLast = 1
    vPosNext = InStr(vPosLast, vData, "&") + 1

    'Definir vWhere
    While (vPosNext <> 0 And vPosNext < Len(vData))
        vPosValue = InStr(vPosLast, vData, "=") + 1

        If vPosValue = 1 Then
            vValue = ""
            vPosNext = InStr(vPosLast, vData, "&") + 1
        Else
            vField = Mid(vData, vPosLast, vPosValue - vPosLast - 1)
            vCondition = Mid(vField, 1, 2)
            vField = "{" & vReportName & "." & Mid(vField, 3) & "}"

            vPosNext = InStr(vPosLast, vData, "&") + 1

            If vPosNext > 1 Then
                vValue = Mid(vData, vPosValue, vPosNext - vPosValue - 1)
            Else
                vValue = Mid(vData, vPosValue, vPosValue - 1)
                vPosNext = 0
            End If
        End If

        'Evalua los valores recibidos en el Form o QueryString (los que empiezan con % para Like, y _ para igual)
        If Len(vValue) > 0 Then

            Select Case vCondition
            Case "C1"    'Consultas por igual
                If Not Len(vWhere) = 0 Then
                    vWhere = vWhere & " AND "
                End If

                If vValue = strNullValue Then
                    vWhere = vWhere & "IsNull(" & vField & ")"
                Else
                    If IsDate(vValue) Then
                        vValue = "DateTime(" & Year(vValue) & "," & Month(vValue) & "," & Day(vValue) & ")"
                    ElseIf Not IsNumeric(vValue) Then
                        vValue = "'" & vValue & "'"
                    End If

                    vWhere = vWhere & vField & "=" & vValue
                End If
            Case "C2"    'Consultas por like
                If Not Len(vWhere) = 0 Then
                    vWhere = vWhere & " AND "
                End If

                vWhere = vWhere & vField & " like'" & vValue & "'"
            Case "C3"    'Consultas por Menor o igual
                If Not Len(vWhere) = 0 Then
                    vWhere = vWhere & " AND "
                End If

                If IsDate(vValue) Then
                    vValue = "DateTime(" & Year(vValue) & "," & Month(vValue) & "," & Day(vValue) & ")"
                ElseIf Not IsNumeric(vValue) Then
                    vValue = "'" & vValue & "'"
                End If

                vWhere = vWhere & vField & " <=" & vValue
            Case "C4"    'Consultas por Mayor o igual
                If Not Len(vWhere) = 0 Then
                    vWhere = vWhere & " AND "
                End If

                If IsDate(vValue) Then
                    vValue = "DateTime(" & Year(vValue) & "," & Month(vValue) & "," & Day(vValue) & ")"
                ElseIf Not IsNumeric(vValue) Then
                    vValue = "'" & vValue & "'"
                End If

                vWhere = vWhere & vField & " >=" & vValue
            Case "C5"    'Consultas por distinto
                If Not Len(vWhere) = 0 Then
                    vWhere = vWhere & " AND "
                End If

                If vValue = strNullValue Then
                    vWhere = vWhere & "NOT IsNull(" & vField & ")"
                Else
                    If IsDate(vValue) Then
                        vValue = "DateTime(" & Year(vValue) & "," & Month(vValue) & "," & Day(vValue) & ")"
                    ElseIf Not IsNumeric(vValue) Then
                        vValue = "'" & vValue & "'"
                    End If

                    vWhere = vWhere & "NOT " & vField & "=" & vValue
                End If
            End Select
        End If


        vPosLast = vPosNext
    
    Wend

    GetFormulaFromRequest = vWhere
    
End Function

Function GetFilterTextFromRequest(vRequest, oRSFields)

Dim vFilterText, vCurrentFieldText, vValor1, vValor2, vEntidad, oRSList

	Set vEntidad = Server.CreateObject("MTNeon.Entidad")
	Set oRSList = Server.CreateObject("ADODB.Recordset")

	If Not oRSFields.BOF Then
		oRSFields.MoveFirst
	End If
	
	While Not oRSFields.EOF
		If vRequest("C1" & oRSFields("FieldName")).Count Then
			vValor1 = vRequest("C1" & oRSFields("FieldName"))
		End If
		If vRequest("C2" & oRSFields("FieldName")).Count Then
			vValor1 = vRequest("C2" & oRSFields("FieldName"))
		End If
		If vRequest("C4" & oRSFields("FieldName")).Count Then
			vValor1 = vRequest("C4" & oRSFields("FieldName"))
		End If
		If vRequest("C3" & oRSFields("FieldName")).Count Then
			If Len(vValor1) = 0 Then
				vValor1 = vRequest("C3" & oRSFields("FieldName"))
			Else
				vValor2 = vRequest("C3" & oRSFields("FieldName"))
			End If
		End If

		If Not Len(vValor1) = 0 Then
			If Not Len(oRSFields("ListSource")) = 0 Then
				'Buscar el valor para el ID de Cada valor
				Set oRSList = vEntidad.GetOne(oRSFields("ListSource"),vValor1)
				vValor1 = oRSList(1)
				
				If Not Len(vValor2) = 0 Then
					Set oRSList = vEntidad.GetOne(oRSFields("ListSource"),vValor2)
					vValor2 = oRSList(1)
				End If				
				oRSList.Close
			End If
			
			Select Case oRSFields("DataType")
			Case mtBoolean
				If vValor1 = "0" Then
					vValor1 = "No"
				Else
					vValor1 = "Si"
				End If
				If vValor2 = "0" Then
					vValor2 = "No"
				Else
					vValor2 = "Si"
				End If
			End Select
			
			vCurrentFieldText = oRSFields("Caption") & ": "
			Select Case oRSFields("ConditionType")
			Case mtBetween
				vCurrentFieldText = vCurrentFieldText & "entre " & vValor1 & " y " & vValor2
			Case Else
				vCurrentFieldText = vCurrentFieldText & vValor1
			End Select
			
			If Not Len(vFilterText) = 0 Then
				vFilterText = vFilterText & "; "
			End If
			vFilterText = vFilterText & vCurrentFieldText

		End If
		
		vCurrentFieldText = ""
		vValor1 = ""
		vValor2 = ""
		
		oRSFields.MoveNext
	Wend

	If Not oRSFields.BOF Then
		oRSFields.MoveFirst
	End If
	
	If Len(vFilterText)=0 Then
		vFilterText = "Datos sin filtrar"
	End If
	
	GetFilterTextFromRequest = vFilterText
	    
End Function

</SCRIPT>
