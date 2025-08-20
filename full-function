Function NumberToSpanishText(ByVal number As Double) As String
Dim units As Variant
Dim tens As Variant
Dim hundreds As Variant
Dim thousands As Variant
Dim millions As Variant
Dim result As String
Dim intPart As Long

' variables
result = ""
intPart = Int(number)

' arrays de todos los numeros
units = Array("", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE")
tens = Array("", "", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
hundreds = Array("", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
thousands = Array("", "MIL", "DOS MIL", "TRES MIL", "CUATRO MIL", "CINCO MIL", "SEIS MIL", "SIETE MIL", "OCHO MIL", "NUEVE MIL")
millions = Array("", "UN MILLÓN", "DOS MILLONES", "TRES MILLONES", "CUATRO MILLONES", "CINCO MILLONES", "SEIS MILLONES", "SIETE MILLONES", "OCHO MILLONES", "NUEVE MILLONES")

' intPart chequea que el numero este dentro del rango hasta 9 millones
If intPart < 0 Or intPart > 9999999 Then
    NumberToSpanishText = "NÚMERO FUERA DE RANGO"
    Exit Function
End If

' millones
If intPart >= 1000000 Then
    result = millions(Int(intPart / 1000000))
    intPart = intPart Mod 1000000
End If

' miles
If intPart >= 1000 Then
    result = result & " " & NumberToSpanishText(Int(intPart / 1000)) & " MIL"
    intPart = intPart Mod 1000
End If

' cientos
If intPart >= 100 Then
    result = Trim(result & " " & hundreds(Int(intPart / 100)))
    intPart = intPart Mod 100
End If

' decenas, units como ordinales simples 
If intPart < 20 Then
    result = Trim(result & " " & units(intPart))
Else
    result = Trim(result & " " & tens(Int(intPart / 10)))
    If (intPart Mod 10) > 0 Then
        result = Trim(result & " Y " & units(intPart Mod 10))
    End If
End If

' for "CIEN"
If result = "CIENTO " And intPart = 0 Then
    result = "CIEN"
End If

' uppercase para todos
NumberToSpanishText = UCase(Trim(result))

End Function

Sub ProcessData()
Dim ws As Worksheet
Dim r As Integer
Dim result As String

' hoja de trabajo - modificar si es necesario
Set ws = ThisWorkbook.Sheets("Hoja 1") 

' loop desde fila 2, fila 1 es header
For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' si columna 7 está vacía...
    If IsEmpty(ws.Cells(r, 7).Value) Then
        
        result = NumberToSpanishText(ws.Cells(r, 5).Value)
    Else
        ' si columna 7 tiene contenido, usa esa
        result = NumberToSpanishText(ws.Cells(r, 7).Value)
    End If

    ' resultado en columna 8
    ws.Cells(r, 8).Value = result
Next r

End Sub
