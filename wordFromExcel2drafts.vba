Sub CreateNewWordDocs()
Dim wdApp As Object
Dim wdDoc As Object
Dim ws As Worksheet
Dim r As Integer
Dim templatePath As String

' elegir la hoja de excel en la que se trabaja, modificar si es necesario
Set ws = ThisWorkbook.Sheets("Hoja 1")

' abre y chequea word
On Error Resume Next
Set wdApp = GetObject(, "Word.Application")
If wdApp Is Nothing Then
    Set wdApp = CreateObject("Word.Application")
End If
wdApp.Visible = True
On Error GoTo 0

For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' elije la plantilla/template the word
    If IsEmpty(ws.Cells(r, 7)) Then
        templatePath = "C:\\folder\\file1.dotx"
    Else
        templatePath = "C:\\folder\\file2.dotx"
    End If

    ' crea un nuevo documento basado en la plantilla seleccionada
    Set wdDoc = Nothing
    On Error Resume Next
    Set wdDoc = wdApp.Documents.Add(Template:=templatePath)
    On Error GoTo 0

    If wdDoc Is Nothing Then
        MsgBox "No se pudo abrir la plantilla: " & templatePath, vbCritical
        GoTo NextIteration
    End If

    ' rellena los espacios de la plantilla con la data de la hoja de excel
    ' modificar de acuerdo a la plantilla seleccionada
    With wdDoc.Content
        
        ReplaceTag .Find, "<<name>>", ws.Cells(r, 1).value
        ReplaceTag .Find, "<<data1>>", ws.Cells(r, 2).value
        ' relleno de fecha con formato harcodeado
        ReplaceTag .Find, "<<date>>", Format(ws.Cells(r, 3).value, "dd/mm/yyyy")
        ReplaceTag .Find, "<<data2>>", ws.Cells(r, 4).value
        ReplaceTag .Find, "<<data3>>", ws.Cells(r, 8).value

        ' rellenos numericos, con formato español harcodeado
        ReplaceTag .Find, "<<num1>>", Format(ws.Cells(r, 5).value, "#,##0.00")
        ReplaceTag .Find, "<<num2>>", Format(ws.Cells(r, 7).value, "#,##0.00")
        ReplaceTag .Find, "<<num3>>", Format(ws.Cells(r, 14).value, "#,##0.00")
        ReplaceTag .Find, "<<num4>>", Format(ws.Cells(r, 13).value, "#,##0.00")

        ReplaceTag .Find, "<<data4>>", ws.Cells(r, 9).value
        ReplaceTag .Find, "<<data5>>", ws.Cells(r, 10).value
        ReplaceTag .Find, "<<data6>>", ws.Cells(r, 11).value
        ReplaceTag .Find, "<<data7>>", ws.Cells(r, 12).value
    End With

    ' guarda el documento de word creado con el nombre de la columna 1
    ' modificar la carpeta donde se guardarían los documentos
    wdDoc.SaveAs2 "C:\\folder\\" & ws.Cells(r, 1).Text & ".docx"
    wdDoc.Close SaveChanges:=False

NextIteration:
Next r

' cierra word
wdApp.Quit
Set wdDoc = Nothing
Set wdApp = Nothing

End Sub

' funcion para verificar que se rellenen los placeholderd
Private Sub ReplaceTag(ByRef findObj As Object, ByVal tag As String, ByVal value As Variant)
With findObj
.Text = tag
.Replacement.Text = value
.Forward = True
.Wrap = 1 ' wdFindContinue
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False
.Execute Replace:=2 ' wdReplaceAll
End With
End Sub
