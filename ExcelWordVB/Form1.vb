Imports System.IO
Imports Microsoft.Office.Interop
''Crear una aplicacion en VS  (vb y c#) para generar un documento en formato Word y excel,
''con un texto variable al inicio del documento, solicitar ruta y nombre del archivo
''para guardarlo y mostrarlo automaticamente.
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Dim ruta, dato As String
        Dim dialogoGuardar As SaveFileDialog = New SaveFileDialog()
        dialogoGuardar.Filter = "Archivos Word (*.docx)|*.docx|Hoja de cálculo de Microsoft Excel (*.xlsx)|*.xlsx"
        If (dialogoGuardar.ShowDialog() <> DialogResult.OK) Then

            Return

        ElseIf (dialogoGuardar.FilterIndex = 1) Then
            ruta = dialogoGuardar.FileName
            Dim WordApp = New Word.Application()
            WordApp.Visible = True
            WordApp.Documents.Add()
            dato = txtDato.Text
            WordApp.Selection.TypeText(dato)
            WordApp.ActiveDocument.SaveAs2(ruta)
            MessageBox.Show("Archivo creado")
        ElseIf (dialogoGuardar.FilterIndex = 2) Then
            Dim aplicacion As Excel.Application
            Dim libros_trabajo As Excel.Workbook
            Dim hoja_trabajo As Excel.Worksheet
            aplicacion = New Excel.Application()
            aplicacion.Visible = True
            libros_trabajo = aplicacion.Workbooks.Add()
            hoja_trabajo = libros_trabajo.Worksheets(1)
            hoja_trabajo.Cells(1, 1) = txtDato.Text
            libros_trabajo.SaveAs(dialogoGuardar.FileName)
            MessageBox.Show("Archivo creado")
        End If

    End Sub
End Class
