Imports System.Web.Mvc
Imports Excel = Microsoft.Office.Interop
Imports System.IO



Namespace Controllers
    Public Class ProductController
        Inherits Controller

        ' GET: Product
        Function Index() As ActionResult
            Return View()
        End Function


        <HttpPost()>
        Function Import(ByVal excelfile As HttpPostedFileBase) As ActionResult

            If (excelfile Is Nothing Or excelfile.ContentLength = 0) Then
                ViewBag.Error = "Por favor ingrese Excel<br>"
                Return View("Index")
            Else
                If (excelfile.FileName.EndsWith("xls") Or excelfile.FileName.EndsWith("xlsx")) Then
                    Dim path As String = Server.MapPath("~/Content/" + excelfile.FileName)
                    If (System.IO.File.Exists(path)) Then
                        'System.IO.File.Delete(path)
                    End If
                    'excelfile.SaveAs(path)
                    Dim application As Excel.Excel.Application = New Excel.Excel.Application
                    Dim WorkBook As Excel.Excel.Workbook = application.Workbooks.Open(path)
                    Dim WorkSheet As Excel.Excel.Worksheet = WorkBook.ActiveSheet
                    Dim Range As Excel.Excel.Range = WorkSheet.UsedRange
                    Dim i As Integer
                    Dim Lista As List(Of Prueba) = New List(Of Prueba)


                    For i = 14 To Range.Rows.Count
                        Dim p As Prueba = New Prueba
                        p.area = Range.Cells(i, 1).ToString
                        p.elemento = Range.Cells(i, 2).ToString
                        p.cantidad = Range.Cells(i, 3).ToString
                        p.largo = Range.Cells(i, 4).ToString
                        p.ancho = Range.Cells(i, 5).ToString
                        p.alto = Range.Cells(i, 6).ToString
                        p.lado = Range.Cells(i, 7).ToString

                        Lista.Add(p)
                    Next
                    ViewBag.Lista = Lista
                    Return View("Success")
                    Else
                        ViewBag.Error = "Archivo Incorrecto <br>"
                    Return View("Index")
                End If
            End If

        End Function


    End Class
End Namespace