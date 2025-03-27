' EPPlus로 데이터 처리 후 Interop.Excel로 결과를 편집 가능한 워크북에 출력
' 지연시간이 완전히 사라졌지만 몇몇 기능이 동작 안함
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports OfficeOpenXml

Public Class Form1
    Dim filePath As String = ""

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not LoadExcelPath() Then Return
        ' Dim savePath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "지역별수량.xlsx")

        Dim savePath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "지역별수량.xlsx")
        Using package As New ExcelPackage()
            Dim sheet = package.Workbook.Worksheets.Add("지역별 수량 합계")
            sheet.Cells(1, 1).Value = "지역"
            sheet.Cells(1, 2).Value = "수량"

            Dim rowIndex As Integer = 2
            Dim 전체합계 As Integer = 0

            Using origin = New ExcelPackage(New FileInfo(filePath))
                For Each s In origin.Workbook.Worksheets
                    Dim 지역명 = s.Cells("B4").Text.Trim()
                    Dim 합계 = s.Cells(7, 4, 66, 4).Where(Function(c) IsNumeric(c.Value)).Sum(Function(c) Convert.ToInt32(c.Value))
                    If 합계 > 0 Then
                        sheet.Cells(rowIndex, 1).Value = 지역명
                        sheet.Cells(rowIndex, 2).Value = 합계
                        전체합계 += 합계
                        rowIndex += 1
                    End If
                Next
            End Using

            sheet.Cells(rowIndex, 1).Value = "전체 합계"
            sheet.Cells(rowIndex, 2).Value = 전체합계
            package.SaveAs(New FileInfo(savePath))
        End Using

        Dim xlApp As New Application
        xlApp.Workbooks.Open(savePath)
        xlApp.Visible = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not LoadExcelPath() Then Return

        Dim savePath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "지역별금액.xlsx")
        Using package As New ExcelPackage()
            Dim sheet = package.Workbook.Worksheets.Add("지역별 금액 합계")
            sheet.Cells(1, 1).Value = "지역"
            sheet.Cells(1, 2).Value = "금액"

            Dim rowIndex As Integer = 2
            Dim 전체합계 As Long = 0

            Using origin = New ExcelPackage(New FileInfo(filePath))
                For Each s In origin.Workbook.Worksheets
                    Dim 지역명 = s.Cells("B4").Text.Trim()
                    Dim 합계 As Long = 0
                    For Each cell In s.Cells(7, 6, 66, 6)
                        Dim val = cell.Text.Replace(",", "").Trim()
                        Dim parsed As Long
                        If Long.TryParse(val, parsed) Then 합계 += parsed
                    Next
                    If 합계 > 0 Then
                        sheet.Cells(rowIndex, 1).Value = 지역명
                        sheet.Cells(rowIndex, 2).Value = 합계
                        전체합계 += 합계
                        rowIndex += 1
                    End If
                Next
            End Using

            sheet.Cells(rowIndex, 1).Value = "전체 금액 합계"
            sheet.Cells(rowIndex, 2).Value = 전체합계
            package.SaveAs(New FileInfo(savePath))
        End Using

        Dim xlApp As New Application
        xlApp.Workbooks.Open(savePath)
        xlApp.Visible = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Not LoadExcelPath() Then Return

        Dim savePath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "지역별품목수량.xlsx")
        Using package As New ExcelPackage()
            Using origin = New ExcelPackage(New FileInfo(filePath))
                For Each s In origin.Workbook.Worksheets
                    Dim 지역명 = s.Cells("B4").Text.Trim()
                    Dim sheet = package.Workbook.Worksheets.Add(지역명 & " 수량")
                    sheet.Cells(1, 1).Value = "품목"
                    sheet.Cells(1, 2).Value = "수량"

                    Dim dict As New Dictionary(Of String, Integer)
                    For i = 7 To 66
                        Dim 품목 = s.Cells(i, 2).Text.Trim()
                        Dim 수량 = s.Cells(i, 4).Text.Trim()
                        If 품목 <> "" AndAlso IsNumeric(수량) Then
                            If dict.ContainsKey(품목) Then dict(품목) += Convert.ToInt32(수량) Else dict.Add(품목, Convert.ToInt32(수량))
                        End If
                    Next

                    Dim row = 2
                    For Each kvp In dict
                        sheet.Cells(row, 1).Value = kvp.Key
                        sheet.Cells(row, 2).Value = kvp.Value
                        row += 1
                    Next
                Next
            End Using
            package.SaveAs(New FileInfo(savePath))
        End Using

        Dim xlApp As New Application
        xlApp.Workbooks.Open(savePath)
        xlApp.Visible = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Not LoadExcelPath() Then Return

        Dim savePath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "지역별품목금액.xlsx")
        Using package As New ExcelPackage()
            Using origin = New ExcelPackage(New FileInfo(filePath))
                For Each s In origin.Workbook.Worksheets
                    Dim 지역명 = s.Cells("B4").Text.Trim()
                    Dim sheet = package.Workbook.Worksheets.Add(지역명 & " 금액")
                    sheet.Cells(1, 1).Value = "품목"
                    sheet.Cells(1, 2).Value = "금액"

                    Dim dict As New Dictionary(Of String, Decimal)
                    For i = 7 To 66
                        Dim 품목 = s.Cells(i, 2).Text.Trim()
                        Dim val = s.Cells(i, 6).Text.Replace(",", "").Trim()
                        Dim 금액 As Decimal
                        If 품목 <> "" AndAlso Decimal.TryParse(val, 금액) Then
                            If dict.ContainsKey(품목) Then dict(품목) += 금액 Else dict.Add(품목, 금액)
                        End If
                    Next

                    Dim row = 2
                    For Each kvp In dict
                        sheet.Cells(row, 1).Value = kvp.Key
                        sheet.Cells(row, 2).Value = kvp.Value
                        row += 1
                    Next
                Next
            End Using
            package.SaveAs(New FileInfo(savePath))
        End Using

        Dim xlApp As New Application
        xlApp.Workbooks.Open(savePath)
        xlApp.Visible = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Not LoadExcelPath() Then Return

        Dim savePath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "통합품목수량.xlsx")
        Using package As New ExcelPackage()
            Dim sheet = package.Workbook.Worksheets.Add("통합 수량")
            sheet.Cells(1, 1).Value = "품목"
            sheet.Cells(1, 2).Value = "수량"

            Dim dict As New Dictionary(Of String, Integer)
            Using origin = New ExcelPackage(New FileInfo(filePath))
                For Each s In origin.Workbook.Worksheets
                    For i = 7 To 66
                        Dim 품목 = s.Cells(i, 2).Text.Trim()
                        Dim 수량 = s.Cells(i, 4).Text.Trim()
                        If 품목 <> "" AndAlso IsNumeric(수량) Then
                            If dict.ContainsKey(품목) Then dict(품목) += Convert.ToInt32(수량) Else dict.Add(품목, Convert.ToInt32(수량))
                        End If
                    Next
                Next
            End Using

            Dim row = 2
            For Each kvp In dict
                sheet.Cells(row, 1).Value = kvp.Key
                sheet.Cells(row, 2).Value = kvp.Value
                row += 1
            Next

            package.SaveAs(New FileInfo(savePath))
        End Using

        Dim xlApp As New Application
        xlApp.Workbooks.Open(savePath)
        xlApp.Visible = True
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Not LoadExcelPath() Then Return

        Dim savePath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "통합품목금액.xlsx")
        Using package As New ExcelPackage()
            Dim sheet = package.Workbook.Worksheets.Add("통합 금액")
            sheet.Cells(1, 1).Value = "품목"
            sheet.Cells(1, 2).Value = "금액"

            Dim dict As New Dictionary(Of String, Decimal)
            Using origin = New ExcelPackage(New FileInfo(filePath))
                For Each s In origin.Workbook.Worksheets
                    For i = 7 To 66
                        Dim 품목 = s.Cells(i, 2).Text.Trim()
                        Dim val = s.Cells(i, 6).Text.Replace(",", "").Trim()
                        Dim 금액 As Decimal
                        If 품목 <> "" AndAlso Decimal.TryParse(val, 금액) Then
                            If dict.ContainsKey(품목) Then dict(품목) += 금액 Else dict.Add(품목, 금액)
                        End If
                    Next
                Next
            End Using

            Dim row = 2
            For Each kvp In dict
                sheet.Cells(row, 1).Value = kvp.Key
                sheet.Cells(row, 2).Value = kvp.Value
                row += 1
            Next

            package.SaveAs(New FileInfo(savePath))
        End Using

        Dim xlApp As New Application
        xlApp.Workbooks.Open(savePath)
        xlApp.Visible = True
    End Sub

    Private Function LoadExcelPath() As Boolean
        If String.IsNullOrEmpty(filePath) Then
            ' OpenFileDialog 사용
            Using openFileDialog As New OpenFileDialog()
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                openFileDialog.Title = "엑셀 파일 선택"

                If openFileDialog.ShowDialog() = DialogResult.OK Then
                    filePath = openFileDialog.FileName
                Else
                    MessageBox.Show("파일을 선택하지 않았습니다.")
                    Return False
                End If
            End Using
        End If

        ' 파일 존재 여부 확인
        If Not File.Exists(filePath) Then
            MessageBox.Show("엑셀 파일이 존재하지 않습니다.")
            Return False
        End If
        Return True
    End Function
End Class
