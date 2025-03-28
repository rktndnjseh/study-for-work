' EPPlus로 데이터 처리 후 Interop.Excel로 결과를 편집 가능한 워크북에 출력
' 지연시간이 완전히 사라졌지만 몇몇 기능이 동작 안함
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports OfficeOpenXml
Imports System.Reflection.Emit
Imports System.ComponentModel

Public Class Form1
    Public Sub New()
        InitializeComponent()
        ' EPPlus 라이선스 설정
        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial
    End Sub
    Dim filePath As String = ""
    Dim exportPath As String = Path.Combine(
        System.Windows.Forms.Application.StartupPath,
  "통합분석결과.xlsx"
)
    'Directory.GetParent(Directory.GetParent(Directory.GetParent(Directory.GetParent(System.Windows.Forms.Application.StartupPath).FullName).FullName).FullName).FullName,


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not LoadExcelPath() Then Return

        Dim package As ExcelPackage = LoadOrCreatePackage()

        DeleteSheetIfExists(package, "지역별 수량 합계")
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

        SaveAndOpenPackage(package)
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not LoadExcelPath() Then Return

        Dim package = LoadOrCreatePackage()
        DeleteSheetIfExists(package, "지역별 금액 합계")

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

        SaveAndOpenPackage(package)
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Not LoadExcelPath() Then Return

        Dim package = LoadOrCreatePackage()

        Using origin = New ExcelPackage(New FileInfo(filePath))
            For Each s In origin.Workbook.Worksheets
                Dim 지역명 = s.Cells("B4").Text.Trim()
                If String.IsNullOrWhiteSpace(지역명) Then Continue For

                Dim baseName = 지역명 & " 수량"
                Dim sheetName = baseName
                Dim index = 1
                While package.Workbook.Worksheets.Any(Function(ws) ws.Name = sheetName)
                    sheetName = baseName & " (" & index & ")"
                    index += 1
                End While

                Dim sheet = package.Workbook.Worksheets.Add(sheetName)
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

        SaveAndOpenPackage(package)
    End Sub




    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Not LoadExcelPath() Then Return

        Dim package = LoadOrCreatePackage()

        Using origin = New ExcelPackage(New FileInfo(filePath))
            For Each s In origin.Workbook.Worksheets
                Dim 지역명 = s.Cells("B4").Text.Trim()
                If String.IsNullOrWhiteSpace(지역명) Then Continue For

                Dim baseName = 지역명 & " 금액"
                Dim sheetName = baseName
                Dim index = 1
                While package.Workbook.Worksheets.Any(Function(ws) ws.Name = sheetName)
                    sheetName = baseName & " (" & index & ")"
                    index += 1
                End While

                Dim sheet = package.Workbook.Worksheets.Add(sheetName)
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

        SaveAndOpenPackage(package)
    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Not LoadExcelPath() Then Return

        Dim package = LoadOrCreatePackage()
        DeleteSheetIfExists(package, "통합 수량")

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

        SaveAndOpenPackage(package)
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Not LoadExcelPath() Then Return

        Dim package = LoadOrCreatePackage()
        DeleteSheetIfExists(package, "통합 금액")

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

        SaveAndOpenPackage(package)
    End Sub

    Private Sub txtFilePath_TextChanged(sender As Object, e As EventArgs) Handles txtFilePath.TextChanged
        Dim textSize As Size = TextRenderer.MeasureText(txtFilePath.Text, txtFilePath.Font)

        ' 최소 너비와 최대 너비 설정 (폼을 넘어가지 않게)
        Dim minWidth As Integer = 100
        Dim maxWidth As Integer = Me.ClientSize.Width - txtFilePath.Left - 20 ' 오른쪽 여백 고려

        ' 계산된 너비 + 약간의 여유 (스크롤 생기지 않게)
        Dim newWidth As Integer = textSize.Width + 20

        ' 범위 내에서만 설정
        txtFilePath.Width = Math.Max(minWidth, Math.Min(newWidth, maxWidth))
    End Sub

    Private Function LoadExcelPath() As Boolean
        If String.IsNullOrEmpty(filePath) Then
            ' OpenFileDialog 사용
            Using openFileDialog As New OpenFileDialog()
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                openFileDialog.Title = "엑셀 파일 선택"

                If openFileDialog.ShowDialog() = DialogResult.OK Then
                    filePath = openFileDialog.FileName
                    txtFilePath.Text = "사용하고 있는 파일 폴더 주소: " & filePath
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
    Private Function LoadOrCreatePackage() As ExcelPackage
        If File.Exists(exportPath) Then
            Return New ExcelPackage(New FileInfo(exportPath))
        Else
            Return New ExcelPackage()
        End If
    End Function

    Private Sub SaveAndOpenPackage(package As ExcelPackage)
        package.SaveAs(New FileInfo(exportPath))
        Dim xlApp As New Application
        xlApp.Workbooks.Open(exportPath)
        xlApp.Visible = True
    End Sub

    Private Sub DeleteSheetIfExists(package As ExcelPackage, sheetName As String)
        Dim existingSheet = package.Workbook.Worksheets.FirstOrDefault(Function(ws) ws.Name = sheetName)
        If existingSheet IsNot Nothing Then
            package.Workbook.Worksheets.Delete(existingSheet)
        End If
    End Sub

End Class