Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Runtime.InteropServices
' 애플리케이션 설명 - 파일 선택을 할 수 있습니다.
' 지역명이 있는 곳의 지역별 수량 합계, 지역별 금액 합계를 계산하는 프로그램입니다.
' 지역명은 B4에서 가져옵니다.
' 수량 합계는 D4열의 7행~66행을 더합니다.
' 금액 합계는 F4열의 7행~66행을 더합니다.
' 
Public Class Form1
    Dim filePath As String = ""

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Filter = "Excel 파일|*.xlsx;*.xls"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            filePath = OpenFileDialog1.FileName
            ListBox1.Items.Add("선택된 파일: " & filePath)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim xlSheet As Worksheet = Nothing
        ' 👉 기본 경로 설정: 실행 파일이 있는 폴더 기준
        If filePath = "" Then
            filePath = GetDefaultExcelPath()
            If Not File.Exists(filePath) Then
                MessageBox.Show("엑셀 파일을 선택하지 않았고, 기본 파일도 없습니다.")
                Return
            End If
        End If

        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlBook As Workbook = xlApp.Workbooks.Open(filePath)
        Dim 전체합계 As Integer = 0

        ListBox1.Items.Add("-------- 지역별 수량 --------")

        For Each sheet As Worksheet In xlBook.Sheets
            xlSheet = sheet
            Dim 지역명 As String = ""
            If Not IsNothing(xlSheet.Range("B4").Value) Then
                지역명 = xlSheet.Range("B4").Value.ToString().Trim()
            End If

            Dim 지역합계 As Integer = 0
            For i As Integer = 7 To 66
                Dim val = xlSheet.Cells(i, 4).Value
                If IsNumeric(val) Then
                    지역합계 += Convert.ToInt32(val)
                End If
            Next

            If 지역합계 > 0 Then
                ListBox1.Items.Add(지역명 & " : " & 지역합계 & "개")
                전체합계 += 지역합계
            End If
            ReleaseComObject(xlSheet)
        Next

        ListBox1.Items.Add("------------------------------")
        ListBox1.Items.Add("전체 합계 : " & 전체합계 & "개")

        ' 종료 및 해제
        xlBook.Close(False)
        xlApp.Quit()
        ReleaseComObject(xlBook)
        ReleaseComObject(xlApp)
    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub
    Private Function OpenExcelWorkbook(path As String) As (Excel.Application, Excel.Workbook)
        Dim app As New Excel.Application
        Dim book As Excel.Workbook = app.Workbooks.Open(path)
        Return (app, book)
    End Function

    Private Function GetDefaultExcelPath() As String
        Dim slnPath = System.IO.Directory.GetParent(
                    System.IO.Directory.GetParent(
                        System.IO.Directory.GetParent(
                            System.IO.Directory.GetParent(
                                System.Windows.Forms.Application.StartupPath).FullName).FullName).FullName).FullName
        Return Path.Combine(slnPath, "주문내역 - 복사본.xlsx")
    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim xlSheet As Worksheet = Nothing
        If filePath = "" Then
            filePath = GetDefaultExcelPath()
            If Not File.Exists(filePath) Then
                MessageBox.Show("엑셀 파일을 선택하지 않았고, 기본 파일도 없습니다.")
                Return
            End If
        End If
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlBook As Workbook = xlApp.Workbooks.Open(filePath)
        Dim 전체금액합계 As Long = 0

        ListBox1.Items.Add("-------- 지역별 금액 합계 --------")

        For Each sheet As Worksheet In xlBook.Sheets
            xlSheet = sheet
            Dim 지역명 As String = ""
            If Not IsNothing(xlSheet.Range("B4").Value) Then
                지역명 = xlSheet.Range("B4").Value.ToString().Trim()
            End If

            Dim 지역금액합계 As Long = 0

            For i As Integer = 7 To 66
                Dim rawVal = xlSheet.Cells(i, 6).Value ' F열 (금액)

                If Not IsNothing(rawVal) Then
                    Dim textVal As String = rawVal.ToString().Replace(",", "").Trim()

                    Dim parsedVal As Long
                    If Long.TryParse(textVal, parsedVal) Then
                        지역금액합계 += parsedVal
                    Else
                        ListBox1.Items.Add($"{지역명} 시트의 {i}행 F열에 숫자가 아닌 값이 있습니다: {rawVal}")
                    End If
                End If
            Next

            If 지역금액합계 > 0 Then
                ListBox1.Items.Add(지역명 & " : " & 지역금액합계.ToString("N0") & "원")
                전체금액합계 += 지역금액합계
            End If
            ReleaseComObject(xlSheet)
        Next

        ListBox1.Items.Add("------------------------------")
        ListBox1.Items.Add("전체 금액 합계 : " & 전체금액합계.ToString("N0") & "원")

        xlBook.Close(False)
        xlApp.Quit()
        ReleaseComObject(xlBook)
        ReleaseComObject(xlApp)
    End Sub
    Private Sub ListBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.C Then
            If ListBox1.SelectedItem IsNot Nothing Then
                Clipboard.SetText(ListBox1.SelectedItem.ToString())
            End If
        End If
    End Sub

End Class
