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

    ' OpenFileDialog1이 제대로 선언되어 있어야 합니다.
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Filter 설정
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
        Dim slnPath = System.Windows.Forms.Application.StartupPath
        ' MessageBox.Show(Path.Combine(slnPath, "주문내역 - 복사본.xlsx"))
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
                        ListBox1.Items.Add(String.Format("[{0}] 시트의 {1}행 F열에 숫자가 아닌 값이 있습니다: {2}", 지역명, i, rawVal))
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
    ' 품목 수량 합산을 계산하는 버튼
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
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

        ' 품목별 수량 합산을 위한 Dictionary
        Dim 품목수량 As New Dictionary(Of String, Integer)
        Dim 전체합계 As Integer = 0

        ListBox1.Items.Add("-------- 품목별 수량 합산 --------")

        For Each sheet As Worksheet In xlBook.Sheets
            xlSheet = sheet
            Dim 지역명 As String = ""
            If Not IsNothing(xlSheet.Range("B4").Value) Then
                지역명 = xlSheet.Range("B4").Value.ToString().Trim()
            End If

            ' 각 품목의 수량을 합산
            For i As Integer = 7 To 66
                Dim 품목 As String = xlSheet.Cells(i, 2).Value ' B열 (품목 이름)
                Dim 수량 As Integer = xlSheet.Cells(i, 4).Value ' D열 (수량)

                ' 품목과 수량이 비어있지 않다면 합산
                If Not String.IsNullOrEmpty(품목) AndAlso IsNumeric(수량) Then
                    If 품목수량.ContainsKey(품목) Then
                        품목수량(품목) += 수량
                    Else
                        품목수량.Add(품목, 수량)
                    End If
                End If
            Next

            ' 지역별 품목 수량 합산 결과 출력
            If 품목수량.Count > 0 Then
                ListBox1.Items.Add(지역명 & " 지역의 품목별 수량 합산:")
                For Each 품목 In 품목수량.Keys
                    ListBox1.Items.Add($"- {품목}: {품목수량(품목)}개")
                    전체합계 += 품목수량(품목)
                Next
            End If

            품목수량.Clear() ' 품목수량 초기화
            ReleaseComObject(xlSheet)
        Next

        ' 전체 합계 결과 출력
        ListBox1.Items.Add("------------------------------")
        ListBox1.Items.Add("전체 품목 수량 합계 : " & 전체합계 & "개")

        ' 종료 및 해제
        xlBook.Close(False)
        xlApp.Quit()
        ReleaseComObject(xlBook)
        ReleaseComObject(xlApp)
    End Sub

    ' 품목 금액 합산을 계산하는 버튼
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
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

        ' 품목별 금액 합산을 위한 Dictionary
        Dim 품목금액 As New Dictionary(Of String, Decimal)
        Dim 전체금액합계 As Decimal = 0

        ListBox1.Items.Add("-------- 품목별 금액 합산 --------")

        For Each sheet As Worksheet In xlBook.Sheets
            xlSheet = sheet
            Dim 지역명 As String = ""
            If Not IsNothing(xlSheet.Range("B4").Value) Then
                지역명 = xlSheet.Range("B4").Value.ToString().Trim()
            End If

            ' 각 품목의 금액을 합산
            For i As Integer = 7 To 66
                Dim 품목 As String = xlSheet.Cells(i, 2).Value ' B열 (품목 이름)
                Dim 금액 As Decimal = xlSheet.Cells(i, 6).Value ' F열 (금액)

                ' 품목과 금액이 비어있지 않다면 합산
                If Not String.IsNullOrEmpty(품목) AndAlso IsNumeric(금액) Then
                    If 품목금액.ContainsKey(품목) Then
                        품목금액(품목) += 금액
                    Else
                        품목금액.Add(품목, 금액)
                    End If
                End If
            Next

            ' 지역별 품목 금액 합산 결과 출력
            If 품목금액.Count > 0 Then
                ListBox1.Items.Add(지역명 & " 지역의 품목별 금액 합산:")
                For Each 품목 In 품목금액.Keys
                    ListBox1.Items.Add($"- {품목}: {품목금액(품목).ToString("N0")}원")
                    전체금액합계 += 품목금액(품목)
                Next
            End If

            품목금액.Clear() ' 품목금액 초기화
            ReleaseComObject(xlSheet)
        Next

        ' 전체 금액 결과 출력
        ListBox1.Items.Add("------------------------------")
        ListBox1.Items.Add("전체 품목 금액 합계 : " & 전체금액합계.ToString("N0") & "원")

        ' 종료 및 해제
        xlBook.Close(False)
        xlApp.Quit()
        ReleaseComObject(xlBook)
        ReleaseComObject(xlApp)
    End Sub
    ' 품목별 수량 합산을 계산하는 버튼
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
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

        ' 품목별 수량 합산을 위한 Dictionary
        Dim 품목수량 As New Dictionary(Of String, Integer)
        Dim 전체합계 As Integer = 0

        ListBox1.Items.Add("-------- 통합 품목별 수량 합산 --------")

        For Each sheet As Worksheet In xlBook.Sheets
            xlSheet = sheet

            ' 각 품목의 수량을 합산
            For i As Integer = 7 To 66
                Dim 품목 As String = xlSheet.Cells(i, 2).Value ' B열 (품목 이름)
                Dim 수량 As Integer = xlSheet.Cells(i, 4).Value ' D열 (수량)

                ' 품목과 수량이 비어있지 않다면 합산
                If Not String.IsNullOrEmpty(품목) AndAlso IsNumeric(수량) Then
                    If 품목수량.ContainsKey(품목) Then
                        품목수량(품목) += 수량
                    Else
                        품목수량.Add(품목, 수량)
                    End If
                End If
            Next

        Next

        ' 통합 품목 수량 합산 결과 출력
        If 품목수량.Count > 0 Then
            For Each 품목 In 품목수량.Keys
                ListBox1.Items.Add($"- {품목}: {품목수량(품목)}개")
                전체합계 += 품목수량(품목)
            Next
        End If

        ' 전체 합계 결과 출력
        ListBox1.Items.Add("------------------------------")
        ListBox1.Items.Add("전체 품목 수량 합계 : " & 전체합계 & "개")

        품목수량.Clear() ' 품목수량 초기화

        ' 종료 및 해제
        xlBook.Close(False)
        xlApp.Quit()
        ReleaseComObject(xlBook)
        ReleaseComObject(xlApp)
    End Sub

    ' 품목 금액 합산을 계산하는 버튼
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
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

        ' 품목별 금액 합산을 위한 Dictionary
        Dim 품목금액 As New Dictionary(Of String, Decimal)
        Dim 전체금액합계 As Decimal = 0

        ListBox1.Items.Add("-------- 통합 품목별 금액 합산 --------")

        For Each sheet As Worksheet In xlBook.Sheets
            xlSheet = sheet

            ' 각 품목의 금액을 합산
            For i As Integer = 7 To 66
                Dim 품목 As String = xlSheet.Cells(i, 2).Value ' B열 (품목 이름)
                Dim 금액 As Decimal = xlSheet.Cells(i, 6).Value ' F열 (금액)

                ' 품목과 금액이 비어있지 않다면 합산
                If Not String.IsNullOrEmpty(품목) AndAlso IsNumeric(금액) Then
                    If 품목금액.ContainsKey(품목) Then
                        품목금액(품목) += 금액
                    Else
                        품목금액.Add(품목, 금액)
                    End If
                End If
            Next
        Next

        ' 통합 품목 금액 합산 결과 출력
        If 품목금액.Count > 0 Then
            For Each 품목 In 품목금액.Keys
                ListBox1.Items.Add($"- {품목}: {품목금액(품목).ToString("N0")}원")
                전체금액합계 += 품목금액(품목)
            Next
        End If

        ' 전체 금액 결과 출력
        ListBox1.Items.Add("------------------------------")
        ListBox1.Items.Add("전체 품목 금액 합계 : " & 전체금액합계.ToString("N0") & "원")

        품목금액.Clear() ' 품목금액 초기화

        ' 종료 및 해제
        xlBook.Close(False)
        xlApp.Quit()
        ReleaseComObject(xlBook)
        ReleaseComObject(xlApp)
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ListBox1의 선택 모드를 여러 항목 선택 가능하도록 설정
        ListBox1.SelectionMode = SelectionMode.MultiExtended
    End Sub

    ' 여러 줄 복사
    Private Sub ListBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.C Then
            ' 선택된 항목들을 결합하여 하나의 문자열로 만듦
            Dim selectedItems As New Text.StringBuilder()

            For Each item As Object In ListBox1.SelectedItems
                selectedItems.AppendLine(item.ToString())
            Next

            ' 결합된 텍스트를 클립보드에 복사
            If selectedItems.Length > 0 Then
                Clipboard.SetText(selectedItems.ToString())
            End If
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button8.Click

    End Sub
End Class
