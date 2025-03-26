<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 19
        Me.ListBox1.Location = New System.Drawing.Point(163, 13)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(398, 669)
        Me.ListBox1.TabIndex = 0
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button2.Location = New System.Drawing.Point(737, 470)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(136, 33)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "수량 계산"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button3.Location = New System.Drawing.Point(737, 411)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(160, 33)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "금액 합계 계산"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button1.Location = New System.Drawing.Point(737, 541)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(154, 33)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "액셀 파일 선택"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button4.Location = New System.Drawing.Point(737, 290)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(221, 34)
        Me.Button4.TabIndex = 5
        Me.Button4.Text = "지역별 품목 수량 계산"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button5.Location = New System.Drawing.Point(737, 349)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(221, 35)
        Me.Button5.TabIndex = 6
        Me.Button5.Text = "지역별 품목 금액 계산"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button8.Location = New System.Drawing.Point(737, 157)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(221, 38)
        Me.Button8.TabIndex = 7
        Me.Button8.Text = "전체 품목 금액 계산"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button7.Location = New System.Drawing.Point(737, 225)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(221, 34)
        Me.Button7.TabIndex = 8
        Me.Button7.Text = "전체 품목 수량 계산"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1183, 718)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Button1 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents Button7 As Button
End Class
