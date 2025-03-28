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
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.txtFilePath = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button1.Location = New System.Drawing.Point(118, 131)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(173, 40)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "지역별 수량 합계"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button2.Location = New System.Drawing.Point(335, 131)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(173, 40)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "지역별 금액 합계"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button3.Location = New System.Drawing.Point(543, 131)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(157, 40)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "지역별품목수량"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button4.Location = New System.Drawing.Point(118, 214)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(173, 43)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "지역별품목금액"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button5.Location = New System.Drawing.Point(335, 214)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(173, 43)
        Me.Button5.TabIndex = 4
        Me.Button5.Text = "통합품목수량"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Font = New System.Drawing.Font("굴림", 14.0!)
        Me.Button6.Location = New System.Drawing.Point(543, 214)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(157, 43)
        Me.Button6.TabIndex = 5
        Me.Button6.Text = "통합품목금액"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'txtFilePath
        '
        Me.txtFilePath.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtFilePath.Font = New System.Drawing.Font("굴림", 11.0!)
        Me.txtFilePath.Location = New System.Drawing.Point(54, 40)
        Me.txtFilePath.Multiline = True
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.ReadOnly = True
        Me.txtFilePath.Size = New System.Drawing.Size(100, 21)
        Me.txtFilePath.TabIndex = 6
        Me.txtFilePath.TabStop = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.txtFilePath)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents txtFilePath As TextBox
End Class
