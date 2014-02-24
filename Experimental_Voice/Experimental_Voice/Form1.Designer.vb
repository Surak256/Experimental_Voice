<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMonitor
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtOut = New System.Windows.Forms.TextBox
        Me.btnNewCommand = New System.Windows.Forms.Button
        Me.btnRemove = New System.Windows.Forms.Button
        Me.btnDisplay = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'txtOut
        '
        Me.txtOut.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOut.Location = New System.Drawing.Point(13, 13)
        Me.txtOut.Multiline = True
        Me.txtOut.Name = "txtOut"
        Me.txtOut.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOut.Size = New System.Drawing.Size(353, 221)
        Me.txtOut.TabIndex = 0
        '
        'btnNewCommand
        '
        Me.btnNewCommand.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnNewCommand.Location = New System.Drawing.Point(12, 240)
        Me.btnNewCommand.Name = "btnNewCommand"
        Me.btnNewCommand.Size = New System.Drawing.Size(106, 23)
        Me.btnNewCommand.TabIndex = 1
        Me.btnNewCommand.Text = "Add command"
        Me.btnNewCommand.UseVisualStyleBackColor = True
        '
        'btnRemove
        '
        Me.btnRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRemove.Location = New System.Drawing.Point(12, 270)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(106, 23)
        Me.btnRemove.TabIndex = 2
        Me.btnRemove.Text = "Remove Command"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'btnDisplay
        '
        Me.btnDisplay.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDisplay.Location = New System.Drawing.Point(410, 370)
        Me.btnDisplay.Name = "btnDisplay"
        Me.btnDisplay.Size = New System.Drawing.Size(96, 23)
        Me.btnDisplay.TabIndex = 3
        Me.btnDisplay.Text = "Display grammar"
        Me.btnDisplay.UseVisualStyleBackColor = True
        '
        'frmMonitor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(378, 313)
        Me.Controls.Add(Me.btnRemove)
        Me.Controls.Add(Me.btnNewCommand)
        Me.Controls.Add(Me.txtOut)
        Me.Name = "frmMonitor"
        Me.Text = "Monitor"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtOut As System.Windows.Forms.TextBox
    Friend WithEvents btnNewCommand As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnDisplay As System.Windows.Forms.Button

End Class
