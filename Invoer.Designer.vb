<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InvoerVenster
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
        Me.OKKnop = New System.Windows.Forms.Button()
        Me.AnnulerenKnop = New System.Windows.Forms.Button()
        Me.TekstVeld = New System.Windows.Forms.TextBox()
        Me.Paneel = New System.Windows.Forms.Panel()
        Me.PromptLabel = New System.Windows.Forms.Label()
        Me.Paneel.SuspendLayout()
        Me.SuspendLayout()
        '
        'OKKnop
        '
        Me.OKKnop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OKKnop.Location = New System.Drawing.Point(260, 10)
        Me.OKKnop.Name = "OKKnop"
        Me.OKKnop.Size = New System.Drawing.Size(67, 23)
        Me.OKKnop.TabIndex = 1
        Me.OKKnop.TabStop = False
        Me.OKKnop.Text = "&OK"
        '
        'AnnulerenKnop
        '
        Me.AnnulerenKnop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AnnulerenKnop.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.AnnulerenKnop.Location = New System.Drawing.Point(260, 39)
        Me.AnnulerenKnop.Name = "AnnulerenKnop"
        Me.AnnulerenKnop.Size = New System.Drawing.Size(67, 23)
        Me.AnnulerenKnop.TabIndex = 2
        Me.AnnulerenKnop.TabStop = False
        Me.AnnulerenKnop.Text = "&Annuleren"
        '
        'TekstVeld
        '
        Me.TekstVeld.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.TekstVeld.Location = New System.Drawing.Point(12, 68)
        Me.TekstVeld.Name = "TekstVeld"
        Me.TekstVeld.Size = New System.Drawing.Size(315, 20)
        Me.TekstVeld.TabIndex = 0
        '
        'Paneel
        '
        Me.Paneel.Controls.Add(Me.PromptLabel)
        Me.Paneel.Location = New System.Drawing.Point(12, 10)
        Me.Paneel.Name = "Paneel"
        Me.Paneel.Size = New System.Drawing.Size(242, 51)
        Me.Paneel.TabIndex = 4
        '
        'PromptLabel
        '
        Me.PromptLabel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PromptLabel.AutoEllipsis = True
        Me.PromptLabel.AutoSize = True
        Me.PromptLabel.Location = New System.Drawing.Point(0, 0)
        Me.PromptLabel.Name = "PromptLabel"
        Me.PromptLabel.Size = New System.Drawing.Size(0, 13)
        Me.PromptLabel.TabIndex = 5
        '
        'InvoerVenster
        '
        Me.AcceptButton = Me.OKKnop
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.AnnulerenKnop
        Me.ClientSize = New System.Drawing.Size(339, 98)
        Me.Controls.Add(Me.Paneel)
        Me.Controls.Add(Me.TekstVeld)
        Me.Controls.Add(Me.AnnulerenKnop)
        Me.Controls.Add(Me.OKKnop)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "InvoerVenster"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Paneel.ResumeLayout(False)
        Me.Paneel.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OKKnop As System.Windows.Forms.Button
   Friend WithEvents AnnulerenKnop As System.Windows.Forms.Button
   Friend WithEvents TekstVeld As System.Windows.Forms.TextBox
   Friend WithEvents Paneel As System.Windows.Forms.Panel
   Friend WithEvents PromptLabel As System.Windows.Forms.Label
End Class
