<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InloggenVenster
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
      Me.GebruikerLabel = New System.Windows.Forms.Label()
      Me.WachtwoordLabel = New System.Windows.Forms.Label()
      Me.GebruikerVeld = New System.Windows.Forms.TextBox()
      Me.WachtwoordVeld = New System.Windows.Forms.TextBox()
      Me.InloggenKnop = New System.Windows.Forms.Button()
      Me.AnnulerenKnop = New System.Windows.Forms.Button()
      Me.SuspendLayout()
      '
      'GebruikerLabel
      '
      Me.GebruikerLabel.AutoSize = True
      Me.GebruikerLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.GebruikerLabel.Location = New System.Drawing.Point(12, 9)
      Me.GebruikerLabel.Name = "GebruikerLabel"
      Me.GebruikerLabel.Size = New System.Drawing.Size(66, 13)
      Me.GebruikerLabel.TabIndex = 0
      Me.GebruikerLabel.Text = "Gebruiker:"
      Me.GebruikerLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'WachtwoordLabel
      '
      Me.WachtwoordLabel.AutoSize = True
      Me.WachtwoordLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.WachtwoordLabel.Location = New System.Drawing.Point(12, 32)
      Me.WachtwoordLabel.Name = "WachtwoordLabel"
      Me.WachtwoordLabel.Size = New System.Drawing.Size(82, 13)
      Me.WachtwoordLabel.TabIndex = 1
      Me.WachtwoordLabel.Text = "Wachtwoord:"
      Me.WachtwoordLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'GebruikerVeld
      '
      Me.GebruikerVeld.Enabled = False
      Me.GebruikerVeld.Location = New System.Drawing.Point(117, 9)
      Me.GebruikerVeld.MaxLength = 255
      Me.GebruikerVeld.Name = "GebruikerVeld"
      Me.GebruikerVeld.Size = New System.Drawing.Size(161, 20)
      Me.GebruikerVeld.TabIndex = 0
      '
      'WachtwoordVeld
      '
      Me.WachtwoordVeld.Enabled = False
      Me.WachtwoordVeld.Location = New System.Drawing.Point(117, 32)
      Me.WachtwoordVeld.MaxLength = 255
      Me.WachtwoordVeld.Name = "WachtwoordVeld"
      Me.WachtwoordVeld.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
      Me.WachtwoordVeld.Size = New System.Drawing.Size(161, 20)
      Me.WachtwoordVeld.TabIndex = 1
      '
      'InloggenKnop
      '
      Me.InloggenKnop.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.InloggenKnop.Location = New System.Drawing.Point(203, 58)
      Me.InloggenKnop.Name = "InloggenKnop"
      Me.InloggenKnop.Size = New System.Drawing.Size(75, 23)
      Me.InloggenKnop.TabIndex = 3
      Me.InloggenKnop.Text = "&Inloggen"
      Me.InloggenKnop.UseVisualStyleBackColor = True
      '
      'AnnulerenKnop
      '
      Me.AnnulerenKnop.DialogResult = System.Windows.Forms.DialogResult.Cancel
      Me.AnnulerenKnop.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.AnnulerenKnop.Location = New System.Drawing.Point(134, 58)
      Me.AnnulerenKnop.Name = "AnnulerenKnop"
      Me.AnnulerenKnop.Size = New System.Drawing.Size(75, 23)
      Me.AnnulerenKnop.TabIndex = 2
      Me.AnnulerenKnop.Text = "&Annuleren"
      Me.AnnulerenKnop.UseVisualStyleBackColor = True
      '
      'InloggenVenster
      '
      Me.AcceptButton = Me.InloggenKnop
      Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.CancelButton = Me.AnnulerenKnop
      Me.ClientSize = New System.Drawing.Size(285, 88)
      Me.Controls.Add(Me.InloggenKnop)
      Me.Controls.Add(Me.AnnulerenKnop)
      Me.Controls.Add(Me.WachtwoordVeld)
      Me.Controls.Add(Me.GebruikerVeld)
      Me.Controls.Add(Me.WachtwoordLabel)
      Me.Controls.Add(Me.GebruikerLabel)
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
      Me.KeyPreview = True
      Me.MaximizeBox = False
      Me.Name = "InloggenVenster"
      Me.Text = "Inloggen"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub

   Friend WithEvents GebruikerLabel As System.Windows.Forms.Label
   Friend WithEvents WachtwoordLabel As System.Windows.Forms.Label
   Friend WithEvents GebruikerVeld As System.Windows.Forms.TextBox
   Friend WithEvents WachtwoordVeld As System.Windows.Forms.TextBox
   Friend WithEvents InloggenKnop As System.Windows.Forms.Button
   Friend WithEvents AnnulerenKnop As System.Windows.Forms.Button
End Class
