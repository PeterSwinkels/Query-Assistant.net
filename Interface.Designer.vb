<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InterfaceVenster
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
      Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(InterfaceVenster))
      Me.QueryFrame = New System.Windows.Forms.GroupBox()
      Me.QueryOpenenKnop = New System.Windows.Forms.Button()
      Me.QuerySelecterenKnop = New System.Windows.Forms.Button()
      Me.QueryUitvoerenKnop = New System.Windows.Forms.Button()
      Me.QueryPadVeld = New System.Windows.Forms.TextBox()
      Me.QueryLabel = New System.Windows.Forms.Label()
      Me.ParametersFrame = New System.Windows.Forms.GroupBox()
      Me.ParameterFrameSchuifBalk = New System.Windows.Forms.VScrollBar()
      Me.ParameterVeldHouder = New System.Windows.Forms.Panel()
      Me.ParameterVelden00000 = New System.Windows.Forms.TextBox()
      Me.ParameterLabels00000 = New System.Windows.Forms.Label()
      Me.ExportFrame = New System.Windows.Forms.GroupBox()
      Me.ExportPadSelecterenKnop = New System.Windows.Forms.Button()
      Me.ResultaatExporterenKnop = New System.Windows.Forms.Button()
      Me.MaakEMailMetExportBijgevoegdVeld = New System.Windows.Forms.CheckBox()
      Me.OpenResultaatNaExportVeld = New System.Windows.Forms.CheckBox()
      Me.AutomatischResultaatExporterenVeld = New System.Windows.Forms.CheckBox()
      Me.ExportPadVeld = New System.Windows.Forms.TextBox()
      Me.ExporteerResultaatNaarLabel = New System.Windows.Forms.Label()
      Me.ResultaatFrame = New System.Windows.Forms.GroupBox()
      Me.QueryResultaatVeld = New System.Windows.Forms.TextBox()
      Me.MenuBalk = New System.Windows.Forms.MenuStrip()
      Me.ProgrammaHoofdMenu = New System.Windows.Forms.ToolStripMenuItem()
      Me.InformatieMenu = New System.Windows.Forms.ToolStripMenuItem()
      Me.SluitenMenu = New System.Windows.Forms.ToolStripMenuItem()
      Me.StatusVeld = New System.Windows.Forms.TextBox()
      Me.QueryFrame.SuspendLayout()
      Me.ParametersFrame.SuspendLayout()
      Me.ParameterVeldHouder.SuspendLayout()
      Me.ExportFrame.SuspendLayout()
      Me.ResultaatFrame.SuspendLayout()
      Me.MenuBalk.SuspendLayout()
      Me.SuspendLayout()
      '
      'QueryFrame
      '
      Me.QueryFrame.Controls.Add(Me.QueryOpenenKnop)
      Me.QueryFrame.Controls.Add(Me.QuerySelecterenKnop)
      Me.QueryFrame.Controls.Add(Me.QueryUitvoerenKnop)
      Me.QueryFrame.Controls.Add(Me.QueryPadVeld)
      Me.QueryFrame.Controls.Add(Me.QueryLabel)
      Me.QueryFrame.Controls.Add(Me.ParametersFrame)
      Me.QueryFrame.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.QueryFrame.Location = New System.Drawing.Point(12, 27)
      Me.QueryFrame.Name = "QueryFrame"
      Me.QueryFrame.Size = New System.Drawing.Size(281, 177)
      Me.QueryFrame.TabIndex = 0
      Me.QueryFrame.TabStop = False
      Me.QueryFrame.Text = "Query"
      '
      'QueryOpenenKnop
      '
      Me.QueryOpenenKnop.Image = Global.Quey_Assistent.NET.My.Resources.Resources.Diskette
      Me.QueryOpenenKnop.Location = New System.Drawing.Point(250, 19)
      Me.QueryOpenenKnop.Name = "QueryOpenenKnop"
      Me.QueryOpenenKnop.Size = New System.Drawing.Size(25, 23)
      Me.QueryOpenenKnop.TabIndex = 2
      Me.QueryOpenenKnop.UseVisualStyleBackColor = True
      '
      'QuerySelecterenKnop
      '
      Me.QuerySelecterenKnop.Image = Global.Quey_Assistent.NET.My.Resources.Resources.Map
      Me.QuerySelecterenKnop.Location = New System.Drawing.Point(223, 20)
      Me.QuerySelecterenKnop.Name = "QuerySelecterenKnop"
      Me.QuerySelecterenKnop.Size = New System.Drawing.Size(25, 23)
      Me.QuerySelecterenKnop.TabIndex = 1
      Me.QuerySelecterenKnop.UseVisualStyleBackColor = True
      '
      'QueryUitvoerenKnop
      '
      Me.QueryUitvoerenKnop.Location = New System.Drawing.Point(143, 141)
      Me.QueryUitvoerenKnop.Name = "QueryUitvoerenKnop"
      Me.QueryUitvoerenKnop.Size = New System.Drawing.Size(120, 23)
      Me.QueryUitvoerenKnop.TabIndex = 5
      Me.QueryUitvoerenKnop.Text = "Query &Uitvoeren"
      Me.QueryUitvoerenKnop.UseVisualStyleBackColor = True
      '
      'QueryPadVeld
      '
      Me.QueryPadVeld.AllowDrop = True
      Me.QueryPadVeld.Location = New System.Drawing.Point(56, 23)
      Me.QueryPadVeld.Name = "QueryPadVeld"
      Me.QueryPadVeld.Size = New System.Drawing.Size(161, 20)
      Me.QueryPadVeld.TabIndex = 0
      '
      'QueryLabel
      '
      Me.QueryLabel.AutoSize = True
      Me.QueryLabel.Location = New System.Drawing.Point(6, 26)
      Me.QueryLabel.Name = "QueryLabel"
      Me.QueryLabel.Size = New System.Drawing.Size(44, 13)
      Me.QueryLabel.TabIndex = 0
      Me.QueryLabel.Text = "Query:"
      '
      'ParametersFrame
      '
      Me.ParametersFrame.Controls.Add(Me.ParameterFrameSchuifBalk)
      Me.ParametersFrame.Controls.Add(Me.ParameterVeldHouder)
      Me.ParametersFrame.Location = New System.Drawing.Point(3, 49)
      Me.ParametersFrame.Name = "ParametersFrame"
      Me.ParametersFrame.Size = New System.Drawing.Size(279, 81)
      Me.ParametersFrame.TabIndex = 3
      Me.ParametersFrame.TabStop = False
      Me.ParametersFrame.Text = "Parameters"
      '
      'ParameterFrameSchuifBalk
      '
      Me.ParameterFrameSchuifBalk.Location = New System.Drawing.Point(261, 10)
      Me.ParameterFrameSchuifBalk.Name = "ParameterFrameSchuifBalk"
      Me.ParameterFrameSchuifBalk.Size = New System.Drawing.Size(17, 71)
      Me.ParameterFrameSchuifBalk.TabIndex = 13
      '
      'ParameterVeldHouder
      '
      Me.ParameterVeldHouder.AutoScroll = True
      Me.ParameterVeldHouder.Controls.Add(Me.ParameterVelden00000)
      Me.ParameterVeldHouder.Controls.Add(Me.ParameterLabels00000)
      Me.ParameterVeldHouder.Location = New System.Drawing.Point(6, 16)
      Me.ParameterVeldHouder.Name = "ParameterVeldHouder"
      Me.ParameterVeldHouder.Size = New System.Drawing.Size(238, 62)
      Me.ParameterVeldHouder.TabIndex = 6
      '
      'ParameterVelden00000
      '
      Me.ParameterVelden00000.Location = New System.Drawing.Point(87, 3)
      Me.ParameterVelden00000.Name = "ParameterVelden00000"
      Me.ParameterVelden00000.Size = New System.Drawing.Size(139, 20)
      Me.ParameterVelden00000.TabIndex = 3
      '
      'ParameterLabels00000
      '
      Me.ParameterLabels00000.Location = New System.Drawing.Point(13, 6)
      Me.ParameterLabels00000.Name = "ParameterLabels00000"
      Me.ParameterLabels00000.Size = New System.Drawing.Size(68, 13)
      Me.ParameterLabels00000.TabIndex = 1
      Me.ParameterLabels00000.Text = "Parameter:"
      Me.ParameterLabels00000.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'ExportFrame
      '
      Me.ExportFrame.Controls.Add(Me.ExportPadSelecterenKnop)
      Me.ExportFrame.Controls.Add(Me.ResultaatExporterenKnop)
      Me.ExportFrame.Controls.Add(Me.MaakEMailMetExportBijgevoegdVeld)
      Me.ExportFrame.Controls.Add(Me.OpenResultaatNaExportVeld)
      Me.ExportFrame.Controls.Add(Me.AutomatischResultaatExporterenVeld)
      Me.ExportFrame.Controls.Add(Me.ExportPadVeld)
      Me.ExportFrame.Controls.Add(Me.ExporteerResultaatNaarLabel)
      Me.ExportFrame.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ExportFrame.Location = New System.Drawing.Point(300, 27)
      Me.ExportFrame.Name = "ExportFrame"
      Me.ExportFrame.Size = New System.Drawing.Size(313, 177)
      Me.ExportFrame.TabIndex = 1
      Me.ExportFrame.TabStop = False
      Me.ExportFrame.Text = "Export"
      '
      'ExportPadSelecterenKnop
      '
      Me.ExportPadSelecterenKnop.Image = Global.Quey_Assistent.NET.My.Resources.Resources.Map
      Me.ExportPadSelecterenKnop.Location = New System.Drawing.Point(282, 36)
      Me.ExportPadSelecterenKnop.Name = "ExportPadSelecterenKnop"
      Me.ExportPadSelecterenKnop.Size = New System.Drawing.Size(25, 23)
      Me.ExportPadSelecterenKnop.TabIndex = 7
      Me.ExportPadSelecterenKnop.UseVisualStyleBackColor = True
      '
      'ResultaatExporterenKnop
      '
      Me.ResultaatExporterenKnop.Location = New System.Drawing.Point(173, 141)
      Me.ResultaatExporterenKnop.Name = "ResultaatExporterenKnop"
      Me.ResultaatExporterenKnop.Size = New System.Drawing.Size(134, 23)
      Me.ResultaatExporterenKnop.TabIndex = 11
      Me.ResultaatExporterenKnop.Text = "Resultaat &Exporteren"
      Me.ResultaatExporterenKnop.UseVisualStyleBackColor = True
      '
      'MaakEMailMetExportBijgevoegdVeld
      '
      Me.MaakEMailMetExportBijgevoegdVeld.AutoSize = True
      Me.MaakEMailMetExportBijgevoegdVeld.Location = New System.Drawing.Point(9, 99)
      Me.MaakEMailMetExportBijgevoegdVeld.Name = "MaakEMailMetExportBijgevoegdVeld"
      Me.MaakEMailMetExportBijgevoegdVeld.Size = New System.Drawing.Size(227, 17)
      Me.MaakEMailMetExportBijgevoegdVeld.TabIndex = 10
      Me.MaakEMailMetExportBijgevoegdVeld.Text = "Maak e-&mail met export bijgevoegd."
      Me.MaakEMailMetExportBijgevoegdVeld.UseVisualStyleBackColor = True
      '
      'OpenResultaatNaExportVeld
      '
      Me.OpenResultaatNaExportVeld.AutoSize = True
      Me.OpenResultaatNaExportVeld.Location = New System.Drawing.Point(9, 82)
      Me.OpenResultaatNaExportVeld.Name = "OpenResultaatNaExportVeld"
      Me.OpenResultaatNaExportVeld.Size = New System.Drawing.Size(170, 17)
      Me.OpenResultaatNaExportVeld.TabIndex = 9
      Me.OpenResultaatNaExportVeld.Text = "&Open resultaat na export."
      Me.OpenResultaatNaExportVeld.UseVisualStyleBackColor = True
      '
      'AutomatischResultaatExporterenVeld
      '
      Me.AutomatischResultaatExporterenVeld.AutoSize = True
      Me.AutomatischResultaatExporterenVeld.Location = New System.Drawing.Point(9, 65)
      Me.AutomatischResultaatExporterenVeld.Name = "AutomatischResultaatExporterenVeld"
      Me.AutomatischResultaatExporterenVeld.Size = New System.Drawing.Size(269, 17)
      Me.AutomatischResultaatExporterenVeld.TabIndex = 8
      Me.AutomatischResultaatExporterenVeld.Text = "&Automatisch resultaat exporteren na query."
      Me.AutomatischResultaatExporterenVeld.UseVisualStyleBackColor = True
      '
      'ExportPadVeld
      '
      Me.ExportPadVeld.Location = New System.Drawing.Point(9, 39)
      Me.ExportPadVeld.Name = "ExportPadVeld"
      Me.ExportPadVeld.Size = New System.Drawing.Size(269, 20)
      Me.ExportPadVeld.TabIndex = 6
      '
      'ExporteerResultaatNaarLabel
      '
      Me.ExporteerResultaatNaarLabel.AutoSize = True
      Me.ExporteerResultaatNaarLabel.Location = New System.Drawing.Point(6, 23)
      Me.ExporteerResultaatNaarLabel.Name = "ExporteerResultaatNaarLabel"
      Me.ExporteerResultaatNaarLabel.Size = New System.Drawing.Size(147, 13)
      Me.ExporteerResultaatNaarLabel.TabIndex = 1
      Me.ExporteerResultaatNaarLabel.Text = "Exporteer resultaat naar:"
      '
      'ResultaatFrame
      '
      Me.ResultaatFrame.Controls.Add(Me.QueryResultaatVeld)
      Me.ResultaatFrame.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ResultaatFrame.Location = New System.Drawing.Point(12, 218)
      Me.ResultaatFrame.Name = "ResultaatFrame"
      Me.ResultaatFrame.Size = New System.Drawing.Size(601, 289)
      Me.ResultaatFrame.TabIndex = 2
      Me.ResultaatFrame.TabStop = False
      Me.ResultaatFrame.Text = "Resultaat"
      '
      'QueryResultaatVeld
      '
      Me.QueryResultaatVeld.BackColor = System.Drawing.SystemColors.ControlLightLight
      Me.QueryResultaatVeld.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.QueryResultaatVeld.Location = New System.Drawing.Point(9, 19)
      Me.QueryResultaatVeld.Multiline = True
      Me.QueryResultaatVeld.Name = "QueryResultaatVeld"
      Me.QueryResultaatVeld.ReadOnly = True
      Me.QueryResultaatVeld.ScrollBars = System.Windows.Forms.ScrollBars.Both
      Me.QueryResultaatVeld.Size = New System.Drawing.Size(586, 272)
      Me.QueryResultaatVeld.TabIndex = 12
      Me.QueryResultaatVeld.WordWrap = False
      '
      'MenuBalk
      '
      Me.MenuBalk.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProgrammaHoofdMenu})
      Me.MenuBalk.Location = New System.Drawing.Point(0, 0)
      Me.MenuBalk.Name = "MenuBalk"
      Me.MenuBalk.Size = New System.Drawing.Size(625, 24)
      Me.MenuBalk.TabIndex = 14
      '
      'ProgrammaHoofdMenu
      '
      Me.ProgrammaHoofdMenu.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.InformatieMenu, Me.SluitenMenu})
      Me.ProgrammaHoofdMenu.Name = "ProgrammaHoofdMenu"
      Me.ProgrammaHoofdMenu.Size = New System.Drawing.Size(82, 20)
      Me.ProgrammaHoofdMenu.Text = "&Programma"
      '
      'InformatieMenu
      '
      Me.InformatieMenu.Name = "InformatieMenu"
      Me.InformatieMenu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.I), System.Windows.Forms.Keys)
      Me.InformatieMenu.Size = New System.Drawing.Size(166, 22)
      Me.InformatieMenu.Text = "&Informatie"
      '
      'SluitenMenu
      '
      Me.SluitenMenu.Name = "SluitenMenu"
      Me.SluitenMenu.ShortcutKeyDisplayString = "Ctrl+S"
      Me.SluitenMenu.Size = New System.Drawing.Size(166, 22)
      Me.SluitenMenu.Text = "&Sluiten"
      '
      'StatusVeld
      '
      Me.StatusVeld.BackColor = System.Drawing.SystemColors.Control
      Me.StatusVeld.Location = New System.Drawing.Point(12, 515)
      Me.StatusVeld.Multiline = True
      Me.StatusVeld.Name = "StatusVeld"
      Me.StatusVeld.ReadOnly = True
      Me.StatusVeld.ScrollBars = System.Windows.Forms.ScrollBars.Both
      Me.StatusVeld.Size = New System.Drawing.Size(595, 53)
      Me.StatusVeld.TabIndex = 13
      Me.StatusVeld.WordWrap = False
      '
      'InterfaceVenster
      '
      Me.AcceptButton = Me.QueryUitvoerenKnop
      Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.ClientSize = New System.Drawing.Size(625, 580)
      Me.Controls.Add(Me.StatusVeld)
      Me.Controls.Add(Me.ExportFrame)
      Me.Controls.Add(Me.QueryFrame)
      Me.Controls.Add(Me.MenuBalk)
      Me.Controls.Add(Me.ResultaatFrame)
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.KeyPreview = True
      Me.MaximizeBox = False
      Me.Name = "InterfaceVenster"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.QueryFrame.ResumeLayout(False)
      Me.QueryFrame.PerformLayout()
      Me.ParametersFrame.ResumeLayout(False)
      Me.ParameterVeldHouder.ResumeLayout(False)
      Me.ParameterVeldHouder.PerformLayout()
      Me.ExportFrame.ResumeLayout(False)
      Me.ExportFrame.PerformLayout()
      Me.ResultaatFrame.ResumeLayout(False)
      Me.ResultaatFrame.PerformLayout()
      Me.MenuBalk.ResumeLayout(False)
      Me.MenuBalk.PerformLayout()
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub

   Friend WithEvents QueryFrame As System.Windows.Forms.GroupBox
   Friend WithEvents QuerySelecterenKnop As System.Windows.Forms.Button
   Friend WithEvents ParametersFrame As System.Windows.Forms.GroupBox
   Friend WithEvents ParameterVeldHouder As System.Windows.Forms.Panel
   Friend WithEvents QueryUitvoerenKnop As System.Windows.Forms.Button
   Friend WithEvents QueryPadVeld As System.Windows.Forms.TextBox
   Friend WithEvents QueryLabel As System.Windows.Forms.Label
   Friend WithEvents ExportFrame As System.Windows.Forms.GroupBox
   Friend WithEvents ResultaatFrame As System.Windows.Forms.GroupBox
   Friend WithEvents QueryOpenenKnop As System.Windows.Forms.Button
   Friend WithEvents ParameterVelden00000 As System.Windows.Forms.TextBox
   Friend WithEvents ParameterLabels00000 As System.Windows.Forms.Label
   Friend WithEvents ResultaatExporterenKnop As System.Windows.Forms.Button
   Friend WithEvents MaakEMailMetExportBijgevoegdVeld As System.Windows.Forms.CheckBox
   Friend WithEvents OpenResultaatNaExportVeld As System.Windows.Forms.CheckBox
   Friend WithEvents AutomatischResultaatExporterenVeld As System.Windows.Forms.CheckBox
   Friend WithEvents ExportPadVeld As System.Windows.Forms.TextBox
   Friend WithEvents ExporteerResultaatNaarLabel As System.Windows.Forms.Label
   Friend WithEvents QueryResultaatVeld As System.Windows.Forms.TextBox
   Friend WithEvents ExportPadSelecterenKnop As System.Windows.Forms.Button
   Friend WithEvents MenuBalk As System.Windows.Forms.MenuStrip
   Friend WithEvents ProgrammaHoofdMenu As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents InformatieMenu As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents SluitenMenu As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents StatusVeld As System.Windows.Forms.TextBox
   Friend WithEvents ParameterFrameSchuifBalk As System.Windows.Forms.VScrollBar
End Class
