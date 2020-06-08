'De instellingen en geimporteerde namespaces van deze module.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Environment
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms

'Deze module bevat het interfacevenster van dit programma.
Public Class InterfaceVenster
   Private Geactiveerd As Boolean = False  'Geeft aan of het venster al een keer geactiveerd is.
   Private WithEvents Tip As New ToolTip   'Geeft tips met betrekking tot de knoppen en velden in dit venster weer.

   'Deze procedure stelt dit venster in wanneer het wordt geopend.
   Public Sub New()
      Try
         InitializeComponent()

         HuidigInterfaceVenster = Me

         If Not BATCH_MODUS_ACTIEF() AndAlso Not OpdrachtRegelParameters().QueryPad = Nothing Then Query(OpdrachtRegelParameters().QueryPad)

         ResetVenster()
         ToonStatus(, NieuwVeld:=StatusVeld)

         With Instellingen()
            ExportPadVeld.Text = .ExportStandaardPad

            AutomatischResultaatExporterenVeld.Checked = False
            MaakEMailMetExportBijgevoegdVeld.Checked = False
            OpenResultaatNaExportVeld.Checked = False

            If .ExportAutoOpenen Then OpenResultaatNaExportVeld.Checked = True
            If Not .ExportStandaardPad = Nothing Then AutomatischResultaatExporterenVeld.Checked = True
            If Not (.ExportOntvanger = Nothing AndAlso .ExportCCOntvanger = Nothing) Then MaakEMailMetExportBijgevoegdVeld.Checked = True
         End With
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure geeft de opdracht om de gebruiker te verzoeken een export pad op te geven.
   Private Sub ExportPadSelecterenKnop_Click(sender As Object, e As EventArgs) Handles ExportPadSelecterenKnop.Click
      Try
         ExportPadVeld.Text = VraagExportPad(ExportPadVeld.Text)
         ExportPadVeld.SelectionStart = 0
         If Not ExportPadVeld.Text = Nothing Then ExportPadVeld.SelectionStart = ExportPadVeld.Text.Length - 1
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure geeft opdracht om het queryresultaat te exporteren.
   Private Sub GeefExportOpdracht()
      Try
         Dim EMail As EMailClass = Nothing
         Dim ExportPad As String = ExportPadVeld.Text

         If ExportPad = Nothing Then
            MessageBox.Show("Geen export pad opgegeven.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
         ElseIf Me.Visible Then
            Me.Cursor = Cursors.WaitCursor
            ResultaatExporterenKnop.Enabled = False
            ToonStatus($"Bezig met het exporteren van het queryresultaat...{NewLine}")

            ExportPad = Path.GetFullPath(VerwijderAanhalingsTekens(VervangSymbolen(ExportPad).Trim()))

            If Directory.Exists(Path.GetDirectoryName(ExportPad)) Then
               If ExporteerResultaat(ExportPad) Then
                  If File.Exists(ExportPad) Then
                     If OpenResultaatNaExportVeld.Checked Then
                        ToonStatus($"De export wordt automatisch geopend...{NewLine}")
                        Process.Start(New ProcessStartInfo With {.CreateNoWindow = False, .FileName = ExportPad, .ErrorDialog = True, .UseShellExecute = True, .WindowStyle = ProcessWindowStyle.Normal})
                     End If
                     If MaakEMailMetExportBijgevoegdVeld.Checked Then
                        ToonStatus($"Bezig met het maken van de e-mail met de export...{NewLine}")
                        EMail = New EMailClass
                        EMail.VoegQueryResultatenToe(New List(Of String)({ExportPad}))
                        EMail = Nothing
                     End If
                  End If
                  ToonStatus($"Exporteren gereed.{NewLine}")
               Else
                  ToonStatus($"Export afgebroken.{NewLine}")
               End If
            Else
               MessageBox.Show($"Ongeldig export pad.{NewLine}Huidig pad: ""{Directory.GetCurrentDirectory()}""", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
               ToonStatus($"Ongeldig export pad.{NewLine}")
            End If
         End If

         ResultaatExporterenKnop.Enabled = True
         Me.Cursor = Cursors.Default
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure geeft de opdracht om de geselecteerde query met de opgegeven parameters uit te voeren.
   Private Sub GeefQueryOpdracht()
      Try
         If Not Query().Code = Nothing Then
            QueryUitvoerenKnop.Enabled = False
            If ParametersGeldig(VraagParameterVeldHouderObjectenOp("ParameterVelden")) Then
               Me.Cursor = Cursors.WaitCursor

               ToonStatus($"Bezig met het uitvoeren van de query...{NewLine}")

               QueryResultaten(, ResultatenVerwijderen:=True)
               VoerQueryUit(Query().Code)

               If VERBINDING_GEOPEND(Verbinding()) Then
                  ToonQueryResultaat(QueryResultaatVeld, ResultaatIndex:=0)

                  If Verbinding().Errors.Count = 0 Then
                     If AutomatischResultaatExporterenVeld.Checked Then GeefExportOpdracht()
                  Else
                     ToonStatus(FoutenLijstTekst(Verbinding().Errors))
                  End If

                  Verbinding(, , Reset:=True)
               End If
            End If
         End If
      Catch
         HandelFoutAf()
      Finally
         QueryUitvoerenKnop.Enabled = ((Not (Query().Pad = Nothing)) AndAlso VERBINDING_GEOPEND(Verbinding()))
         Me.Cursor = Cursors.Default

         If (Instellingen().QueryAutoSluiten) OrElse (Not VerwerkSessieLijst() = Nothing) Then Me.Close()
      End Try
   End Sub

   'Deze procedure geeft de opdracht om programmainformatie te tonen.
   Private Sub InformatieMenu_Click(sender As Object, e As EventArgs) Handles InformatieMenu.Click
      Try
         ToonProgrammaInformatie()
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure geeft de opdracht om een eventuele bij het starten van dit programma geladen query te tonen.
   Private Sub InterfaceVenster_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
      Try
         If Not Geactiveerd Then
            Geactiveerd = True

            If BATCH_MODUS_ACTIEF() Then
               PasVensterAan()
            Else
               If Not OPDRACHT_REGEL.Trim() = Nothing Then ToonStatus($"Opdrachtregel: {OPDRACHT_REGEL}{NewLine}")
               If Not VerwerkSessieLijst() = Nothing Then ToonStatus($"Sessie lijst: {VerwerkSessieLijst()}{NewLine}")
               TOON_VERBINDINGS_STATUS()
               If Not Query().Pad = Nothing Then ToonQuery(Query().Pad)
            End If
         End If
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure sluit dit venster.
   Private Sub InterfaceVenster_Closed(sender As Object, e As EventArgs) Handles MyBase.Closed
      Try
         HuidigInterfaceVenster = Nothing
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure sluit dit programma na bevestiging van de gebruiker.
   Private Sub Interface_Closing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing
      Try
         Dim Keuze As New DialogResult

         With Instellingen()
            If Not .QueryAutoSluiten Then
               If ((InteractieveBatchAfbreken AndAlso InteractieveBatchModusActief) OrElse Not BATCH_MODUS_ACTIEF()) OrElse (Instellingen().QueryAutoSluiten) OrElse (Not VerwerkSessieLijst() = Nothing) Then
                  Keuze = MessageBox.Show("Dit programma sluiten?", My.Application.Info.Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
                  Select Case Keuze
                     Case DialogResult.No
                        e.Cancel = True
                     Case DialogResult.Yes
                        e.Cancel = False
                        If Not VerwerkSessieLijst() = Nothing Then SessiesAfbreken = True
                  End Select
               End If
            End If
         End With
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure verplaatst de parametervelden wanneer de knop op de schuifbalk verschoven wordt.
   Private Sub ParameterFrameSchuifBalk_Scroll(sender As Object, e As EventArgs) Handles ParameterFrameSchuifBalk.Scroll
      Try
         Dim ParameterLabel As Label = Nothing
         Dim ParameterVeld As TextBox = Nothing
         Dim Rij As Integer = 0

         For ParameterIndex As Integer = 0 To VraagParameterVeldHouderObjectenOp("ParameterVelden").Count - 1
            ParameterLabel = DirectCast(Me.ParameterVeldHouder.Controls($"ParameterLabels{ParameterIndex:D5}"), Label)
            ParameterVeld = DirectCast(Me.ParameterVeldHouder.Controls($"ParameterVelden{ParameterIndex:D5}"), TextBox)
            ParameterLabel.Top = (Rij - ParameterFrameSchuifBalk.Value) * 24
            ParameterVeld.Top = ParameterLabel.Top

            If Not QueryParameters()(ParameterIndex).ParameterNaam = Nothing Then Rij += 1
         Next ParameterIndex
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure selecteert de inhoud van het geactiveerde parameterveld.
   Private Sub ParameterVelden_GotFocus(sender As Object, e As EventArgs)
      Try
         With DirectCast(sender, TextBox)
            If .Top - .Height < 0 OrElse .Top > ParameterVeldHouder.Height Then VerschuifBalk(Integer.Parse(.Name.Substring(.Name.Length - 5)))
            .SelectionStart = 0
            .SelectionLength = .Text.Length
         End With
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure filtert de toetsaanslagen van de gebruiker in een parameterveld.
   Private Sub ParameterVelden_KeyDown(sender As Object, e As KeyEventArgs)
      Try
         Dim ParameterIndex As Integer = Nothing

         With DirectCast(sender, TextBox)
            Integer.TryParse(.Name.Substring(.Name.Length - 6), ParameterIndex)

            If e.KeyCode = Keys.V AndAlso e.Modifiers = Keys.Control Then
               If .SelectionLength = 0 Then
                  .Text = Clipboard.GetText(TextDataFormat.Text)
               Else
                  .SelectedText = Clipboard.GetText(TextDataFormat.Text)
               End If

               e.SuppressKeyPress = True
            ElseIf e.KeyCode = Keys.X AndAlso e.Modifiers = Keys.Control Then
               If .SelectionLength = 0 Then
                  Clipboard.SetText(.Text, TextDataFormat.Text)
                  .Text = QueryParameters()(ParameterIndex).Masker
               Else
                  Clipboard.SetText(.SelectedText, TextDataFormat.Text)
                  .SelectedText = QueryParameters()(ParameterIndex).Masker
               End If

               e.SuppressKeyPress = True
            End If
         End With
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure verwerkt de invoer van de gebruiker in een parameterveld.
   Private Sub ParameterVelden_KeyPress(sender As Object, e As KeyPressEventArgs)
      Try
         Dim CursorPositie As New Integer
         Dim MaskerTeken As New Char
         Dim ParameterIndex As Integer = Nothing
         Dim Teken As Char = Nothing
         Dim Tekst As String = Nothing

         If Not (e.KeyChar.ToString().ToUpper() = Keys.C.ToString() AndAlso My.Computer.Keyboard.CtrlKeyDown) Then
            With DirectCast(sender, TextBox)
               Integer.TryParse(.Name.Substring(.Name.Length - 6), ParameterIndex)

               Select Case e.KeyChar
                  Case ToChar(Keys.Back)
                     If .SelectionStart > 0 Then
                        CursorPositie = .SelectionStart
                        Teken = If(.SelectionStart < QueryParameters()(ParameterIndex).Masker.Length, QueryParameters()(ParameterIndex).Masker.Chars(CursorPositie), Nothing)
                     End If
                  Case Else
                     Teken = ToChar(e.KeyChar.ToString().ToUpper())
                     CursorPositie = .SelectionStart + 1
                     MaskerTeken = If(.SelectionStart < QueryParameters()(ParameterIndex).Masker.Length, QueryParameters()(ParameterIndex).Masker.Chars(.SelectionStart), Nothing)
                     If Not ParameterMaskerTekenGeldig(Teken, MaskerTeken) = Nothing Then Teken = Nothing
               End Select

               If CursorPositie > 0 AndAlso CursorPositie <= QueryParameters()(ParameterIndex).Masker.Length() Then
                  Tekst = $"{ .Text}{QueryParameters()(ParameterIndex).Masker.Substring(.Text.Length)}"
                  Tekst = Tekst.Remove(CursorPositie - 1, 1).Insert(CursorPositie - 1, Teken)
                  .Text = Tekst
                  .SelectionStart = If(e.KeyChar = ToChar(Keys.Back), CursorPositie - 1, CursorPositie)
               End If

               e.Handled = True
            End With
         End If
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure maakt het parameterveld leeg wanneer de gebruiker de delete knop in drukt.
   Private Sub ParameterVelden_KeyUp(sender As Object, e As KeyEventArgs)
      Try
         Dim Masker As String = Nothing
         Dim ParameterIndex As New Integer
         Dim StandaardWaarde As String = Nothing

         With DirectCast(sender, TextBox)
            Integer.TryParse(.Name.Substring(.Name.Length - 6), ParameterIndex)

            If e.KeyCode = Keys.Delete Then
               Masker = QueryParameters()(ParameterIndex).Masker
               StandaardWaarde = QueryParameters()(ParameterIndex).StandaardWaarde
               .Text = $"{StandaardWaarde}{Masker.Substring(StandaardWaarde.Length)}"
            End If
         End With
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure selecteert de inhoud van het geactiveerde parameterveld.
   Private Sub ParameterVelden_MouseDown(sender As Object, e As MouseEventArgs)
      Try
         With DirectCast(sender, TextBox)
            .SelectionStart = 0
            .SelectionLength = .Text.Length
         End With
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure past dit venster aan de geselecteerde query aan.
   Private Sub PasVensterAan()
      Try
         Dim HuidigeVeldIndex As Integer = 0
         Dim ParameterIndex As Integer = 0
         Dim ParameterLabel As Label = Nothing
         Dim ParameterVeld As TextBox = Nothing
         Dim ParameterVelden As New List(Of Object)(VraagParameterVeldHouderObjectenOp("ParameterVelden"))
         Dim VeldIndexTekst As String = $"{HuidigeVeldIndex:D5}"
         Dim VorigeParameterLabel As Label = DirectCast(Me.ParameterVeldHouder.Controls($"ParameterLabels{VeldIndexTekst}"), Label)
         Dim VorigeParameterVeld As TextBox = DirectCast(Me.ParameterVeldHouder.Controls($"ParameterVelden{VeldIndexTekst}"), TextBox)
         Dim ZichtbareVelden As Integer = 0

         Do While ParameterIndex <= QueryParameters().Count - 1
            With QueryParameters()(ParameterIndex)
               VeldIndexTekst = $"{HuidigeVeldIndex:D5}"

               If HuidigeVeldIndex >= ParameterVelden.Count Then
                  ParameterLabel = New Label With {.Anchor = AnchorStyles.None,
                                                   .AutoSize = False,
                                                   .Dock = DockStyle.None,
                                                   .Enabled = True,
                                                   .Height = 20,
                                                   .Left = VorigeParameterLabel.Left,
                                                   .Name = $"ParameterLabels{VeldIndexTekst}",
                                                   .TextAlign = ContentAlignment.MiddleRight,
                  .Width = VorigeParameterLabel.Width}
                  ParameterVeld = New TextBox With {.Anchor = AnchorStyles.None,
                                                    .Dock = DockStyle.None,
                                                    .Left = VorigeParameterVeld.Left,
                                                    .Name = $"ParameterVelden{VeldIndexTekst}",
                                                    .TabIndex = (QueryOpenenKnop.TabIndex + 1) + HuidigeVeldIndex,
                                                    .Width = VorigeParameterVeld.Width}
               Else
                  ParameterLabel = DirectCast(Me.ParameterVeldHouder.Controls($"ParameterLabels{VeldIndexTekst}"), Label)
                  ParameterVeld = DirectCast(Me.ParameterVeldHouder.Controls($"ParameterVelden{VeldIndexTekst}"), TextBox)
               End If

               ParameterLabel.Text = $"{ .ParameterNaam}:"
               ParameterLabel.Visible = .VeldIsZichtbaar
               Tip.SetToolTip(ParameterLabel, .ParameterNaam)

               ParameterVeld.Enabled = True
               ParameterVeld.Text = $"{ .StandaardWaarde}{ .Masker.Substring(.StandaardWaarde.Length)}"
               ParameterVeld.ReadOnly = (.Masker = Nothing)
               ParameterVeld.MaxLength = .Masker.Length
               If Not .Commentaar.Trim() = Nothing Then Tip.SetToolTip(ParameterVeld, .Commentaar)
               ParameterVeld.Visible = .VeldIsZichtbaar
               If ParameterVeld.Visible Then ZichtbareVelden += 1
            End With

            AddHandler ParameterVeld.GotFocus, AddressOf ParameterVelden_GotFocus
            AddHandler ParameterVeld.KeyDown, AddressOf ParameterVelden_KeyDown
            AddHandler ParameterVeld.KeyPress, AddressOf ParameterVelden_KeyPress
            AddHandler ParameterVeld.KeyUp, AddressOf ParameterVelden_KeyUp
            AddHandler ParameterVeld.MouseDown, AddressOf ParameterVelden_MouseDown

            Me.ParameterVeldHouder.Controls.AddRange({ParameterLabel, ParameterVeld})

            ParameterLabel.Top = ((ZichtbareVelden - 1) * 24)
            ParameterVeld.Top = ParameterLabel.Top

            HuidigeVeldIndex += 1
            ParameterIndex += 1
            VorigeParameterLabel = ParameterLabel
            VorigeParameterVeld = ParameterVeld
         Loop

         ParametersFrame.Enabled = (ZichtbareVelden > 0)
         ParameterVeldHouder.AutoScrollPosition = New Point(0, 0)

         ParameterFrameSchuifBalk.Enabled = True
         ParameterFrameSchuifBalk.Minimum = 0
         ParameterFrameSchuifBalk.Maximum = CInt(ZichtbareVelden * 1.25)
         ParameterFrameSchuifBalk.Value = 0

         For VeldIndex As Integer = 0 To ParameterVelden.Count - 1
            ParameterVeld = DirectCast(Me.ParameterVeldHouder.Controls($"ParameterVelden{VeldIndex:D5}"), TextBox)

            If ParameterVeld.Visible Then
               ParameterVeld.Focus()
               Exit For
            End If
         Next VeldIndex

         QueryUitvoerenKnop.Enabled = ((Not (Query().Pad = Nothing)) AndAlso VERBINDING_GEOPEND(Verbinding()))
         If QueryUitvoerenKnop.Enabled AndAlso ZichtbareVelden = 0 Then QueryUitvoerenKnop.Focus()

         If Not Query().Pad = Nothing Then
            Me.Text = $"{My.Application.Info.Title} {PROGRAMMA_VERSIE} - ""{Query().Pad}""{NewLine}"
            ToonStatus($"Query: ""{Query().Pad}""{NewLine}")
         End If

         QueryPadVeld.Text = VerwijderAanhalingsTekens(Query().Pad)
         QueryPadVeld.SelectionStart = 0
         If Not QueryPadVeld.Text = Nothing Then QueryPadVeld.SelectionStart = QueryPadVeld.Text.Length - 1
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure geeft de opdracht om de query op het opgegeven pad te openen.
   Private Sub QueryOpenenKnop_Click(sender As Object, e As EventArgs) Handles QueryOpenenKnop.Click
      Try
         If Not QueryPadVeld.Text = Nothing Then ToonQuery(QueryPadVeld.Text)
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure opent de eerste van een of meer bestanden die in het querypadveld gesleept worden.
   Private Sub QueryPadVeld_DragDrop(sender As Object, e As DragEventArgs) Handles QueryPadVeld.DragDrop
      Try
         If e.Data.GetDataPresent(DataFormats.FileDrop) Then ToonQuery(DirectCast(e.Data.GetData(DataFormats.FileDrop), String()).First())
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure controleert of de eventuele objecten die in dit venster gesleept worden bestanden zijn.
   Private Sub QueryPadVeld_DragEnter(sender As Object, e As DragEventArgs) Handles QueryPadVeld.DragEnter
      Try
         If e.Data.GetDataPresent(DataFormats.FileDrop) Then e.Effect = DragDropEffects.All
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure verwerkt de toetsaanslagen van de gebruiker in het queryresultaat veld.
   Private Sub QueryResultaatVeld_KeyUp(sender As Object, e As KeyEventArgs) Handles QueryResultaatVeld.KeyUp
      Try
         Static ResultaatIndex As Integer = 0

         If My.Computer.Keyboard.CtrlKeyDown Then
            If QueryResultaten().Count > 0 Then
               Select Case e.KeyCode
                  Case Keys.PageUp
                     If ResultaatIndex > 0 Then ResultaatIndex -= 1
                  Case Keys.PageDown
                     If ResultaatIndex < QueryResultaten().Count - 1 Then ResultaatIndex += 1
               End Select

               ToonQueryResultaat(QueryResultaatVeld, ResultaatIndex)
            End If
         End If
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure geeft de opdracht om de gebruiker te verzoeken een query te selecteren.
   Private Sub QuerySelecterenKnop_Click(sender As Object, e As EventArgs) Handles QuerySelecterenKnop.Click
      Try
         Dim QueryPad As String = VraagQueryPad()

         If Not QueryPad = Nothing Then ToonQuery(QueryPad)
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure geeft de opdracht om de geselecteerde query met de opgegeven parameters uit te voeren.
   Private Sub QueryUitvoerenKnop_Click(sender As Object, e As EventArgs) Handles QueryUitvoerenKnop.Click
      Try
         If Instellingen().BatchInteractief Then
            If ParametersGeldig(VraagParameterVeldHouderObjectenOp("ParameterVelden")) Then
               InteractieveBatchAfbreken = False
               Me.Enabled = False
               QueryUitvoerenKnop.Enabled = False
            End If
         Else
            GeefQueryOpdracht()
         End If
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure reset dit venster.
   Private Sub ResetVenster()
      Try
         Dim ParameterLabel As Label = Nothing
         Dim ParameterVeld As TextBox = Nothing
         Dim ParameterVelden As New List(Of Object)(VraagParameterVeldHouderObjectenOp("ParameterVelden"))
         Dim VeldIndexTekst As String = Nothing

         AutomatischResultaatExporterenVeld.Enabled = True
         ExporteerResultaatNaarLabel.Enabled = True
         ExportFrame.Enabled = True
         ExportPadSelecterenKnop.Enabled = True
         MaakEMailMetExportBijgevoegdVeld.Enabled = True
         OpenResultaatNaExportVeld.Enabled = True
         ParametersFrame.Enabled = False
         ParameterFrameSchuifBalk.Enabled = False
         QueryLabel.Enabled = True
         QueryOpenenKnop.Enabled = True
         QueryPadVeld.Enabled = True
         QueryResultaatVeld.Enabled = True
         QuerySelecterenKnop.Enabled = True
         QueryUitvoerenKnop.Enabled = False
         ResultaatExporterenKnop.Enabled = True
         ResultaatFrame.Enabled = True

         QueryResultaatVeld.Text = Nothing

         Tip.SetToolTip(AutomatischResultaatExporterenVeld, "Als dit veld is aangevinkt, dan wordt het queryresultaat naar het opgegeven pad geëxporteerd, ")
         Tip.SetToolTip(ExportPadSelecterenKnop, "Klik hier om een venster te openen om naar een map voor het export bestand te bladeren, ")
         Tip.SetToolTip(ExportPadVeld, "Hier kan het pad waar het queryresultaat naar wordt geëxporteerd opgegeven worden, ")
         Tip.SetToolTip(MaakEMailMetExportBijgevoegdVeld, "Als dit veld is aangevinkt, dan wordt een e-mail met het geëxporteerde queryresultaat gemaakt, ")
         Tip.SetToolTip(OpenResultaatNaExportVeld, "Als dit veld is aangevinkt, dan wordt het geëxporteerde resultaat geopend, ")
         Tip.SetToolTip(QueryOpenenKnop, "Klik hier om het opgegeven query bestand te openen, ")
         Tip.SetToolTip(QueryPadVeld, "Hier kan het pad van een querybestand worden opgegeven, ")
         Tip.SetToolTip(QueryResultaatVeld, "Hier wordt het queryresultaat weergegeven. Druk op de toetsen Control + Page Up of Page Down om te bladeren tussen meerdere queryresultaten, ")
         Tip.SetToolTip(QuerySelecterenKnop, "Klik hier om een venster te openen om naar een query bestand te bladeren, ")
         Tip.SetToolTip(QueryUitvoerenKnop, "Klik hier om de query met de opgegeven parameters uit te voeren, ")
         Tip.SetToolTip(ResultaatExporterenKnop, "Klik hier om het queryresultaat te exporteren naar het opgegeven pad, ")
         Tip.SetToolTip(StatusVeld, "Hier wordt de status informatie weergegeven. Klik met de rechtermuisknop in de tekst voor opties, ")

         For VeldIndex As Integer = 0 To ParameterVelden.Count - 1
            VeldIndexTekst = $"{VeldIndex:D5}"
            ParameterLabel = DirectCast(Me.ParameterVeldHouder.Controls($"ParameterLabels{VeldIndexTekst}"), Label)
            ParameterLabel.Text = "Parameter:"
            ParameterLabel.Enabled = False
            Tip.SetToolTip(ParameterLabel, ParameterLabel.Text)

            ParameterVeld = DirectCast(Me.ParameterVeldHouder.Controls($"ParameterVelden{VeldIndexTekst}"), TextBox)

            RemoveHandler Me.ParameterVeldHouder.Controls(ParameterVeld.Name).GotFocus, AddressOf ParameterVelden_GotFocus
            RemoveHandler Me.ParameterVeldHouder.Controls(ParameterVeld.Name).KeyDown, AddressOf ParameterVelden_KeyDown
            RemoveHandler Me.ParameterVeldHouder.Controls(ParameterVeld.Name).KeyPress, AddressOf ParameterVelden_KeyPress
            RemoveHandler Me.ParameterVeldHouder.Controls(ParameterVeld.Name).KeyUp, AddressOf ParameterVelden_KeyUp
            RemoveHandler Me.ParameterVeldHouder.Controls(ParameterVeld.Name).MouseDown, AddressOf ParameterVelden_MouseDown

            ParameterVeld.Enabled = False
            ParameterVeld.Text = Nothing
            Tip.SetToolTip(ParameterVeld, "Voer hier een waarde in voor de parameter.")

            If VeldIndex > 0 Then
               Me.ParameterVeldHouder.Controls.Remove(ParameterLabel)
               Me.ParameterVeldHouder.Controls.Remove(ParameterVeld)
            End If
         Next VeldIndex

         Me.Text = $"{My.Application.Info.Title} {PROGRAMMA_VERSIE}"
      Catch
         HandelFoutAf()
      Finally
         If Instellingen().BatchInteractief Then ZetVensterInBatchModus()
      End Try
   End Sub

   'Deze procedure geeft de opdracht om het queryresultaat te exporteren.
   Private Sub ResultaatExporterenKnop_Click(sender As Object, e As EventArgs) Handles ResultaatExporterenKnop.Click
      Try
         GeefExportOpdracht()
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure sluit dit venster.
   Private Sub SluitenMenu_Click(sender As Object, e As EventArgs) Handles SluitenMenu.Click
      Try
         Me.Close()
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure geeft de opdracht het opgegeven/eerder geladen querybestand te tonen.
   Private Sub ToonQuery(Optional Pad As String = Nothing)
      Try
         If Not BATCH_MODUS_ACTIEF() Then QueryParameters(If(Pad = Nothing, Query().Code, Query(VerwijderAanhalingsTekens(Pad)).Code))

         QueryResultaten(, ResultatenVerwijderen:=True)
         ResetVenster()
         PasVensterAan()

         If Instellingen().QueryAutoUitvoeren Then GeefQueryOpdracht()
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure verschuift de schuifbalk zodat het opgegeven parameterveld zichtbaar wordt.
   Private Sub VerschuifBalk(VeldIndex As Integer)
      Dim Rij As Integer = 0

      Try
         For ParameterIndex As Integer = 0 To VeldIndex
            If Me.ParameterVeldHouder.Controls($"ParameterVelden{ParameterIndex:D5}").Visible Then Rij += 1
         Next ParameterIndex
      Catch
         HandelFoutAf()
      Finally
         ParameterFrameSchuifBalk.Value = Rij - 1
      End Try
   End Sub

   'Deze procedure stuurt een lijst van parameter veld houder objecten waarvan de naam zoals opgegeven begint terug.
   Private Function VraagParameterVeldHouderObjectenOp(Naam As String) As List(Of Object)
      Try
         Return New List(Of Object)(From VensterObject In Me.ParameterVeldHouder.Controls Where DirectCast(VensterObject, Control).Name.StartsWith(Naam))
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure zet dit venster in interactieve batchmodus.
   Private Sub ZetVensterInBatchModus()
      Try
         AutomatischResultaatExporterenVeld.Enabled = False
         ExporteerResultaatNaarLabel.Enabled = False
         ExportFrame.Enabled = False
         ExportPadSelecterenKnop.Enabled = False
         MaakEMailMetExportBijgevoegdVeld.Enabled = False
         OpenResultaatNaExportVeld.Enabled = False
         QueryLabel.Enabled = False
         QueryOpenenKnop.Enabled = False
         QueryPadVeld.Enabled = False
         QueryResultaatVeld.Enabled = False
         QuerySelecterenKnop.Enabled = False
         ResultaatExporterenKnop.Enabled = False
         ResultaatFrame.Enabled = False

         ExportPadVeld.Text = Nothing
         QueryUitvoerenKnop.Text = "Batch &Uitvoeren"
         Tip.SetToolTip(QueryUitvoerenKnop, "Klik hier om de batch met de opgegeven parameters uit te voeren.")
      Catch
         HandelFoutAf()
      End Try
   End Sub
End Class