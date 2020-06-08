'De instellingen en geimporteerde namespaces van deze module.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Windows.Forms

'Deze module bevat het inlogvenster.
Public Class InloggenVenster

   Private WithEvents Tip As New ToolTip   'Geeft tips met betrekking tot de knoppen en velden in dit venster weer.

   'Deze procedure sluit dit venster.
   Private Sub AnnulerenKnop_Click(sender As Object, e As EventArgs) Handles AnnulerenKnop.Click
      Try
         Me.Close()
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure geeft de opdracht om verbinding met een database te maken.
   Private Sub InloggenKnop_Click(sender As Object, e As EventArgs) Handles InloggenKnop.Click
      Try
         Verbinding(VerwerkInlogGegevens(GebruikerVeld.Text, WachtwoordVeld.Text, Instellingen().VerbindingsInformatie.ToString()))
      Catch
         HandelFoutAf()
      Finally
         If VERBINDING_GEOPEND(Verbinding()) Then
            Me.Close()
         Else
            If GebruikerVeld.Enabled Then
               GebruikerVeld.Focus()
            ElseIf WachtwoordVeld.Enabled Then
               WachtwoordVeld.Focus()
            End If
         End If
      End Try
   End Sub

   'Deze procedure stelt dit venster in.
   Private Sub InloggenVenster_Load(sender As Object, e As EventArgs) Handles MyBase.Load
      Try
         Me.Left = CInt(My.Computer.Screen.WorkingArea.Width / 2) - CInt(Me.Width / 2)
         Me.Top = CInt(My.Computer.Screen.WorkingArea.Height / 3) - CInt(Me.Height / 2)

         GebruikerLabel.Enabled = Instellingen().VerbindingsInformatie.ToString().ToUpper().Contains(GEBRUIKER_VARIABEL)
         WachtwoordLabel.Enabled = Instellingen().VerbindingsInformatie.ToString().ToUpper().Contains(WACHTWOORD_VARIABEL)
         GebruikerVeld.Enabled = GebruikerLabel.Enabled
         WachtwoordVeld.Enabled = WachtwoordLabel.Enabled

         Tip.SetToolTip(AnnulerenKnop, "Klik hier om het inloggen af te breken en het programma te beëindigen.")
         Tip.SetToolTip(GebruikerVeld, "Voer hier een gebruikersnaam in, als deze vereist is voor de verbinding met de database.")
         Tip.SetToolTip(InloggenKnop, Instellingen().Bestand)
         Tip.SetToolTip(WachtwoordVeld, "Voer hier een wachtwoord in, als deze vereist is voor de verbinding met de database.")
      Catch
         HandelFoutAf()
      End Try
   End Sub
End Class