'De instellingen en geimporteerde namespaces van deze module.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Windows.Forms

'Deze module bevat het invoer venster.
Public Class InvoerVenster
   'Deze procedure sluit dit venster en annuleert de invoer.
   Private Sub AnnulerenKnop_Click(sender As Object, e As EventArgs) Handles AnnulerenKnop.Click
      Try
         Me.DialogResult = DialogResult.Cancel
         Me.Close()
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure stelt dit venster in.
   Private Sub InvoerVenster_Shown(sender As Object, e As EventArgs) Handles Me.Shown
      Try
         Me.Text = My.Application.Info.Title
         PromptLabel.MaximumSize = Paneel.Size
         TekstVeld.Focus()
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure sluit dit venster en bevestigt de invoer.
   Private Sub OKKnop_Click(sender As Object, e As EventArgs) Handles OKKnop.Click
      Try
         Me.DialogResult = DialogResult.OK
         Me.Close()
      Catch
         HandelFoutAf()
      End Try
   End Sub
End Class
