'De instellingen en geimporteerde namespaces van deze module.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.Office.Interop
Imports System.Collections.Generic

'Deze module bevat de Microsoft Outlook gerelateerde procedures.
Public Class EMailClass
   Private WithEvents MSOutlook As New Outlook.Application                                                                            'Bevat een verwijzing naar Microsoft Outlook.
   Private WithEvents EMail As Outlook.MailItem = DirectCast(MSOutlook.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)   'Bevat een verwijzing naar een Microsoft Outlook e-mail bericht.

   Private OutlookReedsActief As Boolean = False   'Geeft aan of Microsoft Outlook reeds actief is.

   'Deze procedure stelt deze module in.
   Public Sub New()
      Try
         If MSOutlook IsNot Nothing Then EMail.GetInspector.Activate()
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure wordt uitgevoerd wanneer een nieuwe e-mail wordt geopend.
   Private Sub EMail_Open() Handles EMail.Open
      Try
         With Instellingen()
            If EMail IsNot Nothing Then
               EMail.Body = VervangSymbolen(.EMailTekst.ToString())
               EMail.CC = .ExportCCOntvanger
               EMail.SentOnBehalfOfName = .ExportAfzender
               EMail.Subject = VervangSymbolen(.ExportOnderwerp)
               EMail.To = .ExportOntvanger
            End If
         End With
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure wordt uitgevoerd wanneer een e-mail wordt afgesloten.
   Private Sub EMail_Unload() Handles EMail.Unload
      Try
         If (Not (Instellingen().QueryAutoSluiten OrElse OutlookReedsActief)) AndAlso MSOutlook IsNot Nothing Then
            MSOutlook.GetNamespace("MAPI").Logoff()
            MSOutlook.Quit()
         End If
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure wordt uitgevoerd wanneer Microsoft Outlook wordt gestart.
   Private Sub MSOutlook_Startup() Handles MSOutlook.Startup
      Try
         OutlookReedsActief = True
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure voegt de opgegeven geëxporteerde query resultaten toe aan een e-mail.
   Public Sub VoegQueryResultatenToe(ExportPaden As List(Of String))
      Try
         If EMail IsNot Nothing AndAlso MSOutlook IsNot Nothing Then
            ExportPaden.ForEach(AddressOf EMail.Attachments.Add)
            If Instellingen().ExportAutoVerzenden Then EMail.Send()
         End If
      Catch
         HandelFoutAf()
      End Try
   End Sub
End Class
