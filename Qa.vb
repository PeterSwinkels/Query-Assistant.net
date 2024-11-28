'De instellingen en geimporteerde namespaces van deze module.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports ADODB
Imports Microsoft.Office.Interop
Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Diagnostics
Imports System.Environment
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

'Deze module bevat de kern procedures van dit programma.
Public Module QaModule
   'Bevat een opsomming van de parameter definitie elementen.
   Private Enum ParameterDefinitieOpsomming As Integer
      NaamElement                                  'Naamelement.
      MaskerElement                                'Maskerelement.
      StandaardWaardeElement                       'Standaardwaardeelement.
      CommentaarElement                            'Commentaarelement.
   End Enum

   'Bevat de definities voor de instellingen van dit programma.
   Public Structure InstellingenDefinitie
      Public BatchBereik As String                  'Definieert de volgnummers van de eerste en de laatste query in een uit te voeren batch.
      Public BatchInteractief As Boolean            'Definieert of de gebruiker eerst parameters moet invoeren voordat een batch uitgevoerd kan worden.
      Public BatchQueryPad As String                'Definieert het pad en/of de bestandsnaam zonder volgnummers van de query's in een uit te voeren batch.
      Public Bestand As String                      'Definieert het pad en/of de bestandsnaam van het programmainstellingenbestand.
      Public EMailTekst As StringBuilder            'Definieert de tekst van de e-mail met de geëxporteerde resultaten.
      Public ExportAfzender As String               'Definieert de naam van de afzender van de e-mail met de geëxporteerde resultaten.
      Public ExportAutoOpenen As Boolean            'Definieert of een export automatisch na het exporteren geopend wordt.
      Public ExportAutoOverschrijven As Boolean     'Definieert of een bestand automatisch overschreven wordt bij het exporteren van de queryresultaataten.
      Public ExportAutoVerzenden As Boolean         'Definieert of de e-mail met de geëxporteerde resultaten automatisch verzonden wordt.
      Public ExportCCOntvanger As String            'Definieert het e-mail adres van de ontvanger van het kopie van de e-mail met de geëxporteerde resultaten.
      Public ExportKolomAanvullen As Boolean        'Definieert of de data in een kolom moet worden aangevuld met spaties.
      Public ExportOnderwerp As String              'Definieert het onderwerp van de e-mail met de geëxporteerde resultaten.
      Public ExportOntvanger As String              'Definieert het e-mail adres van de ontvanger van de e-mail met de geëxporteerde resultaten.
      Public ExportStandaardPad As String           'Definieert het standaardpad voor het exporteren van queryresultaataten.
      Public QueryAutoSluiten As Boolean            'Definieert of dit programma na het uitvoeren van een query en een eventuele export automatisch afgesloten wordt.
      Public QueryAutoUitvoeren As Boolean          'Definieert of een query automatisch uitgevoerd wordt na het laden.
      Public QueryRecordSets As Boolean             'Definieert of er meer dan een recordset kan worden teruggestuurd door de database als het resultaat van een query.
      Public QueryTimeout As Integer                'Definieert het aantal seconden dat het programma wacht op het queryresultaat nadat opdracht is gegeven de query uit te voeren.
      Public VoorbeeldKolomBreedte As Integer       'Definieert de maximale kolombreedte die gebruikt wordt om het queryresultaat te tonen in het voorbeeldvenster.
      Public VoorbeeldRegels As Integer             'Definieert het maximum aantal regels dat van het queryresultaat wordt getoond in het voorbeeldvenster.
      Public VerbindingsInformatie As StringBuilder 'Definieert de voor de verbinding met een database noodzakelijke gegevens.
   End Structure

   'Bevat de definities voor de opdrachtregelparameters die eventueel zijn opgegeven bij het starten van dit programma.
   Public Structure OpdrachtRegelParametersDefinitie
      Public InstellingenPad As String   'Definieert het opgegeven instellingenpad.
      Public QueryPad As String          'Definieert het opgegeven querypad.
      Public SessiesPad As String        'Definieert het opgegeven sessielijstpad.
      Public Verwerkt As Boolean         'Definieert of de opdrachtregelparameters zonder fouten zijn verwerkt.
   End Structure

   'Bevat de definities voor een query.
   Public Structure QueryDefinitie
      Public Code As String              'Definieert de code van een query.
      Public Pad As String               'Definieert het pad van een querybestand.
      Public Geopend As Boolean          'Definieert of het querybestand kon worden geopend.
   End Structure

   'Bevat de definities voor de parameter gegevens van de geselecteerde query.
   Public Structure QueryParameterDefinitie
      Public Commentaar As String        'Definieert het commentaar bij de parameter.
      Public Invoer As String            'Definieert de invoer van de gebruiker.
      Public Lengte As Integer           'Definieert de lengte van de parameterdefinitie.
      Public LengteIsVariabel As Boolean 'Definieert of de lengte van de invoer variabel is.
      Public Masker As String            'Definieert het invoer masker van de parameter.
      Public ParameterNaam As String     'Definieert de naam van de parameter.
      Public Positie As Integer          'Definieert de positie relatief ten op zichte van de vorige definitie.
      Public StandaardWaarde As String   'Definieert de standaardwaarde van de parameter.
      Public VeldIsZichtbaar As Boolean  'Definieert of het invoerveld zichtbaar is.
   End Structure

   'Bevat de definities voor het resultaat van een query.
   Public Structure QueryResultaatDefinitie
      Public KolomBreedte As List(Of Integer)      'Definieert per kolom de maximale breedte in bytes van de gegevens.
      Public RechtsUitlijnen As List(Of Boolean)   'Definieert per kolom of de gegevens rechtsuitgelijnd worden bij weergave.
      Public Tabel(,) As String                    'Definieert de door een query opgevraagde gegevens uit een database.
   End Structure

   Public Const GEBRUIKER_VARIABEL As String = "$$GEBRUIKER$$"       'Indien aanwezig in de verbindingsinformatie geeft deze variabel de positie van de gebruikersnaam aan.
   Public Const WACHTWOORD_VARIABEL As String = "$$WACHTWOORD$$"     'Indien aanwezig in de verbindingsinformatie geeft deze variabel de positie van het wachtwoord aan.
   Private Const ALLE_REGELS As Integer = -1                          'Staat voor "alle voorbeeld regels".
   Private Const ASCII_A As Integer = 65                              'De ASCII-waarde voor het teken "A".
   Private Const ASCII_Z As Integer = 90                              'De ASCII-waarde voor het teken "Z".
   Private Const COMMENTAAR_TEKEN As Char = "#"c                      'Geeft aan dat een regel in een instellingenbestand commentaar is.
   Private Const CSV_SCHEIDINGSTEKEN As Char = ";"c                   'Scheidt kolommen in CSV bestanden.
   Private Const DEFINITIE_TEKENS As String = "$$"                    'Geeft het begin en het einde van een parameterdefinitie binnen een query aan.
   Private Const ELEMENT_TEKEN As Char = ":"c                         'Scheidt de parameterdefinitie elementen van elkaar.
   Private Const EXCEL_MAXIMUM_AANTAL_KOLOMMEN As Integer = 255       'Het maximale aantal door Microsoft Excel ondersteunde kolommen.
   Private Const GEEN_LETTER As Integer = 64                          'Staat voor "geen letter". (De ASCII waarde die voor het "A" teken komt.)
   Private Const GEEN_MAXIMALE_BREEDTE As Integer = -1                'Staat voor "geen maximale kolom breedte".
   Private Const MAXIMUM_AANTAL_ELMENTEN As Integer = 4               'Staat voor het maximaal toegestane aantal elementen in een parameterdefinitie.
   Private Const MASKER_CIJFER As Char = "#"c                         'Geeft in een masker aan dat er een cijfer als invoer wordt verwacht.
   Private Const MASKER_HOOFDLETTER As Char = "_"c                    'Geeft in een masker aan dat er een hoofdletter als invoer wordt verwacht.
   Private Const PARAMETER_TEKEN As Char = "?"c                       'Scheidt de opdrachtregelparameters van elkaar.
   Private Const SECTIE_NAAM_BEGIN As String = "["c                   'Geeft het begin van een sectie naam in een instellingenbestand aan.
   Private Const SECTIE_NAAM_EINDE As String = "]"c                   'Geeft het einde van een sectie naam in een instellingenbestand aan.
   Private Const SQL_COMMENTAAR_BLOK_BEGIN As String = "/*"           'Staat voor het begin van een SQL-commentaarblok.
   Private Const SQL_COMMENTAAR_BLOK_EINDE As String = "*/"           'Staat voor het einde van een SQL-commentaarblok.
   Private Const SQL_COMMENTAAR_REGEL_BEGIN As String = "--"          'Staat voor het begin van een SQL-commentaarregel.
   Private Const SQL_COMMENTAAR_REGEL_EINDE As String = Nothing       'Staat voor het einde van een SQL-commentaarregel.
   Private Const SYMBOOL_TEKEN As Char = "*"c                         'Geeft het begin en het einde van een symbool in een tekst aan.
   Private Const TEKENREEKS_TEKENS As String = "'"""                  'Staat voor de tekens die het begin en einde van een tekenreeks aanduiden.
   Private Const VARIABELE_LENGTE_TEKEN As Char = "*"c                'Indien aanwezig aan het begin van een masker geeft dit teken aan dat de invoer lengte variabel is.
   Private Const VERBINDING_AFSCHEIDING_BEGIN As Char = "("c          'Scheidt de verbindingsinformatieparameters van elkaar.
   Private Const VERBINDING_AFSCHEIDING_EINDE As Char = ")"c          'Scheidt de verbindingsinformatieparameters van elkaar.
   Private Const VERBINDING_PARAMETER_TEKEN As Char = ";"c            'Scheidt de verbindingsinformatieparameters van elkaar.
   Private Const WAARDE_TEKEN As Char = "="c                          'Scheidt de naam en waarde van een instellingenparameter van elkaar.

   Public ReadOnly BATCH_MODUS_ACTIEF As Func(Of Boolean) = Function() Not (Instellingen().BatchBereik = Nothing OrElse Instellingen().BatchQueryPad = Nothing)                                                                                                    'Geeft aan of de  batchmodus actief is.
   Public ReadOnly BITS_MODUS As String = $"{If(Is64BitProcess, "64", "32")}-bits"                                                                                                                                                                                 'Geeft de bits-modus voor dit programma op. (32-bits of 64-bits).
   Public ReadOnly OPDRACHT_REGEL As String = String.Join(" "c, GetCommandLineArgs.Skip(1))                                                                                                                                                                        'Bevat de eventuele opdrachtregelparameters.
   Public ReadOnly PROGRAMMA_VERSIE As String = $"v{My.Application.Info.Version} ({BITS_MODUS})"                                                                                                                                                                   'Bevat programma versie informatie.
   Public ReadOnly TOON_VERBINDINGS_STATUS As Action = Sub() ToonStatus(String.Format(If(VERBINDING_GEOPEND(Verbinding), "Verbonden met de database. - Instellingen: {0}{1}", "Er is geen verbinding met een database.{0}{1}"), Instellingen().Bestand, NewLine))  'Deze procedure toont de verbindingsstatus.
   Public ReadOnly VERBINDING_GEOPEND As Func(Of Connection, Boolean) = Function(VerbindingO As Connection) If(VerbindingO IsNot Nothing, VerbindingO.State = ObjectStateEnum.adStateOpen, False)                                                                  'Deze procedure geeft aan of de opgegeven verbinding geopend is.
   Private ReadOnly INTERACTIEVE_BATCH_PARAMETERS As New List(Of String)                                                                                                                                                                                           'Bevat de interactieve batch parameters.
   Private ReadOnly IS_INSTELLINGS_SECTIE As Func(Of String, Boolean) = Function(Regel As String) Regel.Trim().StartsWith(SECTIE_NAAM_BEGIN) AndAlso Regel.Trim().EndsWith(SECTIE_NAAM_EINDE)                                                                      'Deze procedure geeft aan of de opgegeven regel een instellingen sectie naam bevat.
   Private ReadOnly LINKS_UITGELIJNDE_DATA_TYPES As New List(Of DataTypeEnum)({DataTypeEnum.adBSTR, DataTypeEnum.adChar, DataTypeEnum.adDBDate, DataTypeEnum.adDBTime, DataTypeEnum.adDBTimeStamp, DataTypeEnum.adLongVarChar, DataTypeEnum.adLongVarWChar, DataTypeEnum.adVarChar, DataTypeEnum.adVarWChar, DataTypeEnum.adWChar})   'Bevat een lijst van databasedatatypes die linksuitgelijnd worden.
   Private ReadOnly RECORDSET_GEOPEND As Func(Of Recordset, Boolean) = Function(RecordsetO As Recordset) If(RecordsetO IsNot Nothing, RecordsetO.State = ObjectStateEnum.adStateOpen, False)                                                                       'Deze procedure geeft aan of de opgegeven recordset geopend is.
   Private ReadOnly SESSIE_PARAMETERS As New List(Of String)                                                                                                                                                                                                       'Bevat de sessieparameters.

   Public HuidigInterfaceVenster As Form = Nothing                                                                                              'Bevat de verwijzing naar een eventueel interface venster.
   Public InteractieveBatchAfbreken As Boolean = False                                                                                          'Geeft aan of een interactive batch moet worden afgebroken.
   Public InteractieveBatchModusActief As Boolean = InteractieveBatchModusActief = Instellingen.BatchInteractief AndAlso BATCH_MODUS_ACTIEF()   'Geeft aan of de interactieve batchmodus actief is.
   Public SessiesAfbreken As Boolean = False                                                                                                    'Geeft aan of een reeks sessies is afgebroken.

   'Deze procedure bewaart de instellingen van dit programma.
   Private Sub BewaarInstellingen(InstellingenPad As String, TeBewarenInstellingen As InstellingenDefinitie, Bericht As String)
      Try
         Dim BewaardeInstellingen As New List(Of String)

         With TeBewarenInstellingen
            BewaardeInstellingen.Add($"{SECTIE_NAAM_BEGIN}BATCH{SECTIE_NAAM_EINDE}")
            BewaardeInstellingen.Add($"Bereik{WAARDE_TEKEN}{ .BatchBereik}")
            BewaardeInstellingen.Add($"Interactief{WAARDE_TEKEN}{ .BatchInteractief}")
            BewaardeInstellingen.Add($"QueryPad{WAARDE_TEKEN}{ .BatchQueryPad}")
            BewaardeInstellingen.Add(NewLine)

            BewaardeInstellingen.Add($"{SECTIE_NAAM_BEGIN}EMAILTEKST{SECTIE_NAAM_EINDE}")
            BewaardeInstellingen.Add(.EMailTekst.ToString())
            BewaardeInstellingen.Add(NewLine)

            BewaardeInstellingen.Add($"{SECTIE_NAAM_BEGIN}EXPORT{SECTIE_NAAM_EINDE}")
            BewaardeInstellingen.Add($"Afzender{WAARDE_TEKEN}{ .ExportAfzender}")
            BewaardeInstellingen.Add($"AutoOpenen{WAARDE_TEKEN}{ .ExportAutoOpenen}")
            BewaardeInstellingen.Add($"AutoOverschrijven{WAARDE_TEKEN}{ .ExportAutoOverschrijven}")
            BewaardeInstellingen.Add($"AutoVerzenden{WAARDE_TEKEN}{ .ExportAutoVerzenden}")
            BewaardeInstellingen.Add($"CCOntvanger{WAARDE_TEKEN}{ .ExportCCOntvanger}")
            BewaardeInstellingen.Add($"KolomAanvullen{WAARDE_TEKEN}{ .ExportKolomAanvullen}")
            BewaardeInstellingen.Add($"Onderwerp{WAARDE_TEKEN}{ .ExportOnderwerp}")
            BewaardeInstellingen.Add($"Ontvanger{WAARDE_TEKEN}{ .ExportOntvanger}")
            BewaardeInstellingen.Add($"StandaardPad{WAARDE_TEKEN}{ .ExportStandaardPad}")
            BewaardeInstellingen.Add(NewLine)

            BewaardeInstellingen.Add($"{SECTIE_NAAM_BEGIN}QUERY{SECTIE_NAAM_EINDE}")
            BewaardeInstellingen.Add($"AutoSluiten{WAARDE_TEKEN}{ .QueryAutoSluiten}")
            BewaardeInstellingen.Add($"AutoUitvoeren{WAARDE_TEKEN}{ .QueryAutoUitvoeren}")
            BewaardeInstellingen.Add($"Recordsets{WAARDE_TEKEN}{ .QueryRecordSets}")
            BewaardeInstellingen.Add($"Timeout{WAARDE_TEKEN}{ .QueryTimeout}")
            BewaardeInstellingen.Add(NewLine)

            BewaardeInstellingen.Add($"{SECTIE_NAAM_BEGIN}VERBINDING{SECTIE_NAAM_EINDE}")
            BewaardeInstellingen.Add(.VerbindingsInformatie.ToString())
            BewaardeInstellingen.Add(NewLine)

            BewaardeInstellingen.Add($"{ SECTIE_NAAM_BEGIN}VOORBEELD{SECTIE_NAAM_EINDE}")
            BewaardeInstellingen.Add($"KolomBreedte{WAARDE_TEKEN}{ .VoorbeeldKolomBreedte}")
            BewaardeInstellingen.Add($"Regels{WAARDE_TEKEN}{ .VoorbeeldRegels}")
         End With

         Do
            Try
               File.WriteAllLines(InstellingenPad, BewaardeInstellingen)
               Exit Do
            Catch
               If HandelFoutAf() = DialogResult.Ignore Then Exit Do
            End Try
         Loop

         MessageBox.Show($"{Bericht}{NewLine}{InstellingenPad}", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Information)
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure stuurt de Microsoft Excel kolom id. voor het opgegeven kolom nummer terug.
   Private Function ExcelKolomId(Kolom As Integer) As String
      Try
         If Kolom >= 0 AndAlso Kolom <= EXCEL_MAXIMUM_AANTAL_KOLOMMEN Then
            For Letter1 As Integer = GEEN_LETTER To ASCII_Z
               For Letter2 As Integer = ASCII_A To ASCII_Z
                  If Kolom = 0 Then Return If(Letter1 = GEEN_LETTER, ToChar(Letter2), $"{Letter1}{Letter2}")
                  Kolom -= 1
               Next Letter2
            Next Letter1
         End If
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure exporteert het queryresultaat naar een tekstbestand.
   Private Function ExporteerAlsTekst(ExportPad As String, Optional CSV As Boolean = False) As Boolean
      Dim ExportAfgebroken As Boolean = False

      Try
         Dim GeexporteerdeTekst As New StringBuilder

         For ResultaatIndex As Integer = 0 To QueryResultaten().Count - 1
            With QueryResultaten()(ResultaatIndex)
               If .Tabel IsNot Nothing Then
                  If QueryResultaten().Count > 1 Then GeexporteerdeTekst.Append($"[Resultaat: #{ResultaatIndex + 1}]{NewLine}")
                  For Rij As Integer = 0 To .Tabel.GetUpperBound(0)
                     For Kolom As Integer = 0 To .Tabel.GetUpperBound(1)
                        If .Tabel(Rij, Kolom) = Nothing Then .Tabel(Rij, Kolom) = ""

                        If CSV Then
                           GeexporteerdeTekst.Append($"{ If(.Tabel(Rij, Kolom).Contains(CSV_SCHEIDINGSTEKEN), $"""{ .Tabel(Rij, Kolom).Replace("""", """""")}""", .Tabel(Rij, Kolom))}{CSV_SCHEIDINGSTEKEN}")
                        Else
                           If Instellingen().ExportKolomAanvullen Then
                              GeexporteerdeTekst.Append(If(.RechtsUitlijnen(Kolom), .Tabel(Rij, Kolom).PadLeft(.KolomBreedte(Kolom)), .Tabel(Rij, Kolom).PadRight(.KolomBreedte(Kolom))))
                           Else
                              GeexporteerdeTekst.Append($"{ .Tabel(Rij, Kolom)}{Microsoft.VisualBasic.ControlChars.Tab}")
                           End If
                        End If
                     Next Kolom
                     GeexporteerdeTekst.Append(NewLine)
                  Next Rij
               End If
            End With
         Next ResultaatIndex

         Do
            Try
               File.WriteAllText(ExportPad, GeexporteerdeTekst.ToString())
               Exit Do
            Catch
               If HandelFoutAf(TypePad:="Export pad: ", Pad:=ExportPad) = DialogResult.Ignore Then
                  ExportAfgebroken = True
                  Exit Do
               End If
            End Try
         Loop

         Return ExportAfgebroken
      Catch
         ExportAfgebroken = True
         HandelFoutAf()
      End Try

      Return ExportAfgebroken
   End Function

   'Deze procedure exporteert het queryresultaat naar een Microsoft Excel werkmap.
   Private Function ExporteerNaarExcel(ExportPad As String, ExcelFormaat As Excel.XlFileFormat) As Boolean
      Dim ExportAfgebroken As Boolean = False

      Try
         Dim KolomId As String = Nothing
         Dim MSExcel As New Excel.Application
         Dim WerkBlad As Excel.Worksheet = Nothing
         Dim WerkMap As Excel.Workbook = Nothing

         MSExcel.DisplayAlerts = False
         MSExcel.Interactive = False
         MSExcel.ScreenUpdating = False
         MSExcel.Workbooks.Add()

         WerkMap = MSExcel.Workbooks.Item(1)
         WerkMap.Activate()

         Do Until WerkMap.Worksheets.Count <= 1
            DirectCast(WerkMap.Worksheets.Item(WerkMap.Worksheets.Count), Excel.Worksheet).Delete()
         Loop

         Do Until WerkMap.Worksheets.Count >= QueryResultaten().Count
            WerkMap.Worksheets.Add()
         Loop

         For ResultaatIndex As Integer = 0 To QueryResultaten().Count - 1
            With QueryResultaten()(ResultaatIndex)
               If .Tabel IsNot Nothing Then
                  If .Tabel.GetUpperBound(1) > EXCEL_MAXIMUM_AANTAL_KOLOMMEN Then
                     MessageBox.Show($"Het queryresultaat bevat te veel kolommen om deze naar Microsoft Excel te exporteren.{NewLine}Het maximaal toegestane aantal kolommen is: {EXCEL_MAXIMUM_AANTAL_KOLOMMEN}", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                  Else
                     WerkBlad = DirectCast(WerkMap.Worksheets.Item(ResultaatIndex + 1), Excel.Worksheet)
                     WerkBlad.Activate()
                     If QueryResultaten().Count > 0 Then WerkBlad.Name = $"Resultaat {ResultaatIndex + 1}"

                     WerkBlad.Range($"A1:{ExcelKolomId(.Tabel.GetUpperBound(1))}{ .Tabel.GetUpperBound(0) + 1}").Value = .Tabel
                     For Kolom As Integer = 0 To .Tabel.GetUpperBound(1)
                        KolomId = ExcelKolomId(Kolom)
                        WerkBlad.Range($"{KolomId}1:{KolomId}1").Font.Bold = True
                        If .RechtsUitlijnen(Kolom) Then WerkBlad.Range($"{KolomId}1:{KolomId}{ .Tabel.GetUpperBound(0) + 1}").HorizontalAlignment = Excel.Constants.xlRight
                     Next Kolom
                     WerkBlad.Range($"A:{ExcelKolomId(.Tabel.GetUpperBound(1))}").Columns.AutoFit()
                  End If
               End If
            End With
         Next ResultaatIndex

         DirectCast(WerkMap.Worksheets.Item(1), Excel.Worksheet).Activate()
         Do
            Try
               WerkMap.SaveAs(ExportPad, ExcelFormaat)
               Exit Do
            Catch
               If HandelFoutAf(TypePad:="Export pad: ", Pad:=ExportPad) = DialogResult.Ignore Then
                  ExportAfgebroken = True
                  Exit Do
               End If
            End Try
         Loop
         WerkMap.Close()

         If Not MSExcel Is Nothing Then
            MSExcel.Quit()
            MSExcel.DisplayAlerts = True
            MSExcel.Interactive = True
            MSExcel.ScreenUpdating = True
         End If

         MSExcel = Nothing
         WerkBlad = Nothing
         WerkMap = Nothing

         Return ExportAfgebroken
      Catch
         ExportAfgebroken = True
         HandelFoutAf()
      End Try

      Return ExportAfgebroken
   End Function

   'Deze procedure exporteert het queryresultaat.
   Public Function ExporteerResultaat(ExportPad As String) As Boolean
      Dim ExportAfgebroken As Boolean = False

      Try
         Dim BestandsType As String = Path.GetExtension(ExportPad.Trim().ToLower())

         If File.Exists(ExportPad) Then
            If Not Instellingen().ExportAutoOverschrijven Then
               If MessageBox.Show($"Het bestand ""{ExportPad}"" bestaat al. Overschrijven?", My.Application.Info.Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then ExportAfgebroken = True
            End If

            If Not ExportAfgebroken Then
               If BestandsType = ".xls" OrElse BestandsType = ".xlsx" Then DirectCast(Microsoft.VisualBasic.GetObject(ExportPad), Excel.Workbook).Close(SaveChanges:=False)
               File.Delete(ExportPad)
            End If
         End If

         If Not ExportAfgebroken Then
            Select Case BestandsType
               Case ".csv"
                  ExportAfgebroken = ExporteerAlsTekst(ExportPad, CSV:=True)
               Case ".xls"
                  ExportAfgebroken = ExporteerNaarExcel(ExportPad, Excel.XlFileFormat.xlWorkbookNormal)
               Case ".xlsx"
                  ExportAfgebroken = ExporteerNaarExcel(ExportPad, Excel.XlFileFormat.xlWorkbookDefault)
               Case Else
                  ExportAfgebroken = ExporteerAlsTekst(ExportPad)
            End Select
         End If

         Return Not ExportAfgebroken
      Catch
         ExportAfgebroken = True
         HandelFoutAf()
      End Try

      Return Not ExportAfgebroken
   End Function

   'Deze procedure zet de opgegeven foutenlijst om naar tekst.
   Public Function FoutenLijstTekst(Lijst As Errors) As String
      Try
         Dim Tekst As New StringBuilder

         Tekst.Append("Er ")
         If Lijst.Count = 1 Then Tekst.Append("is 1 fout ") Else Tekst.Append($"zijn {Lijst.Count} fouten")
         Tekst.Append($" opgetreden tijdens het uitvoeren van de query:{NewLine}")
         Tekst.Append($"{"Native",-11}{"Code",-11}{"Bron",-36}{"SQL Status",-11}Omschrijving{NewLine}")

         For Each Fout As ADODB.Error In Lijst
            With Fout
               Tekst.Append($"{ .NativeError,10} { .Number,10} { .Source,-35} { .SQLState,10} { .Description}{NewLine}")
            End With
         Next Fout

         Return Tekst.ToString()
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure handelt eventuele fouten af.
   Public Function HandelFoutAf(Optional TypePad As String = Nothing, Optional Pad As String = Nothing, Optional ExtraInformatie As String = Nothing) As DialogResult
      Dim Bron As String = Microsoft.VisualBasic.Err.Source
      Dim FoutCode As Integer = Microsoft.VisualBasic.Err.Number
      Dim FoutOmschrijving As String = Microsoft.VisualBasic.Err.Description
      Static Keuze As DialogResult = DialogResult.Retry

      Try
         Dim Bericht As New StringBuilder

         Microsoft.VisualBasic.Err.Clear()

         Bericht.Append($"{MaakFoutOmschrijvingOp(FoutOmschrijving)}{NewLine}Foutcode: {FoutCode}")
         If Bron IsNot Nothing Then Bericht.Append($"{NewLine}Bron: ""{Bron}""")
         Bericht.Append($"{NewLine}Procedure: ""{(New StackTrace).GetFrames().Skip(1).First().GetMethod().Name}""")
         If TypePad IsNot Nothing AndAlso Pad IsNot Nothing Then Bericht.Append($"{NewLine}{TypePad}""{Path.GetFullPath(Pad)}""")
         If ExtraInformatie IsNot Nothing Then Bericht.Append($"{NewLine}{ExtraInformatie}")

         Keuze = MessageBox.Show(Bericht.ToString(), $"{My.Application.Info.Title} ({BITS_MODUS})", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)

         Return Keuze
      Catch
         End
      Finally
         If Keuze = DialogResult.Abort Then End
      End Try

      Return Nothing
   End Function

   'Deze procedure stuurt de instellingen voor dit programma terug.
   Public Function Instellingen(Optional InstellingenPad As String = Nothing) As InstellingenDefinitie
      Try
         Static ProgrammaInstellingen As InstellingenDefinitie = StandaardInstellingen()

         If Not InstellingenPad = Nothing Then
            If File.Exists(InstellingenPad) Then
               ProgrammaInstellingen = LaadInstellingen(InstellingenPad)
            Else
               If MessageBox.Show($"Kan het instellingenbestand niet vinden.{NewLine}Instellingenbestand: ""{InstellingenPad}""{NewLine}Dit bestand genereren?{NewLine}Huidig pad: ""{Directory.GetCurrentDirectory()}""", My.Application.Info.Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                  BewaarInstellingen(InstellingenPad, StandaardInstellingen(), "De standaardinstellingen zijn weggeschreven naar:")
                  ProgrammaInstellingen = LaadInstellingen(InstellingenPad)
               End If
            End If
         End If

         Return ProgrammaInstellingen
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure toont instellingsbestand gerelateerde foutmeldingen.
   Private Function InstellingenFout(Bericht As StringBuilder, Optional InstellingenPad As String = Nothing, Optional Sectie As String = Nothing, Optional Regel As String = Nothing, Optional Fataal As Boolean = False) As Integer
      Try
         If Sectie IsNot Nothing Then Bericht.Append($"{NewLine}Sectie: {Sectie}")
         If Regel IsNot Nothing Then Bericht.Append($"{NewLine}Regel: ""{Regel}""")
         If InstellingenPad IsNot Nothing Then Bericht.Append($"{NewLine}Instellingenbestand: ""{InstellingenPad}""")

         If Fataal Then
            Return MessageBox.Show(Bericht.ToString(), My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
         Else
            Return MessageBox.Show(Bericht.ToString(), My.Application.Info.Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
         End If
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure laadt de instellingen voor dit programma.
   Private Function LaadInstellingen(InstellingenPad As String) As InstellingenDefinitie
      Try
         Dim Afbreken As Boolean = False
         Dim GeladenInstellingen As New List(Of String)
         Dim ParameterNaam As String = Nothing
         Dim ProgrammaInstellingen As InstellingenDefinitie = StandaardInstellingen()
         Dim RecensteGeldigeSectie As String = Nothing
         Dim Sectie As String = Nothing
         Dim VerbindingsInformatie As String = Nothing
         Dim VerwerkteParameters As New List(Of String)
         Dim VerwerkteSecties As New List(Of String)

         Do
            Try
               With ProgrammaInstellingen
                  .Bestand = InstellingenPad
                  For Each Regel As String In File.ReadAllLines(.Bestand)
                     If Afbreken Then Exit For

                     If Not Regel.Trim().StartsWith(COMMENTAAR_TEKEN) Then
                        If IS_INSTELLINGS_SECTIE(Regel) Then
                           Regel = Regel.Trim()
                           RecensteGeldigeSectie = Sectie
                           Sectie = Regel.Substring(SECTIE_NAAM_BEGIN.Length, Regel.Length - (SECTIE_NAAM_BEGIN.Length + SECTIE_NAAM_EINDE.Length)).ToUpper()
                           If VerwerkteSecties.Contains(Sectie) Then
                              If InstellingenFout(New StringBuilder("Sectie is meerdere keren aanwezig."), InstellingenPad, Sectie, Regel) = DialogResult.Cancel Then Afbreken = True
                           Else
                              VerwerkteSecties.Add(Sectie)
                           End If
                           VerwerkteParameters.Clear()
                        Else
                           Select Case Sectie
                              Case "BATCH", "EXPORT", "QUERY", "VOORBEELD"
                                 If Not Regel.Trim() = Nothing Then
                                    LeesParameter(Regel, ParameterNaam)
                                    If VerwerkteParameters.Contains(ParameterNaam) Then
                                       If InstellingenFout(New StringBuilder("Parameter is meerdere keren aanwezig."), InstellingenPad, Sectie, Regel) = DialogResult.Cancel Then Afbreken = True
                                    Else
                                       VerwerkteParameters.Add(ParameterNaam)
                                    End If
                                 End If
                           End Select
                        End If

                        Select Case Sectie
                           Case "BATCH"
                              If Not (IS_INSTELLINGS_SECTIE(Regel) OrElse Regel.Trim() = Nothing) Then
                                 If Not VerwerkBatchInstellingen(Regel, Sectie, ProgrammaInstellingen) Then Afbreken = True
                              End If
                           Case "EMAILTEKST"
                              If Not IS_INSTELLINGS_SECTIE(Regel) Then .EMailTekst.Append($"{Regel}{NewLine}")
                           Case "EXPORT"
                              If Not (IS_INSTELLINGS_SECTIE(Regel) OrElse Regel.Trim() = Nothing) Then
                                 If Not VerwerkExportInstellingen(Regel, Sectie, ProgrammaInstellingen) Then Afbreken = True
                              End If
                           Case "QUERY"
                              If Not (IS_INSTELLINGS_SECTIE(Regel) OrElse Regel.Trim() = Nothing) Then
                                 If Not VerwerkQueryInstellingen(Regel, Sectie, ProgrammaInstellingen) Then Afbreken = True
                              End If
                           Case "VERBINDING"
                              If Not (IS_INSTELLINGS_SECTIE(Regel) OrElse Regel.Trim() = Nothing) Then .VerbindingsInformatie.Append(Regel.Trim())
                           Case "VOORBEELD"
                              If Not (IS_INSTELLINGS_SECTIE(Regel) OrElse Regel.Trim() = Nothing) Then
                                 If Not VerwerkVoorbeeldInstellingen(Regel, Sectie, ProgrammaInstellingen) Then Afbreken = True
                              End If
                           Case Else
                              If Not Regel.Trim() = Nothing Then
                                 If IS_INSTELLINGS_SECTIE(Regel) Then
                                    Sectie = RecensteGeldigeSectie
                                    If InstellingenFout(New StringBuilder("Niet herkende sectie."), InstellingenPad, Sectie, Regel) = DialogResult.Cancel Then Afbreken = True
                                 Else
                                    If InstellingenFout(New StringBuilder("Niet herkende parameter."), InstellingenPad, Sectie, Regel) = DialogResult.Cancel Then Afbreken = True
                                 End If
                              End If
                        End Select
                     End If
                  Next Regel

                  If .VerbindingsInformatie.ToString() = Nothing AndAlso Not Afbreken Then
                     VerbindingsInformatie = VraagVerbindingsInformatie().Trim()
                     If Not VerbindingsInformatie = Nothing Then
                        .VerbindingsInformatie = New StringBuilder(VerbindingsInformatie)
                        BewaarInstellingen(InstellingenPad, ProgrammaInstellingen, "De instellingen zijn weggeschreven naar:")
                     End If
                  End If

                  .VerbindingsInformatie = New StringBuilder(MaakVerbindingsInformatieOp(.VerbindingsInformatie.ToString()))
               End With
               Exit Do
            Catch
               If HandelFoutAf() = DialogResult.Ignore Then Exit Do
            End Try
         Loop

         Return ProgrammaInstellingen
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure stuurt de waarde en de naam van een instellingenparameter in de opgegeven regel terug.
   Private Function LeesParameter(Regel As String, ByRef ParameterNaam As String) As String
      Try
         Dim Positie As Integer = Regel.IndexOf(WAARDE_TEKEN)
         Dim Waarde As String = Nothing

         ParameterNaam = Nothing
         If Positie >= 0 Then
            ParameterNaam = Regel.Substring(0, Positie).Trim().ToLower()
            Waarde = Regel.Substring(Positie + 1).Trim()
         End If

         Return Waarde
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure maakt de opgegeven foutomschrijving op.
   Private Function MaakFoutOmschrijvingOp(FoutOmschrijving As String) As String
      Try
         Dim Omschrijving As String = FoutOmschrijving.Trim()

         Do
            Select Case Omschrijving.ToCharArray().Last()
               Case Microsoft.VisualBasic.ControlChars.Cr, Microsoft.VisualBasic.ControlChars.Lf
                  Omschrijving = Omschrijving.Substring(0, Omschrijving.Length - 1).Trim()
               Case Else
                  Exit Do
            End Select
         Loop
         If Not Omschrijving.EndsWith("."c) Then Omschrijving.Append("."c)

         Return Omschrijving
      Catch
      End Try

      Return Nothing
   End Function

   'Deze procedure controleert de opgegeven verbindingsinformatie en maakt deze op.
   Private Function MaakVerbindingsInformatieOp(VerbindingsInformatie As String) As String
      Try
         Dim HuidigAfscheidingsTeken As Char = Nothing
         Dim HuidigStringTeken As Char = Nothing
         Dim OpgemaakteVerbindingsInformatie As New StringBuilder
         Dim Parameter As String = Nothing
         Dim ParameterBegin As New Integer
         Dim ParameterNaam As String = Nothing
         Dim ParameterNamen As New List(Of String)
         Dim Positie As Integer = 0
         Dim Teken As New Char
         Dim Waarde As String = Nothing

         If Not VerbindingsInformatie.Trim() = Nothing Then
            ParameterBegin = Positie
            If Not VerbindingsInformatie.Trim().EndsWith(VERBINDING_PARAMETER_TEKEN) Then VerbindingsInformatie = $"{VerbindingsInformatie}{VERBINDING_PARAMETER_TEKEN}"

            Do Until Positie >= VerbindingsInformatie.Length
               Teken = VerbindingsInformatie.Chars(Positie)
               If TEKENREEKS_TEKENS.Contains(Teken) Then
                  If HuidigStringTeken = Nothing Then
                     HuidigStringTeken = Teken
                  ElseIf Teken = HuidigStringTeken Then
                     HuidigStringTeken = Nothing
                  End If
               ElseIf (Teken = VERBINDING_AFSCHEIDING_BEGIN OrElse Teken = VERBINDING_AFSCHEIDING_EINDE) AndAlso HuidigStringTeken = Nothing Then
                  If Teken = VERBINDING_AFSCHEIDING_BEGIN AndAlso HuidigAfscheidingsTeken = Nothing Then
                     HuidigAfscheidingsTeken = Teken
                  ElseIf Teken = VERBINDING_AFSCHEIDING_EINDE AndAlso HuidigAfscheidingsTeken = VERBINDING_AFSCHEIDING_BEGIN Then
                     HuidigAfscheidingsTeken = Nothing
                  Else
                     MessageBox.Show($"Ongeldig afscheidingsteken in verbindingsinformatie: ""{Teken}"".", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                     Return Nothing
                  End If
               ElseIf Teken = VERBINDING_PARAMETER_TEKEN AndAlso HuidigAfscheidingsTeken = Nothing AndAlso HuidigStringTeken = Nothing Then
                  Parameter = VerbindingsInformatie.Substring(ParameterBegin, Positie - ParameterBegin)

                  If Not Parameter.Contains(WAARDE_TEKEN) Then
                     MessageBox.Show($"Ongeldige parameter aanwezig in verbindingsinformatie: ""{Parameter}"". Verwacht teken: {WAARDE_TEKEN}", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                     Return Nothing
                  End If

                  Waarde = LeesParameter(Parameter, ParameterNaam)

                  If ParameterNamen.Contains(ParameterNaam) Then
                     MessageBox.Show($"Parameter meerdere malen aanwezig in verbindingsinformatie: ""{Parameter}"".", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                     Return Nothing
                  Else
                     ParameterNamen.Add(ParameterNaam)
                  End If

                  ParameterBegin = Positie + 1
                  OpgemaakteVerbindingsInformatie.Append($"{ParameterNaam}{WAARDE_TEKEN}{Waarde.Trim()}{VERBINDING_PARAMETER_TEKEN}")
               End If

               Positie += 1
            Loop

            If Not HuidigStringTeken = Nothing Then
               MessageBox.Show($"Niet afgesloten tekenreekswaarde in verbindingsgegevens. Verwacht teken: {HuidigStringTeken}", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
               Return Nothing
            End If
         End If

         Return OpgemaakteVerbindingsInformatie.ToString()
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure wordt uitgevoerd wanneer dit programma wordt gestart.
   Public Sub Main()
      Try
         Dim InstellingenPad As String = Nothing

         Directory.SetCurrentDirectory(My.Application.Info.DirectoryPath)

         With OpdrachtRegelParameters(OPDRACHT_REGEL)
            If .Verwerkt Then
               If .InstellingenPad.Trim().StartsWith(PARAMETER_TEKEN) Then
                  InstellingenPad = .InstellingenPad.Trim().Substring(0, PARAMETER_TEKEN.ToString().Length + 1).Trim(""""c)
                  If InstellingenPad = Nothing Then
                     MessageBox.Show("Kan de instellingen niet bewaren. Geen doel bestand opgegeven.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                  Else
                     BewaarInstellingen(InstellingenPad, StandaardInstellingen(), "De standaardinstellingen zijn weggeschreven naar:")
                  End If
               ElseIf Not .SessiesPad = Nothing Then
                  SESSIE_PARAMETERS.Clear()
                  VerwerkSessieLijst(.SessiesPad)
               Else
                  SESSIE_PARAMETERS.Clear()
                  VoerSessieUit(OPDRACHT_REGEL)
               End If
            End If
         End With
      Catch
         HandelFoutAf()
      Finally
         Verbinding(, VerbindingSluiten:=True)
         SluitAlleVensters()
      End Try
   End Sub

   'Deze procedure controleert de queryparameter invoer en stuurt eventueel de index van een onjuist ingevuld veld en een foutomschrijving terug.
   Private Function OngeldigeParameterInvoer(Optional ByRef FoutInformatie As String = Nothing) As Integer?
      Try
         Dim Lengte As New Integer
         Dim OngeldigVeld As Integer? = Nothing

         For ParameterIndex As Integer = 0 To QueryParameters().Count - 1
            With QueryParameters()(ParameterIndex)
               If .Masker = Nothing Then
                  Lengte = .Invoer.Length
               Else
                  Lengte = If(.LengteIsVariabel, ParameterInvoerLengte(ParameterIndex), .Masker.Length)
                  For Positie As Integer = 0 To Lengte - 1
                     FoutInformatie = ParameterMaskerTekenGeldig(.Invoer.Chars(Positie), .Masker.Chars(Positie))
                     If FoutInformatie IsNot Nothing Then
                        FoutInformatie = $"{NewLine}""{FoutInformatie}"".{NewLine}Teken positie: {Positie + 1}."
                        OngeldigVeld = ParameterIndex
                        Exit For
                     End If
                  Next Positie
               End If

               If OngeldigVeld IsNot Nothing Then Exit For
               QueryParameters(, ParameterIndex, .Invoer.Substring(0, Lengte))
            End With
         Next ParameterIndex

         If OngeldigVeld IsNot Nothing Then
            For ParameterIndex As Integer = 0 To QueryParameters().Count - 1
               QueryParameters(, ParameterIndex, "")
            Next ParameterIndex
         End If

         Return OngeldigVeld
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure beheert de huidige sessie's opdrachtregelparameters.
   Public Function OpdrachtRegelParameters(Optional SessieParameters As String = Nothing) As OpdrachtRegelParametersDefinitie
      Try
         Dim Bericht As New StringBuilder
         Dim Extensie As String = Nothing
         Dim Extensies As New List(Of String)
         Dim Positie As Integer = 0
         Static HuidigeOpdrachtRegelParameters As New OpdrachtRegelParametersDefinitie With {.InstellingenPad = "", .QueryPad = Nothing, .SessiesPad = Nothing, .Verwerkt = True}

         With HuidigeOpdrachtRegelParameters
            If SessieParameters IsNot Nothing Then
               Extensies.Clear()

               Positie = SessieParameters.IndexOf(New String(PARAMETER_TEKEN, 2))
               If Positie >= 0 Then
                  .InstellingenPad = SessieParameters.Substring(Positie + PARAMETER_TEKEN.ToString().Length)
               Else
                  For Each Parameter As String In SessieParameters.Split(PARAMETER_TEKEN)
                     If Not Parameter.Trim() = Nothing Then
                        Parameter = Parameter.Trim(""""c)
                        Extensie = Path.GetExtension(Parameter).ToLower()

                        If Not Extensies.Contains(Extensie) Then
                           Extensies.Add(Extensie)

                           Select Case Extensie
                              Case ".ini"
                                 .InstellingenPad = Parameter
                              Case ".lst"
                                 .SessiesPad = Parameter
                              Case ".txt"
                                 .QueryPad = Parameter
                              Case Else
                                 If Not Parameter.Trim() = Nothing Then
                                    Bericht = New StringBuilder($"Niet herkende opdrachtregelparameter: ""{Parameter}"".")
                                    If VerwerkSessieLijst() = Nothing Then
                                       MessageBox.Show(Bericht.ToString(), My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                    Else
                                       Bericht.Append($"{NewLine}Sessielijst: ""{VerwerkSessieLijst()}"".")
                                       If MessageBox.Show(Bericht.ToString(), My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation) = DialogResult.Cancel Then SessiesAfbreken = True
                                    End If
                                    .Verwerkt = False
                                 End If
                           End Select
                        Else
                           Bericht = New StringBuilder("Er kan maar een instellingenbestand en/of query tegelijk opgegeven worden.")
                           If Not VerwerkSessieLijst() = Nothing Then Bericht.Append($"{NewLine}Sessielijst: ""{VerwerkSessieLijst()}"".")
                           MessageBox.Show(Bericht.ToString(), My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                           .Verwerkt = False
                        End If
                     End If
                  Next Parameter
               End If
            End If
         End With

         Return HuidigeOpdrachtRegelParameters
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure stuurt de lengte van de invoer voor de opgegeven queryparameter terug.
   Private Function ParameterInvoerLengte(ParameterIndex As Integer) As Integer
      Try
         Dim Lengte As Integer = 0

         With QueryParameters()(ParameterIndex)
            For Positie As Integer = 0 To .Invoer.Length - 1
               If Not .Invoer.Chars(Positie) = .Masker.Chars(Positie) Then Lengte = Positie + 1
            Next Positie
         End With

         Return Lengte
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure vergelijkt het opgegeven teken met het opgegeven queryparametermaskerteken.
   Public Function ParameterMaskerTekenGeldig(Teken As Char, MaskerTeken As Char) As String
      Try
         Dim Geldig As String = Nothing

         Select Case MaskerTeken
            Case MASKER_CIJFER
               If Not (Teken >= "0"c AndAlso Teken <= "9"c) Then Geldig = "Cijfer verwacht."
            Case MASKER_HOOFDLETTER
               If Not (Teken >= "A"c AndAlso Teken <= "Z"c) Then Geldig = "Hoofdletter verwacht."
            Case Else
               If Not Teken = MaskerTeken Then Geldig = $"Vast maskerteken ""{ MaskerTeken}"" verwacht."
         End Select

         Return Geldig
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure controleert de door de gebruiker ingevoerde parameters en stuurt het resultaat terug.
   Public Function ParametersGeldig(ParameterVelden As List(Of Object)) As Boolean
      Try
         Dim FoutInformatie As String = Nothing
         Dim Geldig As Boolean = False
         Dim OngeldigeVeldIndex As Integer? = Nothing

         For ParameterIndex As Integer = 0 To ParameterVelden.Count - 1
            QueryParameters(, ParameterIndex, DirectCast(ParameterVelden(ParameterIndex), TextBox).Text)
         Next ParameterIndex

         OngeldigeVeldIndex = OngeldigeParameterInvoer(FoutInformatie)
         Geldig = (OngeldigeVeldIndex Is Nothing)

         If Not Geldig Then
            With DirectCast(ParameterVelden(OngeldigeVeldIndex.Value), TextBox)
               If .Visible Then
                  FoutInformatie = $"Dit veld is niet volledig of onjuist ingevuld:{FoutInformatie}"
               Else
                  FoutInformatie = $"Onzichtbare parameter #{OngeldigeVeldIndex.Value} is niet volledig of onjuist ingevuld:{FoutInformatie}"
               End If
               MessageBox.Show(FoutInformatie, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
               If .Visible Then .Focus()
            End With
         End If

         Return Geldig
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure toont parameter en/of symbool gerelateerde foutmeldingen.
   Private Sub ParameterSymboolFout(Bericht As StringBuilder, Optional ParameterIndex As Integer? = Nothing)
      Try
         If ParameterIndex IsNot Nothing Then
            Bericht.Append($"{NewLine}Parameter definitie: #{ParameterIndex + 1}")
            With QueryParameters()(ParameterIndex.Value)
               If .ParameterNaam IsNot Nothing Then Bericht.Append($"{NewLine} Naam: ""{ .ParameterNaam}""")
               If .Invoer IsNot Nothing Then Bericht.Append($"{NewLine}Invoer: ""{ .Invoer}""")
               If .StandaardWaarde IsNot Nothing Then Bericht.Append($"{NewLine}Standaardwaarde: ""{ .StandaardWaarde}""")
               If .Masker IsNot Nothing Then Bericht.Append($"{NewLine}Masker: ""{ .Masker}""")
            End With
         End If
         If Not Query().Pad = Nothing Then Bericht.Append($"{ NewLine}Query: ""{Query().Pad}""")
         MessageBox.Show(Bericht.ToString(), My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure laadt de opgegeven query of stuurt een al geladen query terug.
   Public Function Query(Optional QueryPad As String = Nothing) As QueryDefinitie
      Try
         Static HuidigeQuery As New QueryDefinitie With {.Code = Nothing, .Geopend = False, .Pad = Nothing}

         Do
            With HuidigeQuery
               Try
                  .Geopend = False

                  If Not QueryPad = Nothing Then
                     .Code = New String((From ByteO In File.ReadAllBytes(QueryPad) Select ToChar(ByteO)).ToArray())
                     .Pad = QueryPad
                     .Geopend = True
                  End If

                  Exit Do
               Catch
                  .Geopend = False
                  If HandelFoutAf(TypePad:="Query pad: ", Pad:=QueryPad) = DialogResult.Ignore Then Exit Do
               End Try
            End With
         Loop

         Return HuidigeQuery
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure doorzoekt de opgegeven query op parameterdefinities of stuurt een eerder gevonden parameterdefinitie terug.
   Public Function QueryParameters(Optional QueryCode As String = Nothing, Optional ParameterIndex As Integer = 0, Optional Invoer As String = Nothing) As List(Of QueryParameterDefinitie)
      Try
         Dim Definitie As String = Nothing
         Dim DefinitieBegin As New Integer
         Dim DefinitieEinde As New Integer
         Dim Elementen() As String = {}
         Dim NieuweParameter As QueryParameterDefinitie = Nothing
         Dim OnverwerkteCode As String = Nothing
         Static Parameters As New List(Of QueryParameterDefinitie)

         If Invoer IsNot Nothing Then
            If Parameters.Count > 0 Then
               NieuweParameter = Parameters(ParameterIndex)
               NieuweParameter.Invoer = Invoer
               Parameters.RemoveAt(ParameterIndex)
               Parameters.Insert(ParameterIndex, NieuweParameter)
            End If
         ElseIf QueryCode IsNot Nothing Then
            Parameters.Clear()

            OnverwerkteCode = QueryCode
            Do
               DefinitieBegin = OnverwerkteCode.IndexOf(DEFINITIE_TEKENS)
               If DefinitieBegin >= 0 Then
                  DefinitieEinde = OnverwerkteCode.IndexOf(DEFINITIE_TEKENS, DefinitieBegin + DEFINITIE_TEKENS.Length)
                  If DefinitieEinde >= 0 Then
                     Definitie = OnverwerkteCode.Substring(DefinitieBegin + DEFINITIE_TEKENS.Length, (DefinitieEinde - DefinitieBegin) - DEFINITIE_TEKENS.Length)
                     OnverwerkteCode = OnverwerkteCode.Substring(DefinitieEinde + DEFINITIE_TEKENS.Length)

                     NieuweParameter = New QueryParameterDefinitie With {.Commentaar = "", .Invoer = "", .Lengte = 0, .LengteIsVariabel = False, .Masker = "", .ParameterNaam = "", .Positie = 0, .StandaardWaarde = "", .VeldIsZichtbaar = False}
                     With NieuweParameter
                        .Lengte = DEFINITIE_TEKENS.Length + Definitie.Length + DEFINITIE_TEKENS.Length
                        .Positie = DefinitieBegin

                        Elementen = Definitie.Split(ELEMENT_TEKEN)
                        If Elementen.Count > MAXIMUM_AANTAL_ELMENTEN Then ParameterSymboolFout(New StringBuilder("Teveel elementen, deze worden genegeerd."), Parameters.Count - 1)
                        ReDim Preserve Elementen(0 To MAXIMUM_AANTAL_ELMENTEN - 1)

                        .ParameterNaam = Elementen(ParameterDefinitieOpsomming.NaamElement)

                        .VeldIsZichtbaar = (.ParameterNaam IsNot Nothing)

                        .Masker = Elementen(ParameterDefinitieOpsomming.MaskerElement)

                        .LengteIsVariabel = .Masker.StartsWith(VARIABELE_LENGTE_TEKEN)
                        If .LengteIsVariabel Then .Masker = .Masker.Substring(1)

                        .StandaardWaarde = VervangSymbolen(Elementen(ParameterDefinitieOpsomming.StandaardWaardeElement))
                        If Not .Masker = Nothing Then If .StandaardWaarde.Length > .Masker.Length Then ParameterSymboolFout(New StringBuilder("De standaardwaarde is langer dan het masker. De overtollige tekens worden verwijderd."), Parameters.Count - 1)

                        If Elementen(ParameterDefinitieOpsomming.CommentaarElement) IsNot Nothing Then .Commentaar = Elementen(ParameterDefinitieOpsomming.CommentaarElement)

                        .Invoer = .StandaardWaarde
                     End With

                     Parameters.Add(NieuweParameter)
                  Else
                     ParameterSymboolFout(New StringBuilder("Geen einde markering. Deze wordt genegeerd."), Parameters.Count - 1)
                     Exit Do
                  End If
               Else
                  Exit Do
               End If
            Loop
         End If

         Return Parameters
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure handelt eventuele queryresultaat leesfouten af.
   Public Function QueryResultaatLeesFout(Optional Rij As Integer? = Nothing, Optional Kolom As Integer? = Nothing, Optional KolomNaam As String = Nothing, Optional VraagNietOmKeuze As Boolean = True) As DialogResult
      Dim Bron As String = Microsoft.VisualBasic.Err.Source
      Dim FoutCode As Integer = Microsoft.VisualBasic.Err.Number
      Dim FoutOmschrijving As String = Microsoft.VisualBasic.Err.Description

      Microsoft.VisualBasic.Err.Clear()

      Try
         Dim Bericht As New StringBuilder
         Static Keuze As DialogResult = DialogResult.Retry

         If Not VraagNietOmKeuze Then
            Bericht.Append($"Er is een fout opgetreden bij het uitlezen van het queryresultaat.{NewLine}Rij: {Rij}{NewLine}Kolom: {Kolom}{NewLine}Kolom naam: {KolomNaam}{NewLine}Omschrijving: {MaakFoutOmschrijvingOp(FoutOmschrijving)}{NewLine}Foutcode: {FoutCode}")
            If Bron IsNot Nothing Then Bericht.Append($"{NewLine}Bron: {Bron}")

            Keuze = MessageBox.Show(Bericht.ToString(), My.Application.Info.Title, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
         End If

         Return Keuze
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure stuurt het queryresultaat terug als tekst.
   Public Function QueryResultaatTekst(Resultaat As QueryResultaatDefinitie) As String
      Try
         Dim Breedte As New Integer
         Dim LaatsteRegel As New Integer
         Dim ResultaatTekst As New StringBuilder
         Dim Tekst As String = Nothing

         With Resultaat
            If .Tabel IsNot Nothing Then
               If Instellingen().VoorbeeldRegels = ALLE_REGELS OrElse Instellingen().VoorbeeldRegels > .Tabel.GetUpperBound(0) + 1 Then
                  LaatsteRegel = .Tabel.GetUpperBound(0)
               Else
                  LaatsteRegel = Instellingen().VoorbeeldRegels - 1
               End If

               For Rij As Integer = 0 To LaatsteRegel
                  For Kolom As Integer = 0 To .Tabel.GetUpperBound(1)
                     Breedte = .KolomBreedte(Kolom)
                     Tekst = If(.Tabel(Rij, Kolom) = Nothing, "", .Tabel(Rij, Kolom))
                     Tekst = Tekst.Replace(Microsoft.VisualBasic.ControlChars.Cr, " "c)
                     Tekst = Tekst.Replace(Microsoft.VisualBasic.ControlChars.Lf, " "c)
                     Tekst = Tekst.Replace(Microsoft.VisualBasic.ControlChars.Tab, " "c)

                     If Not Instellingen().VoorbeeldKolomBreedte = GEEN_MAXIMALE_BREEDTE Then
                        If .KolomBreedte(Kolom) > Instellingen().VoorbeeldKolomBreedte Then
                           Breedte = Instellingen().VoorbeeldKolomBreedte
                           If Tekst.Length > Instellingen.VoorbeeldKolomBreedte Then Tekst = Tekst.Substring(0, Instellingen().VoorbeeldKolomBreedte)
                        End If
                     End If

                     ResultaatTekst.Append($"{If(.RechtsUitlijnen(Kolom), Tekst.PadLeft(Breedte), Tekst.PadRight(Breedte))} ")
                  Next Kolom
                  ResultaatTekst.Append(NewLine)
                  Application.DoEvents()
                  If Application.OpenForms.Count = 0 Then Exit For
               Next Rij
            End If
         End With

         Return ResultaatTekst.ToString()
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure beheert de queryresulaten.
   Public Function QueryResultaten(Optional NieuwQueryResultaat As Recordset = Nothing, Optional ResultatenVerwijderen As Boolean = False) As List(Of QueryResultaatDefinitie)
      Try
         Dim NieuwResultaat As QueryResultaatDefinitie = Nothing
         Dim TijdelijkeTabel As New List(Of List(Of String))
         Static Resultaten As New List(Of QueryResultaatDefinitie)

         If Not NieuwQueryResultaat Is Nothing Then
            With NieuwQueryResultaat
               If RECORDSET_GEOPEND(NieuwQueryResultaat) AndAlso Not .BOF Then
                  NieuwResultaat = New QueryResultaatDefinitie With {.KolomBreedte = New List(Of Integer), .RechtsUitlijnen = New List(Of Boolean), .Tabel = {{}}}

                  TijdelijkeTabel.Add(New List(Of String))
                  For Kolom As Integer = 0 To .Fields.Count - 1
                     TijdelijkeTabel.Last().Add(.Fields.Item(Kolom).Name.Trim())
                     NieuwResultaat.KolomBreedte.Add(TijdelijkeTabel.Last()(Kolom).Length)
                     NieuwResultaat.RechtsUitlijnen.Add(Not LINKS_UITGELIJNDE_DATA_TYPES.Contains(.Fields.Item(Kolom).Type))
                  Next Kolom

                  TijdelijkeTabel.Add(New List(Of String))
                  Do While RECORDSET_GEOPEND(NieuwQueryResultaat) AndAlso Not .EOF
                     For Kolom As Integer = 0 To .Fields.Count - 1
                        Do
                           Try
                              TijdelijkeTabel.Last().Add(If(IsDBNull(.Fields.Item(Kolom).Value), "", CStr(.Fields.Item(Kolom).Value).Trim()))
                              If TijdelijkeTabel.Last()(Kolom).Length > NieuwResultaat.KolomBreedte(Kolom) Then NieuwResultaat.KolomBreedte(Kolom) = TijdelijkeTabel.Last()(Kolom).Length + 1
                              Exit Do
                           Catch
                              If QueryResultaatLeesFout(TijdelijkeTabel.Count - 1, Kolom, TijdelijkeTabel.First()(Kolom), VraagNietOmKeuze:=False) = DialogResult.Abort Then Exit Do
                              If QueryResultaatLeesFout() = DialogResult.Ignore Then Exit Do
                           End Try
                        Loop
                        If QueryResultaatLeesFout() = DialogResult.Abort Then Exit Do
                     Next Kolom
                     .MoveNext()
                     TijdelijkeTabel.Add(New List(Of String))
                  Loop

                  ReDim NieuwResultaat.Tabel(0 To TijdelijkeTabel.Count - 1, 0 To .Fields.Count - 1)
                  For Rij As Integer = 0 To NieuwResultaat.Tabel.GetUpperBound(0) - 1
                     For Kolom As Integer = 0 To NieuwResultaat.Tabel.GetUpperBound(1)
                        NieuwResultaat.Tabel(Rij, Kolom) = TijdelijkeTabel(Rij)(Kolom)
                     Next Kolom
                  Next Rij

                  Resultaten.Add(NieuwResultaat)
               End If
            End With
         ElseIf ResultatenVerwijderen Then
            Resultaten.Clear()
         End If

         Return Resultaten
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure sluit alle eventueel geopende vensters af.
   Public Sub SluitAlleVensters()
      Try
         For Each Venster As Form In Application.OpenForms
            Venster.Close()
         Next Venster
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure stuurt de standaardinstellingen voor dit programma terug.
   Private Function StandaardInstellingen() As InstellingenDefinitie
      Try
         Return New InstellingenDefinitie With {
            .BatchBereik = Nothing,
            .BatchInteractief = False,
            .BatchQueryPad = Nothing,
            .Bestand = "Qa.ini",
            .EMailTekst = New StringBuilder(),
            .ExportAfzender = Nothing,
            .ExportAutoOpenen = False,
            .ExportAutoOverschrijven = False,
            .ExportAutoVerzenden = False,
            .ExportCCOntvanger = Nothing,
            .ExportKolomAanvullen = False,
            .ExportOnderwerp = Nothing,
            .ExportOntvanger = Nothing,
            .ExportStandaardPad = ".\Export.xls",
            .QueryAutoSluiten = False,
            .QueryAutoUitvoeren = False,
            .QueryRecordSets = False,
            .QueryTimeout = 10,
            .VerbindingsInformatie = New StringBuilder(),
            .VoorbeeldKolomBreedte = GEEN_MAXIMALE_BREEDTE,
            .VoorbeeldRegels = 10
            }
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure stuurt de status van het queryresultaat terug nadat een query is uitgevoerd.
   Public Function StatusNaQuery(ResultaatIndex As Integer) As String
      Try
         Dim AantalKolommen As Integer = 0
         Dim AantalResultaten As Integer = 0
         Dim AantalRijen As Integer = 0
         Dim Status As New StringBuilder

         If QueryResultaten.Count > 0 Then
            With QueryResultaten()(ResultaatIndex)
               If .Tabel IsNot Nothing Then
                  AantalKolommen = .Tabel.GetUpperBound(1) + 1
                  AantalRijen = .Tabel.GetUpperBound(0)
                  If AantalRijen < 0 Then AantalKolommen = 0
                  AantalResultaten = QueryResultaten().Count
               End If

               Status.Append($"Query uitgevoerd: {AantalRijen} {If(AantalRijen = 1, "rij", "rijen")} en {AantalKolommen} { If(AantalKolommen = 1, "kolom", "kolommen")}.")
               If AantalResultaten > 1 Then Status.Append($" Resultaat {ResultaatIndex + 1} van {AantalResultaten}.")
               If Instellingen().VoorbeeldRegels >= 0 Then Status.Append($" Voorbeeld limiet: {Instellingen().VoorbeeldRegels} {If(Instellingen().VoorbeeldRegels = 1, "regel", "regels")}.")
            End With
         End If

         Return Status.ToString()
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure toon het invoer venster en stuurt de invoer van de gebruiker terug.
   Public Function ToonInvoerVenster(Optional Prompt As String = Nothing, Optional Invoer As String = Nothing, Optional ByRef Knop As DialogResult = Nothing, Optional MeerdereRegels As Boolean = False) As String
      Try
         With InvoerVenster
            .PromptLabel.Text = Prompt
            .TekstVeld.Multiline = MeerdereRegels
            .TekstVeld.ScrollBars = If(.TekstVeld.Multiline, ScrollBars.Vertical, ScrollBars.None)
            .TekstVeld.Text = Invoer
            Knop = .ShowDialog()
            Return If(Knop = DialogResult.Cancel, Nothing, .TekstVeld.Text)
         End With
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure toont de programmainformatie.
   Public Sub ToonProgrammaInformatie()
      Try
         With My.Application.Info
            MessageBox.Show(.Description, $"{ .Title} {PROGRAMMA_VERSIE} - door: { .CompanyName}, { .Copyright}", MessageBoxButtons.OK, MessageBoxIcon.Information)
         End With
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure toont het opgegeven queryresultaat.
   Public Sub ToonQueryResultaat(QueryResultaatVeld As TextBox, ResultaatIndex As Integer)
      Try
         Dim ResultaatTekst As String = QueryResultaatTekst(If(QueryResultaten().Count > 0, QueryResultaten()(ResultaatIndex), Nothing))

         ToonStatus($"Bezig met maken van voorbeeld weergave voor queryresultaat...{NewLine}")
         QueryResultaatVeld.Text = ResultaatTekst
         If HuidigInterfaceVenster IsNot Nothing AndAlso HuidigInterfaceVenster.Visible AndAlso QueryResultaatVeld.Text.Length < ResultaatTekst.Length Then MessageBox.Show("Het queryresultaat kan niet volledig worden weergegeven.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Information)
         ToonStatus($"{StatusNaQuery(ResultaatIndex)}{ NewLine}")
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure toont de opgegeven tekst in het opgegeven veld.
   Public Sub ToonStatus(Optional Tekst As String = Nothing, Optional NieuwVeld As TextBox = Nothing)
      Try
         Dim VorigeLengte As New Integer
         Static Veld As TextBox = Nothing

         If NieuwVeld IsNot Nothing Then Veld = NieuwVeld

         If Veld IsNot Nothing AndAlso Not Tekst = Nothing Then
            With Veld
               VorigeLengte = .Text.Length
               .AppendText(Tekst)
               If .Text.Length < VorigeLengte + Tekst.Length Then .Text = Tekst
               .Select(If(Tekst.Length > .Text.Length, .Text.Length, .Text.Length - Tekst.Length), 0)
            End With
         End If

         Application.DoEvents()
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure beheert de verbinding met een database.
   Public Function Verbinding(Optional VerbindingsInformatie As String = Nothing, Optional VerbindingSluiten As Boolean = False, Optional Reset As Boolean = False) As Connection
      Static DataBaseVerbinding As New Connection

      Do
         Try
            If DataBaseVerbinding IsNot Nothing Then
               If Reset Then
                  DataBaseVerbinding.Errors.Clear()
               ElseIf Not VerbindingsInformatie = Nothing Then
                  If Not MaakVerbindingsInformatieOp(VerbindingsInformatie) = Nothing Then DataBaseVerbinding.Open(VerbindingsInformatie)
               ElseIf VerbindingSluiten Then
                  If VERBINDING_GEOPEND(DataBaseVerbinding) Then
                     DataBaseVerbinding.Close()
                     DataBaseVerbinding = Nothing
                  End If
               End If
            End If

            Exit Do
         Catch
            If HandelFoutAf() = DialogResult.Ignore Then Exit Do
         End Try
      Loop

      Return DataBaseVerbinding
   End Function

   'Deze procecure vervangt de symbolen in de opgegeven tekst met de tekst waar ze voor staan.
   Public Function VervangSymbolen(Tekst As String) As String
      Try
         Dim Symbool As String = Nothing
         Dim SymboolBegin As New Integer
         Dim SymboolEinde As New Integer
         Dim TekstMetSymbolen As String = Tekst
         Dim TekstZonderSymbolen As New StringBuilder

         If Tekst = Nothing Then
            TekstZonderSymbolen.Append("")
         Else
            Do
               SymboolBegin = TekstMetSymbolen.IndexOf(SYMBOOL_TEKEN)
               If SymboolBegin < 0 Then
                  TekstZonderSymbolen.Append(TekstMetSymbolen)
                  Exit Do
               Else
                  SymboolEinde = TekstMetSymbolen.IndexOf(SYMBOOL_TEKEN, SymboolBegin + 1)
                  If SymboolEinde < 0 Then
                     TekstZonderSymbolen.Append(TekstMetSymbolen)
                     Exit Do
                  Else
                     TekstZonderSymbolen.Append(TekstMetSymbolen.Substring(0, SymboolBegin))
                     Symbool = TekstMetSymbolen.Substring(SymboolBegin + 1, (SymboolEinde - SymboolBegin) - 1)
                     TekstMetSymbolen = TekstMetSymbolen.Substring(SymboolEinde + 1)

                     If Symbool = Nothing Then
                        ParameterSymboolFout(New StringBuilder("Een leeg symbool is gevonden. Deze wordt genegeerd."))
                     Else
                        TekstZonderSymbolen.Append(VerwerkSymbool(Symbool))
                     End If
                  End If
               End If
            Loop
         End If

         Return TekstZonderSymbolen.ToString()
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure verwerkt de batchinstellingen.
   Private Function VerwerkBatchInstellingen(Regel As String, Sectie As String, ByRef BatchInstellingen As InstellingenDefinitie) As Boolean
      Try
         Dim ParameterNaam As String = Nothing
         Dim Verwerkt As Boolean = True
         Dim Waarde As String = LeesParameter(Regel, ParameterNaam)

         With BatchInstellingen
            Select Case ParameterNaam
               Case "bereik"
                  .BatchBereik = Waarde
               Case "interactief"
                  .BatchInteractief = Boolean.Parse(Waarde)
               Case "querypad"
                  .BatchQueryPad = Waarde.Trim(""""c)
               Case Else
                  If InstellingenFout(New StringBuilder("Niet herkende parameter."), BatchInstellingen.Bestand, Sectie, Regel) = DialogResult.Cancel Then Verwerkt = False
            End Select
         End With

         Return Verwerkt
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure verwerkt de exportinstellingen.
   Private Function VerwerkExportInstellingen(Regel As String, Sectie As String, ByRef ExportInstellingen As InstellingenDefinitie) As Boolean
      Try
         Dim ParameterNaam As String = Nothing
         Dim Verwerkt As Boolean = True
         Dim Waarde As String = LeesParameter(Regel, ParameterNaam)

         With ExportInstellingen
            Select Case ParameterNaam
               Case "afzender"
                  .ExportAfzender = Waarde
               Case "autoopenen"
                  .ExportAutoOpenen = Boolean.Parse(Waarde)
               Case "autooverschrijven"
                  .ExportAutoOverschrijven = Boolean.Parse(Waarde)
               Case "autoverzenden"
                  .ExportAutoVerzenden = Boolean.Parse(Waarde)
               Case "ccontvanger"
                  .ExportCCOntvanger = Waarde
               Case "kolomaanvullen"
                  .ExportKolomAanvullen = Boolean.Parse(Waarde)
               Case "onderwerp"
                  .ExportOnderwerp = Waarde
               Case "ontvanger"
                  .ExportOntvanger = Waarde
               Case "standaardpad"
                  .ExportStandaardPad = Waarde
               Case Else
                  If InstellingenFout(New StringBuilder("Niet herkende parameter."), ExportInstellingen.Bestand, Sectie, Regel) = DialogResult.Cancel Then Verwerkt = False
            End Select
         End With

         Return Verwerkt
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure stuurt de verbindingsinformatie met de opgegeven inloggegevens terug.
   Public Function VerwerkInlogGegevens(Gebruiker As String, Wachtwoord As String, VerbindingsInformatie As String) As String
      Try
         Dim LinkerDeel As String = Nothing
         Dim Positie As New Integer
         Dim RechterDeel As String = Nothing
         Dim VerwerkteInlogGegevens As String = VerbindingsInformatie

         Positie = VerwerkteInlogGegevens.ToUpper().IndexOf(GEBRUIKER_VARIABEL)
         If Positie >= 0 Then
            LinkerDeel = VerwerkteInlogGegevens.Substring(0, Positie)
            RechterDeel = VerwerkteInlogGegevens.Substring(Positie + GEBRUIKER_VARIABEL.Length)
            VerwerkteInlogGegevens = $"{LinkerDeel}{Gebruiker}{RechterDeel}"
         End If

         Positie = VerwerkteInlogGegevens.ToUpper().IndexOf(WACHTWOORD_VARIABEL)
         If Positie >= 0 Then
            LinkerDeel = VerwerkteInlogGegevens.Substring(0, Positie)
            RechterDeel = VerwerkteInlogGegevens.Substring(Positie + WACHTWOORD_VARIABEL.Length)
            VerwerkteInlogGegevens = $"{LinkerDeel}{Wachtwoord}{RechterDeel}"
         End If

         Return VerwerkteInlogGegevens
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure verwerkt de queryinstellingen.
   Private Function VerwerkQueryInstellingen(Regel As String, Sectie As String, ByRef QueryInstellingen As InstellingenDefinitie) As Boolean
      Try
         Dim ParameterNaam As String = Nothing
         Dim Verwerkt As Boolean = True
         Dim Waarde As String = LeesParameter(Regel, ParameterNaam)

         With QueryInstellingen
            Select Case ParameterNaam
               Case "autosluiten"
                  .QueryAutoSluiten = Boolean.Parse(Waarde)
               Case "autouitvoeren"
                  .QueryAutoUitvoeren = Boolean.Parse(Waarde)
               Case "recordsets"
                  .QueryRecordSets = Boolean.Parse(Waarde)
               Case "timeout"
                  .QueryTimeout = Integer.Parse(Waarde)
               Case Else
                  If InstellingenFout(New StringBuilder("Niet herkende parameter."), QueryInstellingen.Bestand, Sectie, Regel) = DialogResult.Cancel Then Verwerkt = False
            End Select
         End With

         Return Verwerkt
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure verwerkt de opgegeven sessielijst.
   Public Function VerwerkSessieLijst(Optional SessieLijstPad As String = Nothing) As String
      Try
         Static HuidigeSessieLijstPad As String = Nothing

         If SessieLijstPad IsNot Nothing Then
            SessiesAfbreken = False
            HuidigeSessieLijstPad = SessieLijstPad
            Do
               Try
                  For Each SessieParameters As String In File.ReadAllLines(HuidigeSessieLijstPad)
                     If SessiesAfbreken Then Exit For
                     If Not SessieParameters.Trim() = Nothing Then VoerSessieUit(SessieParameters)
                  Next SessieParameters
                  Exit Do
               Catch
                  If HandelFoutAf() = DialogResult.Ignore Then Exit Do
               End Try
            Loop
         End If
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure verwerkt de voorbeeldinstellingen.
   Private Function VerwerkVoorbeeldInstellingen(Regel As String, Sectie As String, ByRef VoorbeeldInstellingen As InstellingenDefinitie) As Boolean
      Try
         Dim ParameterNaam As String = Nothing
         Dim Verwerkt As Boolean = True
         Dim Waarde As String = LeesParameter(Regel, ParameterNaam)

         With VoorbeeldInstellingen
            Select Case ParameterNaam
               Case "kolombreedte"
                  .VoorbeeldKolomBreedte = Integer.Parse(Waarde)
               Case "regels"
                  .VoorbeeldRegels = Integer.Parse(Waarde)
               Case Else
                  If InstellingenFout(New StringBuilder("Niet herkende parameter."), VoorbeeldInstellingen.Bestand, Sectie, Regel) = DialogResult.Cancel Then Verwerkt = False
            End Select
         End With

         Return Verwerkt
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure stuurt de door het opgegeven symbool vertegenwoordigde waarde terug.
   Private Function VerwerkSymbool(Symbool As String) As String
      Try
         Dim ArgumentGetal As New Integer
         Dim ArgumentIsGetal As New Boolean
         Dim SymboolArgument As String = Nothing
         Dim SymboolGetal As New Integer
         Dim SymboolIsGetal As Boolean = Integer.TryParse(Symbool, SymboolGetal)
         Dim Waarde As String = Nothing

         If SymboolIsGetal Then
            Waarde = If(SymboolGetal = 0, Path.GetFileNameWithoutExtension(Query().Pad), QueryParameters()(SymboolGetal - 1).Invoer)
         Else
            SymboolArgument = Symbool.Substring(1)
            Symbool = Symbool.ToCharArray().First()
            ArgumentIsGetal = Integer.TryParse(SymboolArgument, ArgumentGetal)

            Select Case Symbool
               Case "D"c
                  Waarde = $"{Date.Now.Day:D2}{Date.Now.Month:D2}{Date.Now.Year:D4}"
               Case "b"c
                  If ArgumentIsGetal Then Waarde = INTERACTIEVE_BATCH_PARAMETERS(ArgumentGetal)
               Case "c"c
                  If ArgumentIsGetal Then Waarde = ToChar(ArgumentGetal)
               Case "d"c
                  Waarde = $"{Date.Now.Day:D2}"
               Case "e"c
                  Waarde = GetEnvironmentVariable(SymboolArgument)
               Case "j"c
                  Waarde = $"{Date.Now.Year:D4}"
               Case "m"c
                  Waarde = $"{Date.Now.Month:D2}"
               Case "s"c
                  If ArgumentIsGetal Then Waarde = SESSIE_PARAMETERS(ArgumentGetal)
               Case Else
                  If Symbool IsNot Nothing Then ParameterSymboolFout(New StringBuilder($"Symbool ""{Symbool}"" is onbekend. Deze wordt genegeerd."))
            End Select
         End If

         Return Waarde
      Catch
         ParameterSymboolFout(New StringBuilder($"Symbool ""{Symbool}"" veroorzaakt de volgende fout: {NewLine}{ Microsoft.VisualBasic.Err.Description}{NewLine}Foutcode: {Microsoft.VisualBasic.Err.Number}"))
      End Try

      Return Nothing
   End Function

   'Deze procecure verwijdert eventuele opmaak uit de opgegeven querycode.
   Private Function VerwijderOpmaak(QueryCode As String, CommentaarBegin As String, CommentaarEinde As String, TekenreeksTekens As String) As String
      Try
         Dim HuidigStringTeken As Char? = Nothing
         Dim InCommentaar As Boolean = False
         Dim Positie As Integer = 0
         Dim QueryZonderOpmaak As New StringBuilder
         Dim Teken As Char? = Nothing

         Do Until Positie >= QueryCode.Length
            Teken = QueryCode.Chars(Positie)

            If InCommentaar Then
               If CommentaarEinde = Nothing Then
                  If QueryCode.Chars(Positie) = Microsoft.VisualBasic.ControlChars.Cr OrElse QueryCode.Chars(Positie) = Microsoft.VisualBasic.ControlChars.Lf Then
                     HuidigStringTeken = Nothing
                     InCommentaar = False
                     Teken = " "c
                  End If
               Else
                  If QueryCode.Substring(Positie, CommentaarEinde.Length) = CommentaarEinde Then
                     HuidigStringTeken = Nothing
                     InCommentaar = False
                     Positie += (CommentaarEinde.Length - 1)
                     Teken = " "c
                  End If
               End If
            Else
               If TekenreeksTekens.Contains(QueryCode.Chars(Positie)) Then
                  If HuidigStringTeken Is Nothing Then
                     HuidigStringTeken = Teken.Value
                  ElseIf Teken = HuidigStringTeken Then
                     HuidigStringTeken = Nothing
                  End If
               ElseIf Positie + CommentaarBegin.Length < QueryCode.Length AndAlso QueryCode.Substring(Positie, CommentaarBegin.Length) = CommentaarBegin Then
                  If HuidigStringTeken Is Nothing Then InCommentaar = True
               End If
            End If

            If Not InCommentaar Then
               If HuidigStringTeken Is Nothing Then
                  If QueryCode.Chars(Positie) = Microsoft.VisualBasic.ControlChars.Cr OrElse QueryCode.Chars(Positie) = Microsoft.VisualBasic.ControlChars.Lf Then Teken = " "c

                  If $"{Microsoft.VisualBasic.ControlChars.Tab} ".Contains(Teken.Value) Then
                     Teken = " "c
                     If QueryZonderOpmaak.ToString().EndsWith(" "c) Then Teken = Nothing
                  End If
               End If

               If Teken IsNot Nothing Then QueryZonderOpmaak.Append(Teken)
            End If

            Positie += 1
         Loop

         Return QueryZonderOpmaak.ToString()
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure voert een querybatch uit.
   Private Sub VoerBatchUit()
      Try
         Dim EersteQuery As New Integer
         Dim EMail As EMailClass = Nothing
         Dim ExportPad As String = Nothing
         Dim ExportPaden As New List(Of String)
         Dim ExportUitgevoerd As Boolean = False
         Dim FoutInformatie As String = Nothing
         Dim LaatsteQuery As New Integer
         Dim Positie As New Integer
         Dim QueryPad As String = Nothing
         Dim QueryPadExtensie As String = Nothing

         TOON_VERBINDINGS_STATUS()

         With Instellingen()
            Positie = .BatchBereik.IndexOf("-"c)
            If Positie >= 0 Then
               EersteQuery = Integer.Parse(.BatchBereik.Substring(0, Positie).Trim())
               LaatsteQuery = Integer.Parse(.BatchBereik.Substring(Positie + 1).Trim())
               QueryPadExtensie = Path.GetExtension(.BatchQueryPad)

               If EersteQuery.ToString() = .BatchBereik.Substring(0, Positie).Trim() AndAlso LaatsteQuery.ToString() = .BatchBereik.Substring(Positie + 1).Trim() AndAlso EersteQuery <= LaatsteQuery Then
                  For QueryIndex As Integer = EersteQuery To LaatsteQuery
                     QueryPad = $"{ .BatchQueryPad.Substring(0, .BatchQueryPad.Length - QueryPadExtensie.Length)}{QueryIndex}{QueryPadExtensie}".Trim(""""c)

                     If Query(QueryPad).Geopend Then
                        QueryParameters(Query().Code)

                        If .BatchInteractief AndAlso QueryIndex = EersteQuery Then
                           InteractieveBatchAfbreken = True
                           InterfaceVenster.Show()

                           ToonStatus($"{PROGRAMMA_VERSIE}{NewLine}")
                           If Not OPDRACHT_REGEL.Trim() = Nothing Then ToonStatus($"Opdrachtregel: {OPDRACHT_REGEL}{NewLine}")
                           If Not VerwerkSessieLijst() = Nothing Then ToonStatus($"Sessie lijst: {VerwerkSessieLijst()}{NewLine}")
                           ToonStatus($"Query: {QueryPad}{NewLine}")

                           Do While Application.OpenForms.Count > 0 AndAlso HuidigInterfaceVenster IsNot Nothing AndAlso HuidigInterfaceVenster.Enabled
                              Application.DoEvents()
                           Loop
                           If InteractieveBatchAfbreken Then Exit Sub

                           If HuidigInterfaceVenster IsNot Nothing Then HuidigInterfaceVenster.Cursor = Cursors.WaitCursor
                           INTERACTIEVE_BATCH_PARAMETERS.Clear()

                           QueryParameters.ForEach(Sub(QueryParameter As QueryParameterDefinitie) INTERACTIEVE_BATCH_PARAMETERS.Add(QueryParameter.Invoer))
                        Else
                           If QueryIndex = EersteQuery Then
                              If Not OPDRACHT_REGEL.Trim() = Nothing Then ToonStatus($"Opdrachtregel: {OPDRACHT_REGEL}{NewLine}")
                              If Not VerwerkSessieLijst() = Nothing Then ToonStatus($"Sessie lijst: {VerwerkSessieLijst()}{NewLine}")
                           End If

                           ToonStatus($"Query: {QueryPad}{NewLine}")

                           For ParameterIndex As Integer = 0 To QueryParameters().Count - 1
                              With QueryParameters()(ParameterIndex)
                                 QueryParameters(, ParameterIndex, $"{ .StandaardWaarde}{ .Masker.Substring(If(.StandaardWaarde.Length < .Masker.Length, .StandaardWaarde.Length + 1, .Masker.Length))}")
                                 If Not (.Commentaar = Nothing AndAlso .Masker = Nothing AndAlso .ParameterNaam = Nothing) Then ParameterSymboolFout(New StringBuilder("Genegeerde elementen in batch query gevonden."), ParameterIndex)
                              End With
                           Next ParameterIndex
                        End If

                        If OngeldigeParameterInvoer(FoutInformatie) Is Nothing Then
                           ToonStatus($"Bezig met het uitvoeren van de query...{NewLine}")
                           QueryResultaten(, ResultatenVerwijderen:=True)
                           VoerQueryUit(Query().Code)

                           If VERBINDING_GEOPEND(Verbinding()) Then
                              If Verbinding().Errors.Count = 0 Then
                                 ToonStatus($"{StatusNaQuery(ResultaatIndex:=0)}{NewLine}")
                                 If Not .ExportStandaardPad = Nothing Then
                                    ToonStatus($"Bezig met het exporteren van het queryresultaat...{NewLine}")
                                    ExportPad = Path.GetFullPath(VervangSymbolen(.ExportStandaardPad.Trim()).Trim(""""c))

                                    If Directory.Exists(Path.GetDirectoryName(ExportPad)) Then
                                       ExportPaden.Add(ExportPad)
                                       ExportUitgevoerd = ExporteerResultaat(ExportPad)
                                       If ExportUitgevoerd Then
                                          If File.Exists(ExportPad) AndAlso .ExportAutoOpenen Then
                                             ToonStatus($"De export wordt automatisch geopend...{NewLine}")
                                             Process.Start(New ProcessStartInfo With {.CreateNoWindow = False, .FileName = ExportPad, .ErrorDialog = True, .UseShellExecute = True, .WindowStyle = ProcessWindowStyle.Normal})
                                          End If

                                          ToonStatus($"Exporteren gereed.{NewLine}")
                                       Else
                                          ToonStatus($"Export afgebroken.{NewLine}")
                                       End If
                                    Else
                                       MessageBox.Show($"Ongeldig export pad.{NewLine}Huidig pad: ""{Directory.GetCurrentDirectory()}""", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                       ToonStatus($"Ongeldig export pad.{NewLine}")
                                    End If
                                 Else
                                    ToonStatus(FoutenLijstTekst(Verbinding().Errors))
                                 End If
                              End If

                              Verbinding(, , Reset:=True)
                           End If
                        Else
                           ParameterSymboolFout(New StringBuilder($"Ongeldige parameter invoer: {FoutInformatie}"))
                        End If
                     End If
                  Next QueryIndex

                  If (Not .ExportStandaardPad = Nothing) AndAlso ExportUitgevoerd Then
                     If Not (.ExportOntvanger = Nothing AndAlso .ExportCCOntvanger = Nothing) Then
                        ToonStatus($"Bezig met het maken van de e-mail met de export...{NewLine}")
                        EMail = New EMailClass
                        EMail.VoegQueryResultatenToe(ExportPaden)
                        EMail = Nothing
                     End If
                  End If
               Else
                  MessageBox.Show($"Ongeldige querybatchbereik: ""{ .BatchBereik}"".", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
               End If
            End If
         End With
      Catch
         HandelFoutAf()
      Finally
         If HuidigInterfaceVenster IsNot Nothing Then
            HuidigInterfaceVenster.Cursor = Cursors.Default
            HuidigInterfaceVenster.Close()
         End If
      End Try
   End Sub

   'Deze procedure voert een query uit op een database.
   Public Sub VoerQueryUit(QueryCode As String)
      Try
         Dim Commando As New Command
         Dim Resultaat As Recordset = Nothing
         Dim QueryPad As String = Query().Pad

         If VERBINDING_GEOPEND(Verbinding()) Then
            Commando.ActiveConnection = Verbinding()

            If Commando IsNot Nothing Then
               Do
                  Try
                     Commando.CommandText = VulParametersIn(QueryCode)
                     Commando.CommandText = VerwijderOpmaak(Commando.CommandText, SQL_COMMENTAAR_REGEL_BEGIN, SQL_COMMENTAAR_REGEL_EINDE, TEKENREEKS_TEKENS)
                     Commando.CommandText = VerwijderOpmaak(Commando.CommandText, SQL_COMMENTAAR_BLOK_BEGIN, SQL_COMMENTAAR_BLOK_EINDE, TEKENREEKS_TEKENS)
                     Commando.CommandTimeout = Instellingen().QueryTimeout
                     Commando.CommandType = CommandTypeEnum.adCmdText

                     Resultaat = Commando.Execute

                     Do While Resultaat IsNot Nothing AndAlso VERBINDING_GEOPEND(DirectCast(Resultaat.ActiveConnection, Connection))
                        QueryResultaten(Resultaat)
                        If Instellingen().QueryRecordSets Then Resultaat = Resultaat.NextRecordset Else Exit Do
                     Loop
                     Exit Do
                  Catch
                     If HandelFoutAf(ExtraInformatie:=$"Query: ""{QueryPad}""") = DialogResult.Ignore Then Exit Do
                  End Try
               Loop
            End If
         End If

         ToonStatus($"Uitgevoerde query: {NewLine}{Commando.CommandText}{NewLine}")

         If Resultaat IsNot Nothing AndAlso RECORDSET_GEOPEND(Resultaat) Then Resultaat.Close()

         Commando = Nothing
         Resultaat = Nothing
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure voert een sessie met de opgegeven parameters uit.
   Private Sub VoerSessieUit(SessieParameters As String)
      Try
         Static RecensteVerbindingsInformatie As String = Nothing

         With OpdrachtRegelParameters(SessieParameters)
            If .Verwerkt Then
               If .InstellingenPad = Nothing Then
                  Instellingen(Path.Combine(My.Application.Info.DirectoryPath, StandaardInstellingen().Bestand))
               Else
                  Instellingen(.InstellingenPad)
               End If

               With Instellingen()
                  If Not .VerbindingsInformatie.ToString() = RecensteVerbindingsInformatie Then
                     Verbinding(, VerbindingSluiten:=True)
                     If .VerbindingsInformatie.ToString().ToUpper().Contains(GEBRUIKER_VARIABEL) OrElse .VerbindingsInformatie.ToString().ToUpper().Contains(WACHTWOORD_VARIABEL) Then
                        InloggenVenster.ShowDialog()
                     Else
                        Verbinding(.VerbindingsInformatie.ToString())
                     End If
                  End If

                  RecensteVerbindingsInformatie = .VerbindingsInformatie.ToString()
               End With

               If VERBINDING_GEOPEND(Verbinding()) Then
                  If BATCH_MODUS_ACTIEF() Then
                     VoerBatchUit()
                  Else
                     InterfaceVenster.Show()

                     Do While Application.OpenForms.Count > 0
                        Application.DoEvents()
                     Loop
                  End If
               End If
            End If
         End With
      Catch
         HandelFoutAf()
      End Try
   End Sub

   'Deze procedure opent een dialoogvenster waarmee de gebruiker naar het pad voor het te exporteren queryresultaat kan bladeren.
   Public Function VraagExportPad(HuidigExportPad As String) As String
      Try
         Dim NieuwExportPad As String = HuidigExportPad
         Static ExportPadDialoog As New SaveFileDialog

         With ExportPadDialoog
            .CheckPathExists = True
            .Filter = "Tekstbestand (*.txt)|*.txt|Microsoft Excel bestand (*.xls)|*.xls|Microsoft Excel 2007 bestand (*.xlsx)|*.xlsx"
            .FilterIndex = 1
            .InitialDirectory = My.Application.Info.DirectoryPath
            .RestoreDirectory = True
            .Title = "Exporteer het queryresultaat naar:"
            If .ShowDialog() = DialogResult.OK Then NieuwExportPad = .FileName
         End With

         Return NieuwExportPad
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure opent een dialoogvenster waarmee de gebruiker naar een querybestand kan bladeren.
   Public Function VraagQueryPad() As String
      Try
         Dim NieuwQueryPad As String = Nothing
         Static QueryPadDialoog As New OpenFileDialog

         With QueryPadDialoog
            .CheckPathExists = True
            .Filter = "Tekstbestanden (*.txt)|*.txt"
            .FilterIndex = 1
            .InitialDirectory = My.Application.Info.DirectoryPath
            .RestoreDirectory = True
            .ShowReadOnly = False
            .Title = "Selecteer een query:"
            If .ShowDialog() = DialogResult.OK Then NieuwQueryPad = .FileName
         End With

         Return NieuwQueryPad
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure vraagt de gebruiker om de gegevens voor een verbinding met een database op te geven.
   Private Function VraagVerbindingsInformatie() As String
      Try
         Dim VerbindingsInformatie As String = ""

         Do While VerbindingsInformatie.Trim() = Nothing
            VerbindingsInformatie = ToonInvoerVenster("Informatie voor een verbinding met een database:", MeerdereRegels:=True)
            If VerbindingsInformatie Is Nothing Then
               VerbindingsInformatie = ""
               Exit Do
            ElseIf VerbindingsInformatie.Trim() = Nothing Then
               MessageBox.Show("Deze informatie is vereist.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
         Loop

         Return VerbindingsInformatie.Trim()
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function

   'Deze procedure vult de parameter invoer in de querycode in.
   Private Function VulParametersIn(QueryCode As String) As String
      Try
         Dim ParameterIndex As Integer = 0
         Dim Positie As New Integer
         Dim QueryMetParameters As New StringBuilder
         Dim QueryZonderParameters As String = Nothing

         If QueryParameters().Count = 0 Then
            QueryMetParameters.Append(QueryCode)
         Else
            QueryZonderParameters = QueryCode
            Do Until ParameterIndex > QueryParameters().Count - 1
               With QueryParameters()(ParameterIndex)
                  SESSIE_PARAMETERS.Add(.Invoer)
                  Positie = .Positie
                  QueryMetParameters.Append(QueryZonderParameters.Substring(0, Positie))
                  QueryMetParameters.Append(VervangSymbolen(.Invoer))
                  QueryZonderParameters = QueryZonderParameters.Substring(Positie + .Lengte)
               End With
               ParameterIndex += 1
            Loop
            QueryMetParameters.Append(QueryZonderParameters)
         End If

         Return QueryMetParameters.ToString()
      Catch
         HandelFoutAf()
      End Try

      Return Nothing
   End Function
End Module
