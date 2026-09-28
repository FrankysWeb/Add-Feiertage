<#
.SYNOPSIS
    Ergaenzt fehlende Feiertage (DE, AT) in einem eigenen Kalenderordner in Exchange-Postfaechern.
.DESCRIPTION
    Berechnet bundesweite und regionale Feiertage fuer einen Jahresbereich und legt sie per
    EWS (direkt per SOAP ueber HTTPS, ohne die EWS Managed API) in einem eigenen Kalenderordner
    "Feiertage" an, mit -MainCalendar stattdessen direkt im Hauptkalender. Vorhandene Eintraege
    werden erkannt und nicht doppelt angelegt.
.NOTES
    Braucht keine zusaetzliche DLL. Ein Dienstkonto mit der Rolle ApplicationImpersonation wird
    benoetigt.
#>

param(
    # E-Mail-Adresse(n) des Postfachs, in das die Feiertage eingetragen werden sollen
    # (die SMTP-Adresse des Postfachs, nicht der Anmeldename).
    [Parameter(Mandatory = $true)]
    [string[]]$Mailbox,

    # Jahre, fuer die Feiertage eingetragen werden. Default 2027..2031.
    [int[]]$Year = (2027..2031),

    # Land der Feiertage: DE (Default) oder AT.
    [ValidateSet('DE','AT')]
    [string]$Country = 'DE',

    # Nur relevant fuer -Country DE, z. B. 'BY','NW'. Ohne Angabe werden nur die
    # neun bundesweiten Feiertage eingetragen.
    [ValidateSet('BW','BY','BE','BB','HB','HH','HE','MV','NI','NW','RP','SL','SN','ST','SH','TH')]
    [string[]]$Bundesland = @(),

    # EWS-URL des Exchange Servers, in der Regel https://servername/EWS/Exchange.asmx.
    [Parameter(Mandatory = $true)]
    [string]$EwsUrl,

    # Name des Kalenderordners, in dem die Feiertage angelegt werden. Default "Feiertage".
    [string]$FolderName = "Feiertage",

    # Traegt die Feiertage direkt in den Hauptkalender ein statt in einen eigenen Ordner.
    # Nicht mit -FolderName kombinierbar.
    [switch]$MainCalendar,

    # Dienstkonto mit der Rolle ApplicationImpersonation. Der Benutzername meint hier den
    # Anmeldenamen (UPN, z. B. sa_feiertage@frankysweblab.de oder DOMAIN\sa_feiertage),
    # nicht die E-Mail-Adresse eines Postfachs. Ohne Angabe fragt PowerShell einmalig danach.
    [System.Management.Automation.PSCredential]$Credential = (Get-Credential -Message "Dienstkonto fuer die Feiertage-Impersonation (Benutzername als UPN, z. B. sa_feiertage@frankysweblab.de)"),

    # Zeigt nur, welche Feiertage angelegt wuerden, ohne etwas in Exchange zu aendern.
    [switch]$Preview,

    # Nur fuer Test-/Laborumgebungen mit selbstsigniertem oder nicht vertrauenswuerdigem Zertifikat.
    [switch]$IgnoreCertificateErrors,

    # Gibt die gesendeten und empfangenen SOAP-Envelopes auf der Konsole aus, nur zur Fehlersuche.
    [switch]$Trace
)

# Windows PowerShell aktiviert TLS 1.2 fuer ausgehende Verbindungen nicht automatisch,
# ohne das lehnen moderne Exchange Server die Verbindung waehrend des Handshakes ab.
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

if ($MainCalendar -and $PSBoundParameters.ContainsKey('FolderName')) {
    throw "-MainCalendar und -FolderName schliessen sich aus."
}

# Windows-Zeitzonen-ID fuer Deutschland und Oesterreich (MEZ/MESZ). Wird im TimeZoneContext
# jeder EWS-Anfrage mitgeschickt, damit ganztaegige Termine um lokale Mitternacht liegen.
$script:TimeZoneId = "W. Europe Standard Time"

if ($IgnoreCertificateErrors) {
    Write-Warning "Zertifikatspruefung ist deaktiviert, nur in Test-/Laborumgebungen verwenden."
    # Windows PowerShell (Desktop) nutzt fuer Invoke-WebRequest intern HttpWebRequest, dort
    # greift der ServicePointManager-Callback. PowerShell 7 (Core) nutzt HttpClient, dort
    # ignoriert Invoke-WebRequest den Callback und braucht stattdessen -SkipCertificateCheck.
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $script:SkipCertificateCheck = $true
    } else {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }
}

function Get-EasterSunday {
    param([int]$Year)
    # Gausssche Osterformel
    $a = $Year % 19
    $b = [Math]::Floor($Year / 100)
    $c = $Year % 100
    $d = [Math]::Floor($b / 4)
    $e = $b % 4
    $f = [Math]::Floor(($b + 8) / 25)
    $g = [Math]::Floor(($b - $f + 1) / 3)
    $h = (19 * $a + $b - $d - $g + 15) % 30
    $i = [Math]::Floor($c / 4)
    $k = $c % 4
    $l = (32 + 2 * $e + 2 * $i - $h - $k) % 7
    $m = [Math]::Floor(($a + 11 * $h + 22 * $l) / 451)
    $month = [Math]::Floor(($h + $l - 7 * $m + 114) / 31)
    $day = (($h + $l - 7 * $m + 114) % 31) + 1
    return Get-Date -Year $Year -Month $month -Day $day
}

function Get-BussUndBettag {
    # Mittwoch vor dem 23. November
    param([int]$Year)
    $nov23 = Get-Date -Year $Year -Month 11 -Day 23
    $daysBack = ([int]$nov23.DayOfWeek - 3 + 7) % 7
    if ($daysBack -eq 0) { $daysBack = 7 }
    return $nov23.AddDays(-$daysBack)
}

function Get-Holidays {
    param(
        [int]$Year,
        [string]$Country,
        [string[]]$Bundesland
    )

    $easter = Get-EasterSunday -Year $Year
    $holidays = New-Object System.Collections.Generic.List[PSCustomObject]

    switch ($Country) {
        'DE' {
            $holidays.Add([PSCustomObject]@{ Name = "Neujahr";                   Date = Get-Date -Year $Year -Month 1  -Day 1 })
            $holidays.Add([PSCustomObject]@{ Name = "Karfreitag";                Date = $easter.AddDays(-2) })
            $holidays.Add([PSCustomObject]@{ Name = "Ostermontag";               Date = $easter.AddDays(1) })
            $holidays.Add([PSCustomObject]@{ Name = "Tag der Arbeit";            Date = Get-Date -Year $Year -Month 5  -Day 1 })
            $holidays.Add([PSCustomObject]@{ Name = "Christi Himmelfahrt";       Date = $easter.AddDays(39) })
            $holidays.Add([PSCustomObject]@{ Name = "Pfingstmontag";             Date = $easter.AddDays(50) })
            $holidays.Add([PSCustomObject]@{ Name = "Tag der Deutschen Einheit"; Date = Get-Date -Year $Year -Month 10 -Day 3 })
            $holidays.Add([PSCustomObject]@{ Name = "1. Weihnachtstag";          Date = Get-Date -Year $Year -Month 12 -Day 25 })
            $holidays.Add([PSCustomObject]@{ Name = "2. Weihnachtstag";          Date = Get-Date -Year $Year -Month 12 -Day 26 })

            $regional = @(
                @{ Name = "Heilige Drei Könige";       Date = (Get-Date -Year $Year -Month 1 -Day 6);  Laender = @('BW','BY','ST') }
                @{ Name = "Ostersonntag";               Date = $easter;                                  Laender = @('BB') }
                @{ Name = "Pfingstsonntag";             Date = $easter.AddDays(49);                      Laender = @('BB') }
                @{ Name = "Internationaler Frauentag";  Date = (Get-Date -Year $Year -Month 3 -Day 8);  Laender = @('BE','MV') }
                @{ Name = "Fronleichnam";               Date = $easter.AddDays(60);                      Laender = @('BW','BY','HE','NW','RP','SL') }
                @{ Name = "Mariä Himmelfahrt";          Date = (Get-Date -Year $Year -Month 8 -Day 15); Laender = @('SL') }
                @{ Name = "Weltkindertag";              Date = (Get-Date -Year $Year -Month 9 -Day 20); Laender = @('TH') }
                @{ Name = "Reformationstag";            Date = (Get-Date -Year $Year -Month 10 -Day 31); Laender = @('BB','HB','HH','MV','NI','SH','SN','ST','TH') }
                @{ Name = "Allerheiligen";              Date = (Get-Date -Year $Year -Month 11 -Day 1);  Laender = @('BW','BY','NW','RP','SL') }
                @{ Name = "Buß- und Bettag";            Date = (Get-BussUndBettag -Year $Year);          Laender = @('SN') }
            )

            foreach ($h in $regional) {
                if ($Bundesland | Where-Object { $h.Laender -contains $_ }) {
                    $holidays.Add([PSCustomObject]@{ Name = $h.Name; Date = $h.Date })
                }
            }
        }
        'AT' {
            $holidays.Add([PSCustomObject]@{ Name = "Neujahr";              Date = Get-Date -Year $Year -Month 1  -Day 1 })
            $holidays.Add([PSCustomObject]@{ Name = "Heilige Drei Könige";  Date = Get-Date -Year $Year -Month 1  -Day 6 })
            $holidays.Add([PSCustomObject]@{ Name = "Ostermontag";         Date = $easter.AddDays(1) })
            $holidays.Add([PSCustomObject]@{ Name = "Staatsfeiertag";      Date = Get-Date -Year $Year -Month 5  -Day 1 })
            $holidays.Add([PSCustomObject]@{ Name = "Christi Himmelfahrt"; Date = $easter.AddDays(39) })
            $holidays.Add([PSCustomObject]@{ Name = "Pfingstmontag";       Date = $easter.AddDays(50) })
            $holidays.Add([PSCustomObject]@{ Name = "Fronleichnam";        Date = $easter.AddDays(60) })
            $holidays.Add([PSCustomObject]@{ Name = "Mariä Himmelfahrt";   Date = Get-Date -Year $Year -Month 8  -Day 15 })
            $holidays.Add([PSCustomObject]@{ Name = "Nationalfeiertag";    Date = Get-Date -Year $Year -Month 10 -Day 26 })
            $holidays.Add([PSCustomObject]@{ Name = "Allerheiligen";       Date = Get-Date -Year $Year -Month 11 -Day 1 })
            $holidays.Add([PSCustomObject]@{ Name = "Mariä Empfängnis";    Date = Get-Date -Year $Year -Month 12 -Day 8 })
            $holidays.Add([PSCustomObject]@{ Name = "Christtag";           Date = Get-Date -Year $Year -Month 12 -Day 25 })
            $holidays.Add([PSCustomObject]@{ Name = "Stefanitag";          Date = Get-Date -Year $Year -Month 12 -Day 26 })
        }
    }

    return $holidays
}

# --- EWS per rohem SOAP, ohne EWS Managed API -------------------------------------------------
#
# Die Antworten werden bewusst ueber local-name() statt ueber einen XmlNamespaceManager
# durchsucht. Das macht die Abfragen unabhaengig davon, welchen Namespace-Prefix eine
# bestimmte Antwort verwendet (normale Antworten nutzen z. B. "m"/"t", ein SOAP-Fault wie
# ErrorNonExistentMailbox dagegen "s"/"e"), und umgeht damit auch einen Ueberladungsfehler,
# den SelectSingleNode(xpath, XmlNamespaceManager) in manchen Umgebungen wirft.

function ConvertTo-XmlText {
    # Maskiert &, <, >, " und ', damit Werte wie Ordnername oder Postfachadresse das
    # SOAP-XML nicht zerbrechen koennen.
    param([string]$Value)
    return [System.Security.SecurityElement]::Escape($Value)
}

function Select-EwsNode {
    param([xml]$Xml, [string]$LocalName)
    return $Xml.SelectSingleNode("//*[local-name()='$LocalName']")
}

function Select-EwsNodes {
    param([xml]$Xml, [string]$LocalName)
    return $Xml.SelectNodes("//*[local-name()='$LocalName']")
}

function Invoke-EwsRequest {
    param(
        [string]$EwsUrl,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$SmtpAddress,
        [string]$InnerBody,
        [switch]$Trace
    )

    $envelope = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2016" />
    <t:ExchangeImpersonation>
      <t:ConnectingSID>
        <t:SmtpAddress>$(ConvertTo-XmlText $SmtpAddress)</t:SmtpAddress>
      </t:ConnectingSID>
    </t:ExchangeImpersonation>
    <t:TimeZoneContext>
      <t:TimeZoneDefinition Id="$script:TimeZoneId" />
    </t:TimeZoneContext>
  </soap:Header>
  <soap:Body>
$InnerBody
  </soap:Body>
</soap:Envelope>
"@

    if ($Trace) {
        Write-Host "--- Request ---" -ForegroundColor DarkGray
        Write-Host $envelope -ForegroundColor DarkGray
    }

    # Als UTF-8-Bytes senden, statt PowerShell die Kodierung selbst waehlen zu lassen,
    # das muss zur im XML-Prolog deklarierten Kodierung passen.
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($envelope)

    $requestParams = @{
        Uri             = $EwsUrl
        Method          = "Post"
        Credential      = $Credential
        Body            = $bytes
        ContentType     = "text/xml; charset=utf-8"
        UseBasicParsing = $true
    }
    if ($script:SkipCertificateCheck -and $PSVersionTable.PSEdition -eq 'Core') {
        $requestParams.SkipCertificateCheck = $true
    }

    try {
        $response = Invoke-WebRequest @requestParams
        $responseText = $response.Content
    } catch {
        # Auch bei HTTP 4xx/5xx liefert EWS meistens einen auswertbaren SOAP-Fault im Body mit.
        # $_.ErrorDetails.Message enthaelt den Antwortkoerper in Windows PowerShell und in
        # PowerShell 7 gleichermassen, unabhaengig vom konkreten Ausnahmetyp.
        if ($_.ErrorDetails.Message) {
            $responseText = $_.ErrorDetails.Message
        } else {
            throw
        }
    }

    if ($Trace) {
        Write-Host "--- Response ---" -ForegroundColor DarkGray
        Write-Host $responseText -ForegroundColor DarkGray
    }

    if (-not $responseText) {
        throw "Keine Antwort vom Server erhalten."
    }
    [xml]$xml = $responseText

    # Manche Fehler (z. B. ErrorNonExistentMailbox) kommen als echter SOAP-Fault zurueck,
    # nicht als normale Antwort mit ResponseClass="Error".
    $fault = Select-EwsNode -Xml $xml -LocalName "Fault"
    if ($fault) {
        $faultCode = Select-EwsNode -Xml $xml -LocalName "faultcode"
        $faultString = Select-EwsNode -Xml $xml -LocalName "faultstring"
        throw "EWS-Fehler $($faultCode.InnerText): $($faultString.InnerText)"
    }

    $responseCode = Select-EwsNode -Xml $xml -LocalName "ResponseCode"
    if ($responseCode -and $responseCode.InnerText -ne "NoError") {
        $messageText = Select-EwsNode -Xml $xml -LocalName "MessageText"
        if (-not $messageText) { $messageText = Select-EwsNode -Xml $xml -LocalName "Message" }
        throw "EWS-Fehler $($responseCode.InnerText): $($messageText.InnerText)"
    }

    return [PSCustomObject]@{ Xml = $xml }
}

function Get-OrCreateHolidayFolderId {
    param(
        [string]$EwsUrl,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$SmtpAddress,
        [string]$FolderName,
        [switch]$Preview,
        [switch]$Trace
    )

    $findBody = @"
    <m:FindFolder Traversal="Shallow">
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:FolderShape>
      <m:Restriction>
        <t:IsEqualTo>
          <t:FieldURI FieldURI="folder:DisplayName" />
          <t:FieldURIOrConstant>
            <t:Constant Value="$(ConvertTo-XmlText $FolderName)" />
          </t:FieldURIOrConstant>
        </t:IsEqualTo>
      </m:Restriction>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="calendar" />
      </m:ParentFolderIds>
    </m:FindFolder>
"@

    $result = Invoke-EwsRequest -EwsUrl $EwsUrl -Credential $Credential -SmtpAddress $SmtpAddress -InnerBody $findBody -Trace:$Trace
    $folderIdNode = Select-EwsNode -Xml $result.Xml -LocalName "FolderId"

    if ($folderIdNode) {
        return $folderIdNode.Id
    }

    if ($Preview) {
        Write-Host "  [Preview] Ordner '$FolderName' existiert noch nicht und wuerde angelegt" -ForegroundColor Yellow
        return $null
    }

    $createBody = @"
    <m:CreateFolder>
      <m:ParentFolderId>
        <t:DistinguishedFolderId Id="calendar" />
      </m:ParentFolderId>
      <m:Folders>
        <t:CalendarFolder>
          <t:DisplayName>$(ConvertTo-XmlText $FolderName)</t:DisplayName>
        </t:CalendarFolder>
      </m:Folders>
    </m:CreateFolder>
"@

    $result = Invoke-EwsRequest -EwsUrl $EwsUrl -Credential $Credential -SmtpAddress $SmtpAddress -InnerBody $createBody -Trace:$Trace
    $folderIdNode = Select-EwsNode -Xml $result.Xml -LocalName "FolderId"
    return $folderIdNode.Id
}

function Get-HolidayKey {
    param([string]$Subject, [datetime]$Date)
    return "$($Date.ToString('yyyy-MM-dd'))|$Subject"
}

function Get-ExistingHolidayKeys {
    # Sucht mit einer einzigen Anfrage pro Jahr nach Terminen, deren Betreff einem der
    # Feiertagsnamen entspricht, und liefert sie als Schluessel "Datum|Betreff". Exchange filtert
    # serverseitig, es kommen also nur passende Eintraege zurueck, auch im vollen Hauptkalender.
    # Anders als eine CalendarView loest diese Suche Serientermine nicht auf: Eine vom Nutzer
    # selbst angelegte jaehrliche Serie "Neujahr" wird nur mit ihrem ersten Termin gefunden.
    param(
        [string]$EwsUrl,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$SmtpAddress,
        [string]$FolderXml,
        [int]$Year,
        [string[]]$Subject,
        [switch]$Trace
    )

    # Einen Tag Puffer auf jeder Seite: Ganztaegige Termine liegen in UTC am Vortag ab 23:00
    # Uhr. Termine ausserhalb des Jahres stoeren nicht, weil der Schluessel das Datum enthaelt.
    $start = "$($Year - 1)-12-31T00:00:00"
    $end = "$($Year + 1)-01-02T00:00:00"

    $subjectFilter = ($Subject | ForEach-Object {
@"
            <t:IsEqualTo>
              <t:FieldURI FieldURI="item:Subject" />
              <t:FieldURIOrConstant>
                <t:Constant Value="$(ConvertTo-XmlText $_)" />
              </t:FieldURIOrConstant>
            </t:IsEqualTo>
"@
    }) -join "`n"

    $findBody = @"
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="calendar:Start" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="1000" Offset="0" BasePoint="Beginning" />
      <m:Restriction>
        <t:And>
          <t:IsGreaterThanOrEqualTo>
            <t:FieldURI FieldURI="calendar:Start" />
            <t:FieldURIOrConstant>
              <t:Constant Value="$start" />
            </t:FieldURIOrConstant>
          </t:IsGreaterThanOrEqualTo>
          <t:IsLessThan>
            <t:FieldURI FieldURI="calendar:Start" />
            <t:FieldURIOrConstant>
              <t:Constant Value="$end" />
            </t:FieldURIOrConstant>
          </t:IsLessThan>
          <t:Or>
$subjectFilter
          </t:Or>
        </t:And>
      </m:Restriction>
      <m:ParentFolderIds>
        $FolderXml
      </m:ParentFolderIds>
    </m:FindItem>
"@

    $result = Invoke-EwsRequest -EwsUrl $EwsUrl -Credential $Credential -SmtpAddress $SmtpAddress -InnerBody $findBody -Trace:$Trace

    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($script:TimeZoneId)
    $keys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in (Select-EwsNodes -Xml $result.Xml -LocalName "CalendarItem")) {
        $subjectNode = $item.SelectSingleNode("*[local-name()='Subject']")
        $startNode = $item.SelectSingleNode("*[local-name()='Start']")
        if (-not $subjectNode -or -not $startNode) { continue }

        # EWS liefert Start in der Regel in UTC (mit "Z"). Fuer den Vergleich zaehlt das lokale
        # Datum, sonst laege ein ganztaegiger Termin ab 00:00 MEZ (= 23:00 UTC) am Vortag.
        $startTime = [datetime]::Parse($startNode.InnerText, [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind)
        if ($startTime.Kind -ne [System.DateTimeKind]::Unspecified) {
            $startTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($startTime.ToUniversalTime(), $tz)
        }
        [void]$keys.Add((Get-HolidayKey -Subject $subjectNode.InnerText -Date $startTime))
    }
    return ,$keys
}

function Add-HolidayAppointment {
    param(
        [string]$EwsUrl,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$SmtpAddress,
        [string]$FolderXml,
        [string]$Subject,
        [datetime]$Date,
        [switch]$Trace
    )

    # Ganztaegige Termine muessen um Mitternacht in der Zeitzone des Termins beginnen und
    # enden. Ohne "Z" gelten die Zeiten in der Zeitzone aus dem TimeZoneContext-Header, sonst
    # wuerde Exchange sie als UTC werten und Outlook den Termin ueber zwei Tage anzeigen.
    $start = $Date.Date.ToString("yyyy-MM-ddT00:00:00")
    $end = $Date.Date.AddDays(1).ToString("yyyy-MM-ddT00:00:00")

    $createBody = @"
    <m:CreateItem SendMeetingInvitations="SendToNone">
      <m:SavedItemFolderId>
        $FolderXml
      </m:SavedItemFolderId>
      <m:Items>
        <t:CalendarItem>
          <t:Subject>$(ConvertTo-XmlText $Subject)</t:Subject>
          <t:Start>$start</t:Start>
          <t:End>$end</t:End>
          <t:IsAllDayEvent>true</t:IsAllDayEvent>
          <t:LegacyFreeBusyStatus>Free</t:LegacyFreeBusyStatus>
          <t:ReminderIsSet>false</t:ReminderIsSet>
        </t:CalendarItem>
      </m:Items>
    </m:CreateItem>
"@

    Invoke-EwsRequest -EwsUrl $EwsUrl -Credential $Credential -SmtpAddress $SmtpAddress -InnerBody $createBody -Trace:$Trace | Out-Null
}

# --- Hauptlauf ---------------------------------------------------------------------------------

$failed = New-Object System.Collections.Generic.List[string]

foreach ($smtp in $Mailbox) {
    Write-Host "Verarbeite $smtp" -ForegroundColor Cyan

    # Ein Fehler bei einem Postfach (z. B. fehlende Berechtigung) soll die uebrigen Postfaecher
    # nicht blockieren, deshalb wird er hier abgefangen und am Ende zusammengefasst.
    try {
        # Zielordner als XML-Element fuer die EWS-Anfragen: der Hauptkalender ueber seinen
        # festen Namen, der Feiertagsordner ueber seine ID. Im Preview kann der Ordner noch
        # fehlen, dann bleibt $folderXml leer.
        if ($MainCalendar) {
            $folderXml = '<t:DistinguishedFolderId Id="calendar" />'
        } else {
            $folderId = Get-OrCreateHolidayFolderId -EwsUrl $EwsUrl -Credential $Credential -SmtpAddress $smtp `
                -FolderName $FolderName -Preview:$Preview -Trace:$Trace
            $folderXml = if ($folderId) { "<t:FolderId Id=`"$(ConvertTo-XmlText $folderId)`" />" } else { $null }
        }

        foreach ($y in $Year) {
            $holidays = Get-Holidays -Year $y -Country $Country -Bundesland $Bundesland

            $existing = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            if ($folderXml) {
                $existing = Get-ExistingHolidayKeys -EwsUrl $EwsUrl -Credential $Credential -SmtpAddress $smtp `
                    -FolderXml $folderXml -Year $y -Subject ($holidays | ForEach-Object { $_.Name }) -Trace:$Trace
            }

            foreach ($holiday in $holidays) {
                $exists = $existing.Contains((Get-HolidayKey -Subject $holiday.Name -Date $holiday.Date))

                if (-not $exists) {
                    if ($Preview) {
                        Write-Host "  [Preview] wuerde anlegen: $($holiday.Name) ($($holiday.Date.ToString('dd.MM.yyyy')))" -ForegroundColor Yellow
                    } else {
                        Add-HolidayAppointment -EwsUrl $EwsUrl -Credential $Credential -SmtpAddress $smtp `
                            -FolderXml $folderXml -Subject $holiday.Name -Date $holiday.Date -Trace:$Trace
                        Write-Host "  + $($holiday.Name) ($($holiday.Date.ToString('dd.MM.yyyy'))) angelegt"
                    }
                } else {
                    Write-Host "  = $($holiday.Name) ($($holiday.Date.ToString('dd.MM.yyyy'))) bereits vorhanden, uebersprungen" -ForegroundColor DarkGray
                }
            }
        }
    } catch {
        Write-Warning "  $smtp uebersprungen: $($_.Exception.Message)"
        $failed.Add($smtp)
    }
}

if ($failed.Count -gt 0) {
    Write-Warning "$($failed.Count) von $($Mailbox.Count) Postfaechern mit Fehlern: $($failed -join ', ')"
}
