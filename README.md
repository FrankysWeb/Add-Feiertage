# Add-Feiertage

Trägt fehlende Feiertage (Deutschland, Österreich) per EWS in einen eigenen Kalenderordner
in Exchange-Postfächern ein. Hintergrund: Microsoft pflegt die Feiertagsdaten in Outlook
(`OUTLOOK.HOL`) nur bis Ende 2026, ab 2027 fehlen bewegliche Feiertage wie Ostern komplett.

## Was das Skript macht

- Berechnet die Feiertage selbst (gaußsche Osterformel für die beweglichen Termine), es lädt
  sie nicht aus einer Datei.
- Legt dafür einen eigenen Kalenderordner `Feiertage` an, genau wie Outlook das selbst tut,
  wenn man Feiertage manuell über die Optionen hinzufügt. So landen sie nicht zwischen den
  echten Terminen im Hauptkalender. Mit `-MainCalendar` landen sie stattdessen direkt im
  Hauptkalender.
- Erkennt bereits vorhandene Einträge (Betreff + Datum) und legt nichts doppelt an. Selbst
  angelegte Serientermine (z. B. „Neujahr“ als jährliche Serie) werden dabei nur mit ihrem
  ersten Termin erkannt.
- Deckt zwei Länder ab und unterscheidet bei Deutschland zwischen bundesweiten und regionalen
  Feiertagen (siehe unten).

## Voraussetzungen

- Exchange Server on-premises (2016, 2019 oder Subscription Edition). Exchange Online wird nicht
  unterstützt: Dort gibt es die Rolle `ApplicationImpersonation` nicht mehr, und Microsoft
  schaltet EWS in Exchange Online ab.
- Ein Dienstkonto mit der Rolle `ApplicationImpersonation`:

  ```powershell
  New-ManagementRoleAssignment -Name "Feiertage-Impersonation" -Role "ApplicationImpersonation" -User "DOMAIN\sa_feiertage"
  ```

- Die EWS-URL des Exchange Servers, in der Regel `https://servername/EWS/Exchange.asmx`.
- Windows PowerShell 5.1 oder PowerShell 7, beides wird unterstützt (siehe `-IgnoreCertificateErrors`
  unten für den Unterschied).

## Parameter

| Parameter | Pflicht | Beschreibung |
|---|---|---|
| `-Mailbox` | Ja | E-Mail-Adresse(n) des Postfachs, in das die Feiertage eingetragen werden. Mehrere Adressen als Array. |
| `-EwsUrl` | Ja | EWS-URL des Exchange Servers. |
| `-Credential` | Nein | Dienstkonto mit `ApplicationImpersonation`. Der Benutzername meint den Anmeldenamen (UPN, z. B. `sa_feiertage@domain.de`, oder `DOMAIN\sa_feiertage`), nicht die E-Mail-Adresse eines Postfachs. Ohne Angabe fragt PowerShell einmalig danach. |
| `-Year` | Nein | Jahre, für die Feiertage eingetragen werden. Default `2027..2031`. |
| `-Country` | Nein | `DE` (Default) oder `AT`. |
| `-Bundesland` | Nein | Nur für `-Country DE`. Offizielle zweistellige ISO-3166-2:DE-Kürzel (`BW`, `BY`, `BE`, `BB`, `HB`, `HH`, `HE`, `MV`, `NI`, `NW`, `RP`, `SL`, `SN`, `ST`, `SH`, `TH`), z. B. `"NW","BY"`. Andere Werte werden abgelehnt. Ohne Angabe werden nur die neun bundesweiten Feiertage eingetragen. |
| `-FolderName` | Nein | Name des Kalenderordners. Default `Feiertage`. |
| `-MainCalendar` | Nein | Trägt die Feiertage direkt in den Hauptkalender ein, es wird kein eigener Ordner angelegt. Nicht mit `-FolderName` kombinierbar. |
| `-Preview` | Nein | Zeigt nur, welche Feiertage angelegt würden, ohne etwas in Exchange zu ändern (auch der Ordner wird dabei nicht angelegt). |
| `-IgnoreCertificateErrors` | Nein | Nur für Test-/Laborumgebungen mit selbstsigniertem oder sonst nicht vertrauenswürdigem Zertifikat. Unter Windows PowerShell 5.1 über den `ServicePointManager`, unter PowerShell 7 über `-SkipCertificateCheck`, da PowerShell 7 intern `HttpClient` statt `HttpWebRequest` nutzt und den `ServicePointManager`-Callback nicht beachtet. |
| `-Trace` | Nein | Gibt jeden gesendeten und empfangenen SOAP-Envelope auf der Konsole aus, nur zur Fehlersuche gedacht. |

## Beispiele

Nur die neun bundesweiten deutschen Feiertage, zuerst als Preview:

```powershell
.\Add-Feiertage.ps1 -Mailbox "max.mustermann@example.com" -EwsUrl "https://mail.example.com/EWS/Exchange.asmx" -Preview
```

Inklusive der Feiertage für Nordrhein-Westfalen, für echt:

```powershell
.\Add-Feiertage.ps1 -Mailbox "max.mustermann@example.com" -EwsUrl "https://mail.example.com/EWS/Exchange.asmx" -Bundesland "NW"
```

Österreich, nur für 2027:

```powershell
.\Add-Feiertage.ps1 -Mailbox "max.mustermann@example.at" -EwsUrl "https://mail.example.at/EWS/Exchange.asmx" -Country AT -Year 2027
```

Direkt in den Hauptkalender statt in einen eigenen Ordner:

```powershell
.\Add-Feiertage.ps1 -Mailbox "max.mustermann@example.com" -EwsUrl "https://mail.example.com/EWS/Exchange.asmx" -MainCalendar
```

Mehrere Bundesländer und ein eigener Ordnername:

```powershell
.\Add-Feiertage.ps1 -Mailbox "max.mustermann@example.com" -EwsUrl "https://mail.example.com/EWS/Exchange.asmx" -Bundesland "BY","BW" -FolderName "Feiertage Süd"
```

Dienstkonto vorab abfragen, z. B. für mehrere Läufe hintereinander:

```powershell
$cred = Get-Credential "sa_feiertage@example.com"
.\Add-Feiertage.ps1 -Mailbox "max.mustermann@example.com" -EwsUrl "https://mail.example.com/EWS/Exchange.asmx" -Credential $cred
```

Fehlersuche in einer Testumgebung mit selbstsigniertem Zertifikat, ohne etwas zu ändern:

```powershell
.\Add-Feiertage.ps1 -Mailbox "max.mustermann@example.com" -EwsUrl "https://mail.example.com/EWS/Exchange.asmx" -Preview -Trace -IgnoreCertificateErrors
```

Mehrere Postfächer aus Exchange auslesen:

```powershell
$adressen = (Get-Mailbox -ResultSize Unlimited).PrimarySmtpAddress
.\Add-Feiertage.ps1 -Mailbox $adressen -EwsUrl "https://mail.example.com/EWS/Exchange.asmx"
```

## Einschränkungen

- Für den einmaligen Nachtrag gedacht, nicht für den Dauerbetrieb. Bei sehr vielen Postfächern
  kann Exchange die EWS-Anfragen drosseln.
- Regionale Feiertage einzelner Gemeinden (z. B. Fronleichnam nur in bestimmten bayerischen
  Gemeinden) sind nicht enthalten, nur die auf Bundeslandebene geltenden.
- Die Schweiz ist nicht enthalten: Gesetzlich bundesweit vorgeschrieben ist dort nur der
  Bundesfeiertag, alles andere ist kantonal geregelt.

## Autor

Frank Zöchling, [www.frankysweb.de](https://www.frankysweb.de)

## Lizenz

MIT, siehe [LICENSE](LICENSE).
