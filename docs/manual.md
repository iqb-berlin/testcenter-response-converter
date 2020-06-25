# IQB-Testcenter – AntwortKonverter
Über den Admin-Bereich des Testcenters lassen sich vor allem zwei Dateiarten 
herunterladen: Responses und Logs. Diese Rohdaten sind schlecht auswertbar. 
Die Desktop-Anwendung itc-AntwortKonverter transformiert diese Daten. 
Dieser Text beschreibt die Struktur dieser transformierten Daten.

Der Anwendung wird zunächst ein Verzeichnis mitgeteilt, in dem die 
Response- und Log-Daten im CSV-Format liegen. Bei kleineren Erhebungen 
sind dies zwei Dateien, bei größeren Studien könnte eine Aufteilung in 
viele Dateien erforderlich sein. Zusätzlich kann der Anwendung eine txt-Datei 
übergeben werden, die für jedes Booklet die Größe angibt (pro Zeile jeweils 
Booklet-ID, Leerzeichen, Größe in Bytes). Dies kann zur Schätzung der 
Geschwindigkeit der Netzwerkverbindung des Testcomputers verwendet werden.

Als Ausgabe wird eine Xlsx-Datei erzeugt. Diese enthält in drei Tabellen die 
gewünschten Daten. Nachfolgend wird die Bedeutung der Spalten jeder dieser 
Tabellen beschrieben:

## Tabelle Responses

| Spaltenbezeichnung | Bedeutung |
| :------------- | :---------- |
|ID|Kombination aus anderen (nachfolgenden) Informationen. Diese ID wird benötigt, um eine Zeile eindeutig zu identifizieren. Es handelt sich um eine Testsitzung, also eine Testperson beantwortet ein konkretes Booklet. Diese Kombination ist nötig, weil eine Testperson mehrere Booklets haben könnte und in einem Booklet theoretisch dieselbe Unit enthalten sein kann (z. B. Motivationsabfrage). Es muss eindeutig sein, in welchem Booklet diese Unit platziert war.<br> Diese ID wird auch in den anderen zwei Tabellen verwendet, so dass hierüber eine Zusammenführung der Informationen erfolgen kann.|
|Group|Gruppe, in der das Anmelde-Login platziert war. Dies ist normalerweise nur ein Ordnungsmerkmal für das Monitoring der Durchführung.|
|Login+Code|Entsprechend der Anmeldung der Testperson|
| Booklet | ID des Booklets |
| Variablen nach dem Schema<br>`<Unit-ID>##<innere ID>`<br>z. B. `EL105R##canvasElement10` | Der Player des Testcenters speichert im bisherigen Modell die Antwortdaten als Paarung ID->Wert ab, wobei nicht definiert ist, was ID kennzeichnet (Item, Aspekt eines Items, Eingabeelement des Formulars usw.; hier mal als innere ID bezeichnet). Sicher ist nur, dass diese ID innerhalb der Unit eindeutig ist, und da die Unit-ID eindeutig für das Testheft ist, erlaubt die Kombination Unit-ID mit dieser inneren ID eine eindeutige Zuordnung des Antwort-Wertes zu einer Testperson in einem Booklet, wodurch sich die übliche zweidimensionale Struktur der Antwortdaten ergibt.<br> Es werden nur Units berücksichtigt, die tatsächlich Antwortdaten produziert haben. Reine Textseiten, die z. B. nur Instruktionen enthalten, werden nicht in die Tabelle aufgenommen.<br>Die Variablenspalten werden alphabetisch sortiert ausgegeben.<br>Sollte eine Unit mehrfach in einem Test vorkommen, fügt das System ab dem zweiten Vorkommen der Unit automatisch ein Suffix hinzu:<br>`<Unit-ID>%<n>`<br>n steht hier für die fortlaufende Nummerierung, beginnend mit 1 bei dem zweiten Vorkommen der Unit|

Die Zeilen dieser Tabelle sind nach ID sortiert. Sollte eine Testperson den Test nur gestartet, aber keine Antwortdaten abgeschickt haben, erscheint sie nicht in der Liste.

## Tabelle TimeOnPage

Für die weitere Beurteilung der Antworten schickt das IQB-Testcenter eine größere 
Menge zeitpunktbezogener Daten, sog. Log-Daten. Hierbei wird stets ein Zeitstempel 
mitgeliefert (Datum und Uhrzeit auf dem Computer der Testperson) sowie Art des 
Ereignisses und ggf. weitere Informationen. Aus dieser Folge von Ereignissen lässt 
sich die Navigation zwischen Units und Seiten und somit die Zeit ermitteln, die eine 
Testperson während des Tests auf einer bestimmten Seite verbracht hat.

Eine Unit besteht aus mindestens einer Seite, auf denen die visuellen Elemente 
sowie die Antwort-Elemente platziert sind. Sollte es bei mehreren Seiten eine Seite 
geben, die ständig angezeigt wird, wird diese nicht extra aufgeführt, sondern deren 
Zeit müsste als Summe aller anderen Seitenzeiten ermittelt werden. Die Verteilung 
der Antwort-Elemente (z. B. Items) auf die Seiten wird nicht erfasst.

Die erste Spalte enthält die ID (s. Tabelle Responses), dann folgen für jede Seite 
zwei Spalten:
* `<Unit-ID>##<Seiten-ID>##topTotal`<br>
Zeit, die die Seite angezeigt wurde. Diese Verweildauer ist eine Summe aller 
Besuche in Millisekunden. Konnte für eine Seite kein Endzeitpunkt festgestellt 
werden (z. B. wegen eines Neuladen des Tests), wird hier eine Null eingetragen.
* `<Unit-ID>##<Seiten-ID>##topCount`<br>
Anzahl aller Besuche dieser Seite.

Die Unitseiten-Spalten sind alphabetisch sortiert. Es werden nur Units beachtet, 
die in der Tabelle Responses aufgeführt sind, d. h. nur Units mit Antwortdaten.

Die Verweilzeiten werden nach folgender Methode ermittelt: Startzeitpunkt 
wird durch PAGENAVIGATIONCOMPLETE festgestellt. Dieses Ereignis wird durch 
den Player ausgelöst, wenn alle Elemente der Seite vollständig dargestellt 
sind und ggf. eine vorherige Beantwortung rekonstruiert wurde. Ein Ende der 
Anzeige wird angenommen, wenn ein Ereignis auftritt, das nicht in der folgenden 
Liste aufgeführt ist: RESPONSESCOMPLETE, PRESENTATIONCOMPLETE und UNITTRYLEAVE.

## Tabelle TechData

In dieser Tabelle sind weitere eventuell interessante Daten einer Testsitzung 
aufgelistet:

| Spaltenbezeichnung | Bedeutung |
| :------------- | :---------- |
|ID|ID der Testsitzung wie in den anderen Tabellen|
|Start at|Beginn des ersten Ladens der Testinhalte nach Auswahl des Booklets durch die Testperson. Es handelt sich um eine in JavaScript über Date.now() ermittelte Anzahl der Millisekunden, die seit dem 01.01.1970 00:00:00 UTC vergangen sind. Für Excel muss man den Wert umrechnen: `=<ts>/(1000*60*60*24) + 25569` und dann als Datum+Zeit formatieren: TT.MM.JJJJ h:mm:ss|
|loadcomplete after|Dauer des Ladevorganges in Millisekunden|
|loadspeed|Ladegeschwindigkeit als Quotient aus Bookletgröße (aus der zusätzlich zugewiesenen txt-Datei) und Ladedauer. Wenn die Bookletgröße in Bytes und die Dauer in Millisekunden angegeben werden (wie hier aktuell im Testcenter), dann ist die Einheit des Wertes kBytes/sec|
|firstUnitEnter after|Zeit zwischen Start des Ladens der Testinhalte und Eintritt in die erste Unit. Achtung: In der ersten Version hat das IQB-Testcenter auf den Abschluss des gesamten Ladevorganges gewartet, ehe die erste Unit angezeigt wurde. Nach der ersten Testwoche wurde das Ladeverhalten verändert: Sobald die erste Unit geladen ist, wird sie auch angezeigt. Im Hintergrund werden die anderen Units geladen. Sollte beim Navigieren die nächste Unit noch nicht geladen worden sein (bzw. bei einem Block mit Maximalzeit der gesamte Block), wartet das System. Die Unterscheidung zwischen der ersten Version und den nachfolgenden Testungen ist durch Vergleich „loadcomplete after“ < „firstUnitEnter after“ möglich.|
|os|Betriebssystem (operating system)|
|browser|Name und Version|
|screen|Breite x Höhe in Pixels|
|pages visited ratio|Anteil der in dieser Testsitzung besuchten Seiten an der Gesamtzahl der für dieses Booklet bekannten Seiten. Bekannt sind Seiten durch den Besuch irgendeiner Testperson, d. h. zu einem frühen Zeitpunkt der Testung mögen nicht alle Seiten bekannt sein. Es werden nur Seiten entsprechend der Tabelle TimeOnPages berücksichtigt, d. h. Seiten von Units, zu denen Antworten gespeichert wurden.|
|units fully responded ratio|Anteil der vollständig beantworteten Units an allen bekannten Units des Booklets. Es wird zur Ermittlung das Ereignis RESPONSESCOMPLETE: „all“ verwendet, das der Player schickt. Es werden nur Units berücksichtigt, zu denen Antworten gespeichert wurden. Es kann sein, dass nachträglich Antworten wieder entfernt wurden, was nicht berücksichtigt wird.|
