# AfsprakenProgramma
Excel code base for a Dutch Excel application to generate medical orders

## Description in Dutch
Excel applicatie voor het genereren van voorschriften. Deze applicatie is specifiek bedoeld voor gebruik in combinatie met MetaVision 5.x versies. 

De applicatie is gemaakt voor monitoren met een resolutie van 1900 x 1200.

## Kenmerken

1. Patient gegevens:
  * Koppelling met MetaVision mogelijk met synchronisatie van patient gegevens.
  * Validatie van geboortedatum, leeftijd, gewicht, lengte, zwangerschapsduur en geboortegewicht.
2. Totalen:
  * Berekening van vocht en ingredienten in totalen.
  * Koppeling met MetaVision ten aanzien van relevante lab uitslagen.
3. Continue Infusen:
  * Snelkeuze lijst voor continue medicatie.
  * Berekeningen van doseringen.
  * Berekening van concentraties.
  * Automatische selectie van standaard concentraties.
  * Weergave van onder en bovengrenzen voor doseringen.
  * Controle op oplosmiddel voor specifieke medicamenten.
  * Doorrekening totalen.
4. Discontinue medicatie:
  * Gebaseeerd op de G-Standaard.
  * Indicatie gericht voorschrijven volgens Kinderformularium indicaties.
  * Doorlink moglijkheid naar het Kinderformularium.
  * Gebruik van standaard frequenties.
  * Mogelijkheid tot oplossing en toediening tijd.
  * Berekening van dagdosis en concentratie indien in oplossing.
  * Doorrekening in totalen van de oplossing.
  * Mogelijkheid tot medicatie bewaking via de G-Standaard en het Kinderformularium.
5. Lijnen en Pacemaker:
  * Snelkeuze lijst voor intravasculaire lijnen en flushes over deze lijnen.
  * Doorrekening van de flushes in de totalen.
  * Opgeven van pacemaker instellingen.
6. Voeding en TPN:
  * Snelkeuze lijst voor voedingen en toevoegingen.
  * Configuratie van speciale voeding met ingredienten.
  * Doorrekening van ingredienten en vocht in totalen.
  * Instellen van TPN.
  * Automatische selectie van juiste eiwit samenstelling.
  * Automatische instelling van TPN voor opbouw van TPN.
  * Doorrekening van ingredienten en vocht in totalen.
7. Lab aanvragen:
  * Snelkeuze lijsten voor labaanvragen op standaard tijden.
8. Afspraken en controles:
  * Snelkeuze lijsten en tekst invoer.
9. Infuusbrief voor de Neonatologie:
  * Geintegreerd overzicht van voeding, continue medicatie, en IV toedieningen.
  * Automatische berekening van oplopende vocht toediening naar aanleding van leeftijd en zwangerschapsduur.
  * Correctie van vochttoediening voor fototherapie.
  * Snelkeuze lijst voor voedingen en toevoegingen.
  * Doorrekening van vocht en ingredienten in totalen.
  * Berekening van totalen enteraal vs parenteraal.
10. Generatie van een infuusbrief voor elektroliet oplossingen en TPN.
11. Generatie van werkbrieven met infusen voor Neonatologie.
12. Generatie van bereidingvoorschriften voor klaarmaken van medicatie door de apotheek.
13. Generatie van een lijst met berekende acute medicatie en interventies volgens APLS.
14. Uitprint mogelijkheid van ingevoerde afspraken en medicatie.
15. Koppeling mogelijk voor verwerking van afspraken en medicatie in MetaVision.

## Installatie

Voor de koppeling met MetaVision is een bestand `secret` nodig met daarin:</br>
login</br>
wachtwoord</br>

Dit bestand wordt gebruikt om toegang te krijgen tot de MetaVision database. Daarnaast verwacht de applicatie dat MetaVision 2 registry keys wegschrijft om toegang te krijgen tot de database.

## Beheer

