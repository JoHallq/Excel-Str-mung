Attribute VB_Name = "Wirbelstrasse"
' Steuerung fuer Wirbelstrasse.xlsm
' Ein Zeitschritt = die Rechenbloecke werden in fester Reihenfolge neu berechnet:
'   W0 (alt) -> U -> V -> W1 (neu) -> PSI -> Anzeige
' Range.Calculate rechnet jeden Block zeilenweise von oben nach unten, unabhaengig davon,
' wie Excel die Zirkelbezuege sonst sortieren wuerde. Dadurch ist die Rechnung reproduzierbar.
Option Explicit

Private laeuft As Boolean

Private Function N(ByVal name As String) As Range
    Set N = ThisWorkbook.Names(name).RefersToRange
End Function

Private Sub Vorbereiten()
    Application.Calculation = xlCalculationManual
    Application.Iteration = True
    Application.MaxIterations = 1
End Sub

Private Sub EinSchritt()
    N("blk_W0").Calculate
    N("blk_U").Calculate
    N("blk_V").Calculate
    N("blk_W1").Calculate
    N("blk_PSI").Calculate
    N("Schritt").Calculate
End Sub

Private Sub Anzeigen()
    N("visc").Calculate   ' Re oder D waehrend des Laufs geaendert? -> neue Viskositaet
    N("blk_Anzeige").Calculate
    N("Schrittanzeige").Calculate
End Sub

Public Sub Starten()
    Dim k As Long, proBild As Long
    Vorbereiten
    If laeuft Then Exit Sub
    If N("Reset").Value <> 0 Then N("Reset").Value = 0
    laeuft = True
    Do While laeuft
        proBild = N("Schritte_pro_Bild").Value
        If proBild < 1 Then proBild = 1
        Application.ScreenUpdating = False
        For k = 1 To proBild
            EinSchritt
        Next k
        Anzeigen
        Application.ScreenUpdating = True
        DoEvents
    Loop
End Sub

Public Sub Anhalten()
    laeuft = False
End Sub

Public Sub Einzelschritt()
    Vorbereiten
    If N("Reset").Value <> 0 Then N("Reset").Value = 0
    EinSchritt
    Anzeigen
End Sub

Public Sub Zuruecksetzen()
    laeuft = False
    Vorbereiten
    N("Reset").Value = 1
    ' Parameter koennten sich geaendert haben: alles abhaengige neu rechnen
    N("visc").Calculate
    N("Rzyl").Calculate
    N("psiBody").Calculate
    N("blk_MASK").Calculate
    N("blk_NF").Calculate
    EinSchritt
    EinSchritt
    Anzeigen
    N("Reset").Value = 0
End Sub
