# GUIA UD3.AO.FC

**Repositori:** https://github.com/igarcia21404-dev/UD3.AO.FC

## Exercici 2 — Personalització de barres d'eines (FER A LES MVs)

### LibreOffice Calc (Linux)
1. Obre LibreOffice Calc
2. Fes captura d'estat inicial
3. Visualitza → Barres d'eines → Personalitza...
4. Afegeix: Desa, Desfés, Refés, Moneda, Percentatge, Ordena, Filtre, Suma...
5. Fes captura després dels canvis

### WPS Spreadsheets (Linux)
1. Obre WPS Spreadsheets
2. Captura inicial
3. Clic dret a la barra → Personalitzar
4. Afegeix les mateixes opcions
5. Captura final

Repeteix el mateix a Windows amb dos programes (OpenOffice/LibreOffice + Excel o WPS).

**Total:** 2 minitutorials + 4 imatges a Linux i el mateix a Windows.

## Exercici 3 — Plantilles LibreOffice (.ods)
- **Checklist.ods**: tasques + Sí/No + recompte amb COUNTA i COUNTIF
- **Factura.ods**: imports, subtotal, IVA 21% i total amb fórmules

## Exercici 4 — Plantilles WPS (.xlsx)
- **Quadern_de_notes.xlsx**: activitats + examen + mitjana
- **Llista_notes_finals.xlsx**: 3 avaluacions + nota final

## Exercici 5 i 6 — Importació/Exportació
Les versions ja estan convertides a les carpetes de Windows.

## Exercici 7 — Macros
Document: `Ubuntu/LibreOffice/Tasca3.A.Exercici7.Macros.ods`

### Fórmules:
- Comparació: `=IF(A5>10;"Major que 10";IF(A5=10;"Igual a 10";"Menor que 10"))`
- Parell/Senar: `=IF(MOD(A11;2)=0;"Parell";"Senar")`
- Bases: DEC2BIN, DEC2OCT, DEC2HEX, BIN2DEC, OCT2DEC, HEX2DEC

### Codi Basic (per crear macros):
```
Sub CompararAmb10
    Dim valor As Double
    Dim resultat As String
    valor = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("A5").getValue()
    If valor > 10 Then
        resultat = "Major que 10"
    ElseIf valor = 10 Then
        resultat = "Igual a 10"
    Else
        resultat = "Menor que 10"
    End If
    ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("B5").setString(resultat)
End Sub

Sub ParellOSenar
    Dim valor As Long
    Dim resultat As String
    valor = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("A11").getValue()
    If (valor Mod 2) = 0 Then
        resultat = "Parell"
    Else
        resultat = "Senar"
    End If
    ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("B11").setString(resultat)
End Sub
```

Per crear les macros: Eines → Macros → Organitza macros → LibreOffice Basic → enganxa el codi.
