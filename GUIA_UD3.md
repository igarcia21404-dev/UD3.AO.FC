# UD3.AO.FC — Guia de les plantilles

S'han afegit quatre plantilles CSV per obrir amb Calc o WPS:
- Ubuntu/LibreOffice/Checklist.csv: tasques Sí/No i recompte.
- Ubuntu/LibreOffice/Factura.csv: imports, subtotal, IVA 21% i total.
- Ubuntu/WPS Office/Quadern_de_notes.csv: mitjana d'activitats i examen.
- Ubuntu/WPS Office/Llista_notes_finals.csv: mitjana de tres avaluacions.

Obre cada CSV i selecciona la coma com a separador. Comprova les fórmules i adapta els separadors d'arguments a la configuració regional. Després desa checklist i factura com ODS i els llibres de notes com XLSX.

Exercici 2: fes una captura abans; obre Visualitza > Barres d'eines > Personalitza (o Eines > Personalitza), afegeix Desa, Desfés, Refés, Moneda, Percentatge, Ordena i Filtre; fes la captura després. Repeteix per als dos programes de Linux i els dos de Windows.

Exercicis 5 i 6: obre les plantilles corresponents a OpenOffice Calc o Excel, revisa el format i utilitza Fitxer > Anomena i desa / Desa com per exportar-les al format demanat.

Exercici 7: a B1 escriu =IF(A1>10;"Major que 10";IF(A1=10;"Igual a 10";"Menor que 10")); a B2 escriu =IF(MOD(A2;2)=0;"Parell";"Senar"). Per a conversió des de decimal: DEC2BIN, DEC2OCT i DEC2HEX; cap a decimal: BIN2DEC, OCT2DEC i HEX2DEC.

Les captures, la personalització, les importacions i l'exportació s'han de fer a les màquines virtuals; no es poden simular des del repositori.
