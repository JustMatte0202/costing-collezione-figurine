# Analisi tecnica e dei costi di una collezione di figurine

Caso di studio che ricostruisce da zero il costo di una collezione fittizia, **"Campioni 2027"**: album da 48 pagine, bustine da 6 figurine autoadesive, espositore da banco.
Le materie prime usano prezzi di mercato 2025-2026 con fonte; le lavorazioni sono stime dichiarate.

**[Pagina interattiva con preventivatore](https://JustMatte0202.github.io/costing-collezione-figurine/)** · [Modello Excel](output/modello_costi.xlsx) · [Report PDF](output/report_caso_studio.pdf)

## Domande a cui risponde
| Domanda | Dove |
|---|---|
| Di cosa è fatto il prodotto e come si stampa? | `Distinta_Tecnica` |
| Quanto costa ogni unità? | `Scheda_Costi` |
| Come cambia il costo con tiratura e prezzo della carta? | `Preventivatore`, `Sensitivita` |
| Quale fornitore scegliere? | `Fornitori` (scoring pesato) |
| Stampare internamente o acquistare? | `Make_or_Buy` |
| Quanto costano le innovazioni? | `Nuovo_Prodotto` |

## Risultati principali
| Indicatore | Valore |
|---|---|
| Costo pieno bustina (1.000.000 pz) | € 0,041 |
| Costo pieno album (100.000 pz) | € 0,65 |
| Quota carta e cartoncino sul costo variabile dell'album | 63% |
| Tiratura di pareggio make-or-buy figurine | ≈ 12.700 bustine |
| Extra costo bustina plastic-free | +4,9% sul costo variabile |

## Metodo
- Imposizione su foglio macchina 70×100 (16 pp/foglio interni, 180 figurine/foglio).
- Costo pieno = (costo variabile + costi di avviamento ÷ tiratura) × (1 + 15% overhead).
- Tutte le ipotesi sono nel foglio `Ipotesi` (input in blu) e ogni risultato è una formula collegata.

## Struttura
```
output/   modello_costi.xlsx, report_caso_studio.pdf
docs/     index.html (pagina GitHub Pages)
src/      build_model.py (genera il modello Excel con openpyxl)
```
Rigenerare il modello: `pip install openpyxl` e poi `python src/build_model.py`.

## Fonti
- Pricepedia, *Graphic Paper Prices*, luglio 2025
- FreshPlaza, aprile 2026 · Pulpapernews, maggio 2026 (rincari carta)
- ChemAnalyst, *BOPP Film Prices* Q2 2026
- Energia B2B su dati ARERA (costo energia imprese)
- Lexplain, tabelle CCNL Grafici Editoriali industria 2024-2026
- Muscara.com (prezzo al pubblico bustina)

## Limiti
Collezione e fornitori sono fittizi e non rappresentano dati di alcuna azienda. Carta adesiva, cartoncino GC1, lavorazioni esterne e costi orari sono stime da validare.

---
Autore: Matteo · settembre 2026
