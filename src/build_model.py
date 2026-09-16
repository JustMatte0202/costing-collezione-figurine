from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference, BarChart
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation

NAVY="1F3A5F"; ACC="E8A33D"; INP="0000FF"
F=lambda **k: Font(name="Arial", **k)
thin=Side(style="thin",color="D0D5DD"); B=Border(bottom=thin)
HFILL=PatternFill("solid",fgColor=NAVY); INFILL=PatternFill("solid",fgColor="FFF4DC"); TOT=PatternFill("solid",fgColor="E7EEF7")
wb=Workbook()

def title(ws,t,sub):
    ws["A1"]=t; ws["A1"].font=F(bold=True,size=16,color=NAVY)
    ws["A2"]=sub; ws["A2"].font=F(italic=True,size=9,color="667085")
    ws.sheet_view.showGridLines=False
def header(ws,r,vals):
    for i,v in enumerate(vals,1):
        c=ws.cell(r,i,v); c.font=F(bold=True,color="FFFFFF"); c.fill=HFILL; c.alignment=Alignment(vertical="center",wrap_text=True)
def widths(ws,w):
    for i,x in enumerate(w): ws.column_dimensions[chr(65+i)].width=x
def style_row(ws,r,n,fmt=None,inp=False,tot=False):
    for i in range(1,n+1):
        c=ws.cell(r,i); c.border=B
        if c.font is None or not c.font.bold: c.font=F(size=10)
    if tot:
        for i in range(1,n+1): ws.cell(r,i).fill=TOT; ws.cell(r,i).font=F(bold=True,size=10)

# ---------------- LEGENDA / SINTESI ----------------
S=wb.active; S.title="Sintesi"
# ---------------- IPOTESI ----------------
I=wb.create_sheet("Ipotesi")
title(I,"Ipotesi e prezzi di mercato","Celle blu su fondo giallo = input modificabili. Fonti indicate per ogni valore; 'Stima' = ipotesi dichiarata dell'autore. Aggiornato: settembre 2026.")
header(I,4,["Parametro","Valore","Unità","Tipo","Fonte / nota"])
P={}
rows=[
("carta_base","Carta patinata in fogli – prezzo medio UE 2025",1031,"€/t","Mercato","Pricepedia, 'Graphic Paper Prices: July 2025' (media 2025)"),
("carta_adj","Rincaro 2026 applicato",0.05,"%","Stima","Aumenti annunciati 2026 da Sappi (+50/70 €/t, poi +65 €/t) e Arctic Paper; FreshPlaza apr-2026, Pulpapernews mag-2026"),
("carta","Carta patinata 2026 (calcolata)","=B5*(1+B6)","€/t","Calcolo","Base 2025 × (1 + rincaro)"),
("gc1","Cartoncino GC1 per copertine",None,"€/t","Stima","Patinata × (1 + premio cartoncino); nessun listino pubblico gratuito"),
("gc1_prem","Premio cartoncino vs patinata",0.30,"%","Stima","Differenziale tipico FBB/patinata; da verificare con fornitori"),
("adesiva","Carta autoadesiva patinata (facestock+adesivo+liner)",0.55,"€/m²","Stima","Ordine di grandezza per etichettifici; da validare"),
("bopp_usd","Film BOPP – Germania Q2 2026",2653,"USD/t","Mercato","ChemAnalyst, BOPP Film Prices Q2 2026"),
("fx","Cambio USD per 1 EUR",1.16,"USD/€","Stima","Cambio medio ipotizzato"),
("bopp","Film BOPP (calcolato)","=B11/B12","€/t","Calcolo","USD/t ÷ cambio"),
("kwh","Energia elettrica imprese",0.265,"€/kWh","Mercato","ARERA via Energia B2B: 26,5 c€/kWh imprese italiane (2024)"),
("op","Costo pieno orario operatore",25,"€/h","Stima","CCNL Grafici-Editoriali industria, livello B2: minimo 1.713,68 €/mese (lug-dic 2026, Lexplain) + contingenza, 14 mensilità, contributi e TFR ≈ +35%, 1.650 h/anno"),
("mac","Costo orario macchina offset 70x100 (ammortamento + manutenzione)",60,"€/h","Stima","Esclusi energia e personale"),
("kw","Potenza media assorbita macchina offset",80,"kW","Stima","Macchina 4 colori formato 70x100"),
("vel","Velocità netta di stampa",9000,"fogli/h","Stima",""),
("ink","Inchiostro offset quadricromia",12,"€/kg","Stima",""),
("ink_g","Consumo inchiostro (4 colori)",1.6,"g/m² per lato","Stima",""),
("lastra","Lastra CTP",12,"€/lastra","Stima",""),
("avv_h","Ore di avviamento per lavoro",3,"h","Stima","Messa a registro, colore, prove"),
("scarto","Scarto di avviamento e tiratura",0.05,"%","Stima",""),
("oh","Costi generali (overhead)",0.15,"%","Stima","Applicati al costo industriale"),
("markup","Ricarico commerciale preventivi",0.25,"%","Stima",""),
("retail","Prezzo al pubblico bustina (6 figurine)",1.00,"€","Mercato","Rivenditore online Muscara.com, bustina Calciatori 2025-26"),
("canale","Quota distribuzione/edicola sul prezzo",0.35,"%","Stima","Aggio + distribuzione"),
]
r=5
for k,n,v,u,t,s in rows:
    I.cell(r,1,n); I.cell(r,2,v); I.cell(r,3,u); I.cell(r,4,t); I.cell(r,5,s)
    P[k]=f"Ipotesi!$B${r}"; style_row(I,r,5)
    c=I.cell(r,2)
    if t!="Calcolo" and v is not None: c.font=F(color=INP,size=10); c.fill=INFILL
    c.number_format='0%' if u=="%" else ('#,##0.000' if u in("€/kWh","€/m²") else '#,##0.00')
    I.cell(r,5).alignment=Alignment(wrap_text=True,vertical="top")
    r+=1
I[P["gc1"].split("!")[1].replace("$","")]=f"={P['carta']}*(1+{P['gc1_prem']})"
widths(I,[52,12,14,11,90])

def p(k): return P[k]

# ---------------- DISTINTA TECNICA ----------------
D=wb.create_sheet("Distinta_Tecnica")
title(D,"Distinta tecnica – Collezione 'Campioni 2027' (caso di studio)","Collezione fittizia costruita per l'esercizio: album + bustine di figurine autoadesive + espositore da banco.")
header(D,4,["Componente","Elemento costruttivo","Materiale","Grammatura / spessore","Formato","Tecnica di stampa","Lavorazioni cartotecniche","Interno / esterno"])
dist=[
("Album","Copertina","Cartoncino GC1","300 g/m²","47 × 30 cm aperta (con dorso)","Offset 4+0","Plastificazione lucida, cordonatura","Interno"),
("Album","Pagine interne (48 pp)","Carta patinata opaca","115 g/m²","23 × 30 cm","Offset 4+4","Piega in segnature, raccolta","Interno"),
("Album","Legatura","Punti metallici","—","—","—","Cucitura a punto metallico, rifilo 3 lati","Interno"),
("Figurina","Supporto","Carta autoadesiva patinata","80 g + liner","5 × 7 cm","Offset 4+0 (+ retro 1 colore sul liner)","Fustellatura a mezzo taglio, taglio in lotti da 6","Interno"),
("Figurina speciale","Variante olografica (1 ogni 6 bustine)","Film olografico laminato","—","5 × 7 cm","Offset UV 4+0","Laminazione a caldo","Esterno"),
("Bustina","Involucro","Film BOPP","30 µm","7 × 9 cm (fronte+retro)","Flessografia 4 colori","Saldatura a flow-pack","Esterno (film stampato)"),
("Espositore","Box da banco 50 bustine","Cartone microonda E","≈ 1,5 mm","15 × 12 × 8 cm","Offset 4+0 su copertina accoppiata","Accoppiatura, fustellatura, incollaggio","Esterno"),
]
for i,row in enumerate(dist,5):
    for j,v in enumerate(row,1): D.cell(i,j,v).alignment=Alignment(wrap_text=True,vertical="top")
    style_row(D,i,8)
widths(D,[16,26,24,18,24,22,34,18])
D["A14"]="Parametri di imposizione"; D["A14"].font=F(bold=True,color=NAVY,size=12)
header(D,15,["Parametro","Valore","Nota"])
imp=[("Formato foglio macchina (m²)",0.70,"70 × 100 cm"),
("Pagine interne per foglio (bianca+volta)",16,"8 pagine per lato"),
("Pagine interne album",48,""),
("Copertine per foglio",4,"Copertina aperta 47 × 30"),
("Figurine per foglio (con margini)",180,"Teorico 200; 10% perso in margini e crocini"),
("Figurine per bustina",6,""),
("Lastre per lavoro album",24,"3 segnature × 8 + copertina 4 (arrotondato)"),
("Lastre per lavoro figurine",8,"Per ogni 'serie' da ristampare"),
("Area film per bustina (m²)",0.0126,"2 × 0,07 × 0,09"),
("Grammatura film BOPP 30 µm (g/m²)",27.3,"Densità 0,91 g/cm³"),
("Bustine per espositore",50,""),
]
G={}
for i,(n,v,no) in enumerate(imp,16):
    D.cell(i,1,n); c=D.cell(i,2,v); c.font=F(color=INP,size=10); c.fill=INFILL; D.cell(i,3,no); style_row(D,i,3)
    G[n]=f"Distinta_Tecnica!$B${i}"
g=lambda n: G[n]

# ---------------- SCHEDA COSTI ----------------
C=wb.create_sheet("Scheda_Costi")
title(C,"Scheda costi di prodotto","Costi variabili per unità + costi fissi di lavoro (avviamento). Tutte le celle sono formule collegate a Ipotesi e Distinta_Tecnica.")
header(C,4,["Voce","Quantità per unità","Unità","Prezzo unitario (€)","Costo per unità (€)","Logica di calcolo"])
def line(r,name,q,u,pr,note,cost=None):
    C.cell(r,1,name); C.cell(r,2,q); C.cell(r,3,u); C.cell(r,4,pr)
    C.cell(r,5,cost if cost else f"=B{r}*D{r}"); C.cell(r,6,note)
    C.cell(r,2).number_format='#,##0.00000'; C.cell(r,4).number_format='#,##0.0000'; C.cell(r,5).number_format='€ #,##0.0000'
    style_row(C,r,6)
def sec(r,t):
    C.cell(r,1,t).font=F(bold=True,size=12,color=NAVY)

sec(6,"A. ALBUM (48 pagine + copertina)")
fogli_int=f"{g('Pagine interne album')}/{g('Pagine interne per foglio (bianca+volta)')}"
line(7,"Fogli macchina interni",f"={fogli_int}*(1+{p('scarto')})","fogli","", "pagine ÷ pagine per foglio × (1+scarto)",cost="=0")
line(8,"Fogli macchina copertina",f"=1/{g('Copertine per foglio')}*(1+{p('scarto')})","fogli","","1 ÷ copertine per foglio × (1+scarto)",cost="=0")
line(9,"Carta patinata 115 g",f"=B7*{g('Formato foglio macchina (m²)')}*115/1000000","t",f"={p('carta')}","fogli × m² × g/m² → tonnellate")
line(10,"Cartoncino GC1 300 g",f"=B8*{g('Formato foglio macchina (m²)')}*300/1000000","t",f"={p('gc1')}","")
line(11,"Inchiostro",f"=(B7*2+B8)*{g('Formato foglio macchina (m²)')}*{p('ink_g')}/1000","kg",f"={p('ink')}","interni 2 lati, copertina 1 lato")
line(12,"Tempo macchina (tiratura)",f"=(B7+B8)/{p('vel')}","h",f"={p('mac')}+{p('op')}*2","macchina + 2 addetti")
line(13,"Energia di stampa",f"=B12*{p('kw')}","kWh",f"={p('kwh')}","ore × kW")
line(14,"Plastificazione copertina",1,"pz",0.03,"Stima listino terzisti")
line(15,"Legatoria punto metallico + rifilo",1,"pz",0.04,"Stima")
for rr in (14,15): C.cell(rr,4).font=F(color=INP,size=10); C.cell(rr,4).fill=INFILL
C["A16"]="Costo variabile album"; C["E16"]="=SUM(E9:E15)"; style_row(C,16,6,tot=True); C["E16"].number_format='€ #,##0.0000'
C["A17"]="  di cui carta e cartoncino"; C["E17"]="=E9+E10"; C["E17"].number_format='€ #,##0.0000'; style_row(C,17,6)
C["A18"]="Costi fissi di lavoro album (avviamento + lastre)"; C["E18"]=f"={p('avv_h')}*({p('mac')}+{p('op')}*2+{p('kw')}*{p('kwh')})+{g('Lastre per lavoro album')}*{p('lastra')}"
C["F18"]="€ per lavoro, indipendente dalla tiratura"; style_row(C,18,6,tot=True); C["E18"].number_format='€ #,##0.00'

sec(20,"B. BUSTINA (6 figurine autoadesive)")
line(21,"Fogli macchina figurine",f"={g('Figurine per bustina')}/{g('Figure per foglio')}" if False else f"={g('Figurine per bustina')}/{g('Figurine per foglio (con margini)')}*(1+{p('scarto')})","fogli","","",cost="=0")
line(22,"Carta autoadesiva",f"=B21*{g('Formato foglio macchina (m²)')}","m²",f"={p('adesiva')}","")
line(23,"Inchiostro",f"=B21*{g('Formato foglio macchina (m²)')}*{p('ink_g')}*1.25/1000","kg",f"={p('ink')}","+25% per retro 1 colore")
line(24,"Tempo macchina stampa",f"=B21/{p('vel')}","h",f"={p('mac')}+{p('op')}*2","")
line(25,"Energia di stampa",f"=B24*{p('kw')}","kWh",f"={p('kwh')}","")
line(26,"Fustellatura mezzo taglio e taglio",1,"bustina",0.004,"Stima")
line(27,"Film BOPP",f"={g('Area film per bustina (m²)')}*{g('Grammatura film BOPP 30 µm (g/m²)')}/1000000","t",f"={p('bopp')}","m² × g/m² → t")
line(28,"Stampa flessografica film",1,"bustina",0.004,"Stima terzista")
line(29,"Imbustamento flow-pack",f"=1/12000","h",f"={p('op')}+30","12.000 bustine/h; macchina 30 €/h")
line(30,"Espositore da banco (quota)",f"=1/{g('Bustine per espositore')}","espositore",0.35,"Prezzo espositore stimato")
for rr in (26,28,30): C.cell(rr,4).font=F(color=INP,size=10); C.cell(rr,4).fill=INFILL
C["A31"]="Costo variabile bustina"; C["E31"]="=SUM(E22:E30)"; style_row(C,31,6,tot=True); C["E31"].number_format='€ #,##0.0000'
C["A32"]="  di cui carta autoadesiva"; C["E32"]="=E22"; C["E32"].number_format='€ #,##0.0000'; style_row(C,32,6)
C["A33"]="Costi fissi di lavoro bustine"; C["E33"]=f"={p('avv_h')}*({p('mac')}+{p('op')}*2+{p('kw')}*{p('kwh')})+{g('Lastre per lavoro figurine')}*{p('lastra')}"
style_row(C,33,6,tot=True); C["E33"].number_format='€ #,##0.00'

sec(35,"C. Tiratura di riferimento e costo pieno")
header(C,36,["Voce","Album","Bustina"])
C["A37"]="Tiratura di riferimento (pz)"; C["B37"]=100000; C["C37"]=1000000
for x in("B37","C37"): C[x].font=F(color=INP,size=10); C[x].fill=INFILL; C[x].number_format='#,##0'
C["A38"]="Costo industriale unitario"; C["B38"]="=E16+E18/B37"; C["C38"]="=E31+E33/C37"
C["A39"]="Costo pieno (con overhead)"; C["B39"]=f"=B38*(1+{p('oh')})"; C["C39"]=f"=C38*(1+{p('oh')})"
C["A40"]="Prezzo al pubblico"; C["B40"]=""; C["C40"]=f"={p('retail')}"
C["A41"]="Ricavo netto editore stimato"; C["C41"]=f"=C40*(1-{p('canale')})"
C["A42"]="Incidenza costo pieno sul ricavo netto"; C["C42"]="=C39/C41"
C["A43"]="Nota"; C["B43"]="Il margine residuo copre licenze, marketing, logistica e utile: voci escluse dal perimetro dell'analisi."
for rr in range(37,43):
    style_row(C,rr,3)
    for x in "BC":
        if rr in(38,39,40,41): C[f"{x}{rr}"].number_format='€ #,##0.0000'
C["C42"].number_format='0.0%'
widths(C,[44,18,12,18,18,50])

# ---------------- PREVENTIVATORE ----------------
Q=wb.create_sheet("Preventivatore")
title(Q,"Preventivatore e curva di costo","Inserisci tiratura e ricarico: il preventivo si aggiorna. Tabella sotto = economie di scala.")
Q["A4"]="Prodotto"; Q["B4"]="Bustina"
dv=DataValidation(type="list",formula1='"Album,Bustina"'); Q.add_data_validation(dv); dv.add("B4")
Q["A5"]="Tiratura richiesta (pz)"; Q["B5"]=750000
Q["A6"]="Ricarico commerciale"; Q["B6"]=f"={p('markup')}"
for x in("B4","B5"): Q[x].font=F(color=INP,size=11,bold=True); Q[x].fill=INFILL
Q["B5"].number_format='#,##0'; Q["B6"].number_format='0%'
Q["A8"]="Costo variabile unitario"; Q["B8"]='=IF(B4="Album",Scheda_Costi!E16,Scheda_Costi!E31)'
Q["A9"]="Costi fissi di lavoro"; Q["B9"]='=IF(B4="Album",Scheda_Costi!E18,Scheda_Costi!E33)'
Q["A10"]="Costo pieno unitario"; Q["B10"]=f"=(B8+B9/B5)*(1+{p('oh')})"
Q["A11"]="Prezzo unitario di offerta"; Q["B11"]="=B10*(1+B6)"
Q["A12"]="Valore totale preventivo"; Q["B12"]="=B11*B5"
Q["A13"]="Margine lordo del lavoro"; Q["B13"]="=(B11-B10)*B5"
for rr in range(4,14):
    Q[f"A{rr}"].font=F(size=11,bold=rr in(11,12)); Q[f"A{rr}"].border=B; Q[f"B{rr}"].border=B
for x in("B8","B10","B11"): Q[x].number_format='€ #,##0.0000'
for x in("B9","B12","B13"): Q[x].number_format='€ #,##0'
Q["B11"].fill=TOT; Q["B12"].fill=TOT
header(Q,16,["Tiratura","Costo pieno album (€)","Costo pieno bustina (€)","Incidenza fissi bustina"])
tir=[25000,50000,100000,250000,500000,1000000,2000000,5000000]
for i,t in enumerate(tir,17):
    Q.cell(i,1,t).number_format='#,##0'
    Q.cell(i,2,f"=(Scheda_Costi!$E$16+Scheda_Costi!$E$18/A{i})*(1+{p('oh')})").number_format='€ #,##0.0000'
    Q.cell(i,3,f"=(Scheda_Costi!$E$31+Scheda_Costi!$E$33/A{i})*(1+{p('oh')})").number_format='€ #,##0.0000'
    Q.cell(i,4,f"=(Scheda_Costi!$E$33/A{i})/(Scheda_Costi!$E$31+Scheda_Costi!$E$33/A{i})").number_format='0.0%'
    style_row(Q,i,4)
ch=LineChart(); ch.title="Costo pieno unitario bustina vs tiratura"; ch.y_axis.title="€ / bustina"; ch.x_axis.title="Tiratura"
ch.add_data(Reference(Q,min_col=3,min_row=16,max_row=24),titles_from_data=True)
ch.set_categories(Reference(Q,min_col=1,min_row=17,max_row=24)); ch.height=8; ch.width=16
ch.series[0].graphicalProperties.line.solidFill=NAVY; ch.y_axis.delete=False; ch.x_axis.delete=False
Q.add_chart(ch,"F4")
widths(Q,[30,22,24,22])

# ---------------- SENSITIVITY ----------------
T=wb.create_sheet("Sensitivita")
title(T,"Analisi di sensitività – costo pieno album","Righe: variazione prezzo carta/cartoncino. Colonne: tiratura. Valori in €/album.")
T["A4"]="Δ prezzo carta \\ Tiratura"; T["A4"].font=F(bold=True,color="FFFFFF"); T["A4"].fill=HFILL
tt=[25000,50000,100000,250000,500000]
for j,t in enumerate(tt,2):
    c=T.cell(4,j,t); c.number_format='#,##0'; c.font=F(bold=True,color="FFFFFF"); c.fill=HFILL
deltas=[-0.2,-0.1,0,0.1,0.2,0.3]
for i,d in enumerate(deltas,5):
    c=T.cell(i,1,d); c.number_format='+0%;-0%;0%'; c.font=F(bold=True,color=INP); c.fill=INFILL
    for j in range(2,2+len(tt)):
        col=chr(64+j)
        T.cell(i,j,f"=(Scheda_Costi!$E$16+Scheda_Costi!$E$17*$A{i}+Scheda_Costi!$E$18/{col}$4)*(1+{p('oh')})").number_format='€ #,##0.000'
T.conditional_formatting.add("B5:F10",ColorScaleRule(start_type="min",start_color="D9EAD3",mid_type="percentile",mid_value=50,mid_color="FFF2CC",end_type="max",end_color="F4CCCC"))
T["A13"]="Lettura: a tirature basse pesa l'avviamento; a tirature alte il costo è guidato dalla carta, quindi la negoziazione sulle materie prime diventa la leva principale."
T["A13"].font=F(italic=True,size=10)
widths(T,[26,14,14,14,14,14])

# ---------------- FORNITORI ----------------
R=wb.create_sheet("Fornitori")
title(R,"Listino e valutazione fornitori esterni – espositore da banco","Fornitori fittizi. Punteggio pesato: pesi modificabili (somma = 100%).")
header(R,4,["Criterio","Peso"])
crit=[("Prezzo",0.40),("Lead time",0.20),("Qualità (difettosità)",0.25),("Puntualità consegne",0.15)]
for i,(n,w) in enumerate(crit,5):
    R.cell(i,1,n); c=R.cell(i,2,w); c.number_format='0%'; c.font=F(color=INP); c.fill=INFILL; style_row(R,i,2)
R["A9"]="Totale"; R["B9"]="=SUM(B5:B8)"; R["B9"].number_format='0%'; style_row(R,9,2,tot=True)
header(R,11,["Fornitore","Prezzo 1–9.999 pz (€)","Prezzo 10.000–49.999 (€)","Prezzo ≥ 50.000 (€)","Lead time (gg)","Difettosità (ppm)","Puntualità (%)","Punteggio prezzo","Punteggio lead time","Punteggio qualità","Punteggio puntualità","Punteggio totale","Classifica"])
fo=[("Fornitore Alfa – Emilia",0.42,0.36,0.33,15,800,0.96),("Fornitore Beta – Lombardia",0.39,0.34,0.31,25,1500,0.90),("Fornitore Gamma – Est Europa",0.33,0.28,0.25,35,2500,0.85)]
for i,row in enumerate(fo,12):
    for j,v in enumerate(row,1):
        c=R.cell(i,j,v)
        if 2<=j<=7: c.font=F(color=INP,size=10); c.fill=INFILL
        if 2<=j<=4: c.number_format='€ 0.00'
        if j==7: c.number_format='0%'
    R.cell(i,8,f"=MIN($D$12:$D$14)/D{i}*100")
    R.cell(i,9,f"=MIN($E$12:$E$14)/E{i}*100")
    R.cell(i,10,f"=MIN($F$12:$F$14)/F{i}*100")
    R.cell(i,11,f"=G{i}/MAX($G$12:$G$14)*100")
    R.cell(i,12,f"=H{i}*$B$5+I{i}*$B$6+J{i}*$B$7+K{i}*$B$8")
    R.cell(i,13,f"=RANK(L{i},$L$12:$L$14)")
    for j in range(8,13): R.cell(i,j).number_format='0.0'
    style_row(R,i,13)
R["A16"]="Costo annuo per fascia"; R["A16"].font=F(bold=True,color=NAVY)
R["A17"]="Fabbisogno annuo espositori"; R["B17"]=60000; R["B17"].font=F(color=INP); R["B17"].fill=INFILL; R["B17"].number_format='#,##0'
header(R,18,["Fornitore","Costo annuo (€)","Δ vs migliore (€)"])
for i in range(3):
    r0=19+i
    R.cell(r0,1,f"=A{12+i}")
    R.cell(r0,2,f"=$B$17*IF($B$17>=50000,D{12+i},IF($B$17>=10000,C{12+i},B{12+i}))").number_format='€ #,##0'
    R.cell(r0,3,f"=B{r0}-MIN($B$19:$B$21)").number_format='€ #,##0'
    style_row(R,r0,3)
R["A23"]="Lettura: Gamma è il più economico ma Alfa vince sul punteggio pesato. Il risparmio annuo di Gamma va confrontato con lead time più che doppio e difettosità tripla (rischio di fermi e resi)."
R["A23"].font=F(italic=True,size=10)
widths(R,[30,14,14,14,12,12,12,12,12,12,12,12,11])
bc=BarChart(); bc.title="Punteggio pesato"; bc.add_data(Reference(R,min_col=12,min_row=11,max_row=14),titles_from_data=True)
bc.set_categories(Reference(R,min_col=1,min_row=12,max_row=14)); bc.height=7; bc.width=14; bc.legend=None
bc.series[0].graphicalProperties.solidFill=NAVY; bc.y_axis.delete=False; bc.x_axis.delete=False
R.add_chart(bc,"E17")

# ---------------- MAKE OR BUY ----------------
M=wb.create_sheet("Make_or_Buy")
title(M,"Make or buy – stampa figurine","Confronto tra produzione interna (costi da Scheda_Costi) e acquisto da stampatore esterno.")
M["A4"]="Prezzo esterno per 1.000 figurine stampate e fustellate (€)"; M["B4"]=9.5
M["A5"]="Figurine per bustina"; M["B5"]=f"={g('Figurine per bustina')}"
M["A6"]="Quota interna variabile per bustina (solo figurine)"; M["B6"]="=SUM(Scheda_Costi!E22:E26)"
M["A7"]="Costi fissi interni per lavoro"; M["B7"]="=Scheda_Costi!E33"
M["A8"]="Costo esterno per bustina (solo figurine)"; M["B8"]="=B4/1000*B5"
M["A9"]="Tiratura di pareggio (bustine)"; M["B9"]='=IF(B8>B6,B7/(B8-B6),"Esterno sempre conveniente")'
M["B4"].font=F(color=INP,bold=True); M["B4"].fill=INFILL
for rr in range(4,10): style_row(M,rr,2)
M["B4"].font=F(color=INP,bold=True)
for x in("B6","B8"): M[x].number_format='€ #,##0.0000'
M["B7"].number_format='€ #,##0'; M["B9"].number_format='#,##0'; M["B9"].fill=TOT
header(M,11,["Tiratura","Costo interno (€)","Costo esterno (€)","Scelta"])
for i,t in enumerate([5000,10000,25000,50000,100000,500000],12):
    M.cell(i,1,t).number_format='#,##0'
    M.cell(i,2,f"=$B$7+$B$6*A{i}").number_format='€ #,##0'
    M.cell(i,3,f"=$B$8*A{i}").number_format='€ #,##0'
    M.cell(i,4,f'=IF(B{i}<C{i},"MAKE","BUY")')
    style_row(M,i,4)
M["A19"]="Fattori qualitativi: capacità produttiva disponibile, riservatezza sulle collezioni, tempi di reazione per ristampe, controllo qualità del colore."
M["A19"].font=F(italic=True,size=10)
widths(M,[52,18,18,12])

# ---------------- NUOVO PRODOTTO ----------------
N=wb.create_sheet("Nuovo_Prodotto")
title(N,"Studio nuovi prodotti – impatto sul costo della bustina","Due varianti valutate rispetto alla bustina standard.")
header(N,4,["Variante","Parametro","Valore","Δ costo per bustina (€)","Δ % sul costo variabile","Nota"])
N["A5"]="1. Figurina olografica"; N["B5"]="Sovrapprezzo film olografico (€/m²)"; N["C5"]=0.90
N["A6"]=""; N["B6"]="Frequenza (1 ogni N bustine)"; N["C6"]=6
N["D5"]="=C5*0.05*0.07/C6"; N["F5"]="Area figurina 5×7 cm; costo spalmato su N bustine"
N["A8"]="2. Bustina in carta (plastic-free)"; N["B8"]="Carta barriera flessibile (€/t)"; N["C8"]=2100
N["B9"]="Grammatura (g/m²)"; N["C9"]=60
N["D8"]=f"={g('Area film per bustina (m²)')}*C9/1000000*C8-Scheda_Costi!E27"
N["F8"]="Sostituisce il film BOPP; saldatura a freddo da verificare con i fornitori"
N["B10"]="Rallentamento imbustamento"; N["C10"]=0.20
N["D10"]="=Scheda_Costi!E29*C10"; N["F10"]="Velocità macchina ridotta del 20%"
N["A12"]="Totale variante 2"; N["D12"]="=D8+D10"; style_row(N,12,6,tot=True)
for rr in (5,8,12):
    N[f"E{rr}"]=f"=D{rr}/Scheda_Costi!$E$31"; N[f"E{rr}"].number_format='+0.0%;-0.0%'
    N[f"D{rr}"].number_format='€ +0.0000;€ -0.0000'
N["D10"].number_format='€ +0.0000;€ -0.0000'
for x in("C5","C6","C8","C9","C10"): N[x].font=F(color=INP); N[x].fill=INFILL
N["C10"].number_format='0%'; N["C5"].number_format='€ 0.00'; N["C8"].number_format='#,##0'
for rr in range(5,12): style_row(N,rr,6)
N["A14"]="Valutazione: prezzo carta barriera e sovrapprezzo olografico sono stime, da validare con richieste di offerta a 2–3 fornitori prima di una decisione."
N["A14"].font=F(italic=True,size=10)
widths(N,[34,36,12,22,20,56])

# ---------------- SINTESI ----------------
title(S,"Collezione 'Campioni 2027' – Analisi tecnica e dei costi","Caso di studio di Matteo · Settembre 2026 · Prezzi di mercato 2025-2026 con fonti nel foglio Ipotesi")
header(S,4,["Indicatore","Valore","Riferimento"])
kp=[("Costo pieno album (100.000 pz)","=Scheda_Costi!B39","€ #,##0.000","Scheda_Costi"),
("Costo pieno bustina (1.000.000 pz)","=Scheda_Costi!C39","€ #,##0.0000","Scheda_Costi"),
("Incidenza costo pieno su ricavo netto bustina","=Scheda_Costi!C42","0.0%","Scheda_Costi"),
("Quota carta sul costo variabile album","=Scheda_Costi!E17/Scheda_Costi!E16","0.0%","Scheda_Costi"),
("Tiratura di pareggio make-or-buy figurine","=Make_or_Buy!B9","#,##0","Make_or_Buy"),
("Fornitore espositori consigliato",'=INDEX(Fornitori!A12:A14,MATCH(1,Fornitori!M12:M14,0))',"@","Fornitori"),
("Δ costo bustina con figurina olografica","=Nuovo_Prodotto!D5","€ +0.0000","Nuovo_Prodotto"),
("Δ costo bustina plastic-free","=Nuovo_Prodotto!D12","€ +0.0000;€ -0.0000","Nuovo_Prodotto"),
]
for i,(n,f,fm,rf) in enumerate(kp,5):
    S.cell(i,1,n); c=S.cell(i,2,f); c.number_format=fm; c.font=F(bold=True,size=11,color=NAVY); S.cell(i,3,rf); style_row(S,i,3)
    S.cell(i,2).font=F(bold=True,size=11,color=NAVY)
S["A15"]="Struttura del file"; S["A15"].font=F(bold=True,color=NAVY,size=12)
struct=[("Ipotesi","Prezzi di mercato con fonte e stime dichiarate"),("Distinta_Tecnica","Elementi costruttivi, materiali, stampa, cartotecnica, imposizione"),("Scheda_Costi","Costo variabile, costi fissi di lavoro e costo pieno"),("Preventivatore","Offerta per tiratura e curva delle economie di scala"),("Sensitivita","Effetto combinato prezzo carta × tiratura"),("Fornitori","Listino a scaglioni e scoring pesato"),("Make_or_Buy","Produzione interna vs esterna"),("Nuovo_Prodotto","Figurina olografica e bustina plastic-free")]
for i,(a,b) in enumerate(struct,16): S.cell(i,1,a).font=F(bold=True,size=10); S.cell(i,2,b).font=F(size=10)
S["A25"]="Legenda: testo blu su fondo giallo = input modificabile · nero = formula · fondo azzurro = totali."
S["A25"].font=F(italic=True,size=9,color="667085")
widths(S,[48,30,20])
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.font and c.font.name!="Arial": c.font=F(size=10)
wb.save("output/modello_costi.xlsx")
