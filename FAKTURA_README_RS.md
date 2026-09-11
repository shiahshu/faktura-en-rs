# faktura.py — Native PDF Invoice Extraction by Pure Rules (No LLM) / faktura.py — Izvlačenje faktura iz PDF-a čistim pravilima (bez LLM-a)

`faktura.py` koristi pdfplumber za direktno parsiranje **originalnih PDF faktura sa tekst slojem**, izvlačenje zaglavlja i stavki putem regex-a i tabelskih pravila. **Nema poziva LLM-a** — samo lokalno izvlačenje, brzo, nultotroškovno, determinističko i reprodukovano.

> Namena: originalne (ne-skenirane) PDF fakture izvedene iz srpskog eFaktura sistema.
> Skenirani (slike-only) PDF fajlovi nemaju tekst sloj, skripta ne može da ih obradi — koristite `racun.py` / `invoice.py`.

## Brza upotreba

```powershell
python faktura.py
```

Prozor sa listom fajlova se pojavljuje, izaberite PDF i obrada se izvodi automatski, čime se generiše:

```text
Zbir_Faktura.xlsx
```

Ako je Excel fajlo otvoren, automatski se čuva kao `Zbir_Faktura_FIX_HHMMSS.xlsx`.

**Jezičke varijante** (logika potpuno identična, samo jezički interfejs i ime izlaznog fajla se razlikuju):

| Skripta | Jezik | Izlazni fajl |
|---------|-------|--------------|
| `faktura.py` | Kineski | `发票汇总.xlsx` |
| `fakturaen.py` | Engleski | `invoice_summary.xlsx` |
| `fakturars.py` | Srpski | `Zbir_Faktura.xlsx` |

## Glavne funkcije

- Tkinter prozor sa listom fajlova **multiselekt** originalnih PDF faktura (moguće obrađivati više faktura odjednom).
- Čisto-pravilno izvlačenje (zasniva se na `match_invoices.py` pdfplumber logici):
  - Zaglavlje: broj fakture, PIB prodavca, PIB kupca, `DATUM PROMETA`, zvaničan iznos za plaćanje
  - Stavke: opis, količina, cena po jedinici, jedinica mere (`Jedinica_Mere`), popust, iznos bez PDV-a, PDV stopa, tekst specijalnog oporezivanja
  - Stavke podržavaju **tri kolone** (određeno po tome da li zaglavlje sadrži kolonu popusta): standardni režim (Količina Cena Jed Popust Neto PDV), režim specijalnog oporezivanja (Količina Cena Jed Popust Neto Tekst oporezivanja), režim bez popusta (Količina Cena Jed Neto PDV) — automatski prepoznato.
  - Završni nenumerički tokeni uvek se tretiraju kao tekst oporezivanja (npr. `Posebni postupci oporezivanja`) i nikada ne parsiraju kao brojevi.
- **PIB obavezan 9 cifara**: regex `[0-9]{9}` + granica `(?!\d)` tako da registarski brojevi i slični identifikatori nikada ne budu pomešani sa PIB-om.
- **Više faktura po strani**: PDF sa više faktura (npr. `*_sef.pdf` masovni fajlovi) generiše jednu fakturu po stranici.
- **Spajanje nastavaka na više strana**: kada se isti broj fakture proteže preko više strana, stavke nastavka se dodaju na prethodnu stranu; nastavci bez stavki (npr. avans nastavci) se automatski preskoče, sprečavajući dvostruko brojanje.
- **Podrška za eFaktura nastavak broja fakture**: pored `Broj fakture` i `račun br.`, prepoznaje se i format `Generisao sistem eFaktura pod brojem:00084/26`.
- **Spajanje footer nastavaka**: čak i bez stavki, nastavci stranica imaju svoje podatke o zvaničnom iznosu, popustu u avansu, PIB-u i datumu transakcije koji se spajaju naviše.
- PDV po stavci = `round(neto × stopa, 2)`; ukupno sa PDV-om = neto + PDV.
- Stopa se uzima iz stavke `PDV stopa`: čisti brojevi koriste tu stopu; nenumerički tekst oporezivanja (npr. `Nije predmet oporezivanja PDV`) daje 0 PDV — nikada ne primenjuje se 10%/20% iz footer-a.
- **Popust u avansu** (prema `avans3.py` semantici): **dinamička iteracija po svim stopama PDV-a**, prepoznavanje linija `Ukupna osnovica - stopa X%` / `Ukupan iznos osnovice umanjen za osnovicu po avansu - stopa X%` / `Ukupan PDV umanjen za PDV po avansu - stopa X%`, sabiranje po stopama razlike, daje `Prepayment_Deduction` (ukupni popust u avansu) i `Actual_Payment` (stvarni dugovan), uključeno u verifikaciju.
- Svaka faktura dobija sumirani red i verifikuje se sa zvaničnim iznosom (`OK` / `REVIEW` / `ERR`).
- **Razlaganje detalja verifikacije**: konzola štampa po fakturi `Neto / PDV / Ukupno / Avans / Plaćeno / Zvanično`, tako se svaka nesuglasnica može lokalizovati.
- Kreditni računi (storno / knjižno odobrenje / crvene knjige) automatski pretvarani u negativne vrednosti.
- Strukturisana tabela u stilu racun.py u PowerShell-u; isto se zapisuje u Excel.
- **Razlike u verifikaciji → Footer tabla**: kada postoji faktura sa statusom različitim od `OK`, ispod glavne tabele se automatski štampa dodatna tabela čisto na osnovu originalnih footer podataka PDF-a (Neto / PDV / Ukupno / Avans / Plaćeno), sa zglobnim SUM redom, za direktno poređenje sa zvaničnim podacima.
- **Bez kineske kolone za prevod** (bez LLM-a nema prevoda; po potrebi dodajte prevode drugim alatom).

## Razlika u odnosu na racun.py / invoice.py

| Stavka | faktura.py | racun.py / invoice.py |
|--------|------------|------------------------|
| Pozivi LLM-a | Ništa — čisto pdfplumber pravila | Da, Gemini parsiranje |
| Tip PDF-a | Samo originalni (tekst sloj) | Originalni + skenirani (invoice.py) |
| Prevod na kineski | Ne izlazi ta kolonica | Izlazi `Description_CN` |
| Brzina | Sekunde, nema troškova tokena | Zavisno od API-a, ograničen kvotom |
| Determinizam | Determinističko, reproduktivno | LLM halucinacion rizik |
| Masovni fajlovi sa više faktura | Automatski razdvojeno po stranici | Zavisno od prepoznavanja modela |

## Struktura izlaza

### Konzolna tabela

```text
Invoice_ID  Description  Prodavac PIB  Kupac PIB  DATUM PROMETA  Količina  Jed  Cena  Neto  PDV  Ukupno  Avans  Plaćeno  Status
```

Svaki sumirani red faktura prikazuje broj fakture, PIB prodavca i kupca, `DATUM PROMETA`, količinu, jedinicu (`Jed`), neto iznos, PDV, ukupno sa PDV-om, popust u avansu (`Avans`), stvarno plaćeno (`Plaćeno`) i status verifikacije.

### Excel kolone

| Kolona | Opis |
|--------|------|
| `Invoice_ID` | Broj fakture (npr. IF26-0134) |
| `Prodavac_PIB` | PIB prodavca; popunjeno samo u sumiranim redovima |
| `Kupac_PIB` | PIB kupca; popunjeno samo u sumiranim redovima |
| `Datum_prometa` | Datum transakcije/nabavke; popunjeno samo u sumiranim redovima |
| `Description` | Originalni srpski opis |
| `Količina` | Količina |
| `Jedinica_Mere` | Jedinica mere (npr. kom / t / km / paušal); popunjeno samo u redovima stavki |
| `Cena` | Cena po jedinici |
| `Neto` | Iznos bez PDV-a |
| `PDV_Stope` | PDV stopa (npr. 20%) |
| `PDV_Iznos` | Iznos PDV-a = round(Neto × stopa, 2) |
| `Ukupno` | Stavka = neto + PDV (sa PDV-om); sumirani/redovi ukupno = ukupno sa PDV-om pre popusta |
| `Avans` | Popust u avansu (sa PDV-om) = Σ po stopama (osnovica popusta + PDV popusta); popunjeno u sumiranim/ukupnim redovima |
| `Plaćeno` | Stvarno dugovan (nakon popusta) = zvanični iznos (`Iznos za plaćanje`); stavka = te PDV-cene; sumirani red = faktura dugovan |
| `Status` | Status verifikacije |

## Pravila za proveru iznosa

- PDV po stavci je nezavisno zaokružen na 2 decimale, pa se zatim sabira (po pravilu srpske faktura po starci).
- Popust u avansu: `Avans = Σ po stopama [(Ukupna osnovica − osnovica umanjena za avans) + (Ukupan PDV − PDV umanjen za avans)]`.
- Stvarno dugovan = ukupno sa PDV-om stavki − `Avans`.
- **Semantička razlika**: `Ukupno` = ukupno sa PDV-om pre popusta (preračunato iz stavki); `Plaćeno` = dugovan nakon popusta (zvanični `Iznos za plaćanje`, fallback `Ukupan iznos fakture`).

### Određivanje statusa verifikacije

```
calc  = sum(svih stavki Ukupno) − Avans   ← stvarno dugovan
off   = zvanični iznos (Iznos za plaćanje / Ukupan iznos fakture)
diff  = round(abs(calc − off), 2)        ← razlika zaokružena na 2 decimale
```

| Razlika | Status | Znacenje | Izlaz u konzoli |
|---------|--------|----------|-----------------|
| `0.00` | **OK** | Verifikacija prođena, stvarno dugovan odgovara zvaničnom | `✅ OK` |
| `≤ 0.01` | **REVIEW** | Razlika od centa, normalno odstupanje pri zaokruživanju, potrebna ručna provera | `⚠️ Razlika od centa, proverite` |
| `> 0.01` | **ERR** | Neslaganje u verifikaciji, ispitajte neto/PDV/popust/zvanično izvlačenje | `❌ Neslaganje` |

**Zašto 0.01 nije ERR?** PDV po stavci `round(Neto × stopa, 2)` se unosi po centu pa se sabira, dok zvanični broj sabira netoe pa primenjuje stopu — dva različita puta zaokruživanja prirodno proizvode ±0.01 odstupanje. Ovo je uobičajeno u srpskom eFaktura sistemu, a ne greška u izvlačenju.

- Kada zvanični iznos nije pronađen, koristi se stvarno dugovan i označava se `NO_OFFICIAL`.
- Bez avans popusta (`Avans = 0`), stvarno dugovan je jednako ukupnom sa PDV-om, što odgovara starom ponašanju.
- Regex za zvanični iznos je ograničen na tekuću liniju (`[^\r\n]*?`) da ne bi progutao sledeći broj.

## Footer poređenje (prikazuje se automatski kada postoje razlike u verifikaciji)

Kada postoji faktura sa statusom različitim od `OK`, ispod glavne tabele u konzoli se automatski štampa **Footer tabela podataka**:

```text
📋 Footer podaci za poređenje (samo faktura sa razlikama, svi iz originalnog footer PDF-a)
Broj Fakture                  Neto           PDV            Ukupno         Avans          Plaćeno
----------------------------------------------------------------------------------------------
IF26-0134                 12,000.00      2,400.00     14,400.00          0.00      14,400.00
IF26-0135                  8,500.00      1,700.00     10,200.00        500.00       9,700.00
----------------------------------------------------------------------------------------------
SUM                       20,500.00      4,100.00     24,600.00        500.00      24,100.00
```

| Kolona | Izvor | Opis |
|--------|-------|------|
| `Neto` | `footer_net` | Originalni PDF footer `Ukupna osnovica` / `Total net` |
| `PDV` | `footer_vat` | Originalni PDF footer `Ukupan PDV` / `Total VAT` |
| `Ukupno` | `footer_total` | Originalni PDF footer `Ukupan iznos fakture` |
| `Avans` | `prep_deduction` | Popust u avansu (sa PDV-om); određen po razlici ukupnog sa PDV-om i finalnog dugovanja |
| `Plaćeno` | `grand_total` | Stvarno dugovan = `Iznos za plaćanje` (nakon popusta) |
| `SUM` | Zbir svih gore navedenih kolona | Zglobni ukupan red |

> Tabella za poređenje potiče **isključivo iz originalnog footer PDF-a**, ne koristi se u izračunavanju stavki i pomaže brzo lokalizovanje izvora razlike (neto, PDV, popust ili zvanično izvlačenje).

## Pravila za izvlačenje zaglavlja

| Polje | Metod izvlačenja |
|-------|------------------|
| `Invoice_ID` | `Broj fakture` / `račun br.` / `Generisao sistem eFaktura pod brojem` / `IF\d{2}-\d{3,}` itd. |
| `Prodavac_PIB` | PIB u bloku `Prodavac`; `[0-9]{9}` + granica `(?!\d)`, obavezno 9 cifara |
| `Kupac_PIB` | PIB posle `PIB kupca`; isto obavezno 9 cifara |
| `Datum_prometa` | Datum posle `Datum prometa` |
| Zvanični iznos | `Iznos za plaćanje` je prioritet, fallback `Ukupan iznos fakture` / `Grand total` |
| Popust u avansu | Izvlači se iz `Ukupan iznos fakture − Iznos za plaćanje`; prepoznaje se samo kada je ukupno sa PDV-om strogo veće od finalnog dugovanja |

Samo 9-cifreni brojevi koji su eksplicitno prisutni na fakturi se tretiraju kao PIB-ovi; MB, telefonski brojevi i drugi identifikatori nikada ne budu pomešani sa PIB-om.

## Ograničenja popusta u avansu

Skripta više ne veruje individualnim avans poljima, prvo uspostavlja Footer zatvaranje iznosa:

```text
Ukupno sa PDV-om       = Ukupan iznos fakture
Finalno dugovan        = Iznos za plaćanje
Popust u avansu        = Ukupno sa PDV-om − Finalno dugovan
```

Ograničenja:

- Kada je ukupno sa PDV-om jednako finalnom dugovanju, popust je obavezan `0.00`;
- Popust nikada ne sme biti negativan;
- Samo kako u Footer-u postoje važeća ukupna vrednost sa PDV-om i finalno dugovan, i ukupno je strože veće, popust je potvrđen;
- `Ukupna osnovica ... umanjen ...` i `Ukupan PDV ... umanjen ...` linije su samo referentni podaci za osnovicu/PDV nakon popusta i ne mogu samostalno utvrditi popust;
- Kada Footer ne može da formira važeće zatvaranje, sumnjivi popust se ignoriše i štampa upozorenje;
- `0.01` PDV razlike može se tretirati kao zaokružno odstupanje, samo iznad tog opsega postaje `ERR`.

Na primer:

```text
Ukupan iznos fakture = 4,360,000.00
Iznos za plaćanje    = 2,180,000.00
Popust u avansu       = 2,180,000.00
```

Ako su oba `1,286,000.00`, popust u avansu mora biti `0.00` — pojava avans povezanih teksta na stranici sama po sebi ne opravdava dodatno oduzimanje.

> **Detalj implementacije**: Popust se proračunava unutar `build_report` na osnovu Footer zatvaranja, prepisivajući originalnu vrednost `extract_avans_deduction()`. `extract_avans_deduction()` i dalje **dinamički iterira po svim stopama PDV-a**, računajući po stopi `(Ukupna osnovica − osnovica umanjena za avans) + (Ukupan PDV − PDV umanjen za avans)` pa sabira, i služi kao vrednost kolone `Avans` u tabeli za poređenje footer-a kao i referentna vrednost za upozorenje kada zatvaranje ne uspe.

## Nivoi statusa

Svaka faktura zadržava tri glavna statusa verifikacije:

| Status | Uslov |
|--------|-------|
| `OK` | Stvarno dugovan odgovara zvaničnom iznosu |
| `REVIEW` | Samo `0.01` zaokružno odstupanje, potrebna ručna provera |
| `ERR` | Razlika prelazi `0.01`, ili popust/stavke/zvanični iznos ne mogu da se zatvore |

Svaka faktura zadržava tri glavna statusa: `završeno`, `OK`, `REVIEW`, `ERR`, `NO_OFFICIAL`.

`ERR` ne znači nužno da je faktura pogrešna — može takođe značiti da je stavka propuštena, Footer polja su nekompletna ili format nekompatibilan. Uvek proverite originalni PDF.

## Neophodni uslovi

### Radno okruženje

- Python 3.x (Windows zvanični instalator uključuje `tkinter` — nema potrebe za dodatnom instalacijom)
- Trebate instalirati: `pdfplumber`, `pandas`, `openpyxl`
  - `tkinter` je standardna biblioteka (`Tk`/`filedialog`); uključena po default-u sa python.org instalacijama

### Instalacija jednom komandom na novom računaru

```powershell
pip install pdfplumber pandas openpyxl
```

### Verikovane verzije u trenutnom okruženju (referencija)

| Zavisnost | Verifikovana verzija |
|-----------|---------------------|
| Python | 3.12.10 |
| pdfplumber | 0.11.5 |
| pandas | 2.3.3 |
| openpyxl | 3.1.5 |
| tkinter | Standardna biblioteka |

### Zahtevi za PDF fajlove

- Mora biti **originalni PDF faktura sa tekst slojem** (eFaktura export; Ctrl+A selektovan tekst radi).
- Skenirani / samo-slike PDF fajlovi nemaju tekst sloj, skripta će prikazati "prazan PDF / nema tekst sloja" i ne može da se obradi — koristite `racun.py` / `invoice.py`.
- Faktura mora biti srpskog eFaktura formata: zaglavlje sadrži `Broj fakture`, `Iznos za plaćanje`, stavke sadrže količinu/cenu po jedinici/jedinicu/popust/neto/stopu. Drugi jezici ili formati mogu biti nepotpuno prepoznati.
- Jedan PDF može sadržati više faktura (npr. `*_sef.pdf`); skripta automatski razdvaja po stranicama. Jedna faktura može da se proteže preko više strana (nastavci se automatski spajaju/preskaču).

## Licenca
