# ARR Jelentés

Reptéri érkezési transzfer-jelentés készítő. A „printout-with-passengers” PDF-ből kiolvassa a hajó
transzfereit és kísérőit, lekéri a valós érkezési időket, és kitölti az ARR Excel sablont
(`ÉÉÉÉ.HH.NN.HAJÓ.xlsm`).

Két változat, ugyanazzal a logikával:

| | Hol fut | Mappa |
|---|---|---|
| **Asztali** | Windows (EXE) | `arr_jelentes.py` |
| **Web-app** | iPhone, iPad, Android, böngésző – telepítés nélkül, kezdőképernyőre tehető | `mobil/` |

## Működés

1. **PDF** – táblázatként olvassa (cellák + szövegpozíciók): indulási idő, járatok, létszám, jármű kategória, Hosp/Escort nevek.
2. **AeroDataBox (RapidAPI)** – a BUD érkezési lista **egy hívással** (12 óránként 1): ebből jönnek az érkezési idők
   és a codeshare → üzemeltető cserék (pl. AA6685 → BA866). Ha egy járat nincs a listában, csak arra fut külön lekérés.
3. **Aviationstack** – tartalék a DL / AA / UA codeshare cseréhez, ha a listából nem derül ki (60 napos gyorsítótárral).
   A web-app előbb a PDF-ből párosít (azonos város + időpont), mert az ingyenes Aviationstack böngészőből nem érhető el.
4. **Excel** – közvetlen XML-szerkesztés: a makró, a legördülő listák és a formázás megmarad.
   A kezdés/végzés (From/To) üresen marad, a Sum Hours képlet megmarad.

API kulcsok: egyszer kell megadni, törölhetők. Asztalin a Windows-felhasználóhoz titkosítva (DPAPI),
a web-appban csak az adott eszközön tárolódnak. Havi API-számláló és keret-csík; ha a RapidAPI szerint
elfogy a keret, a program nem küld több kérést.

## Építés / feltöltés

- EXE: `.\build.ps1` → `dist\ARR_Jelentes.exe` (a `sablon.xlsm` beépül; a programban a „Másik sablon…” gombbal cserélhető)
- Web-app: `.\deploy-mobil.ps1` (Netlify)

> A beépített sablon (BASZÓ ARR 2024 v3.2) mindkét változatban benne van, és mindkettőben cserélhető.
> A PDF-ek, a generált jelentések és az API kulcsok szándékosan nincsenek a repóban.
> Figyelem: a web-app nyilvános, így a `mobil/sablon.xlsm` a linken keresztül letölthető.
