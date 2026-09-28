# Mini HTML Editor 2.0

En WYSIWYG-editor för innehåll som ska klistras in i ett annat system. Den ger
**kompakt HTML på en enda rad med all CSS inline**, så att målsystemets
redigerare inte får extra blankrader och inte behöver några stilmallar.

Öppna `index.html` i en webbläsare (Chrome, Edge eller Firefox). Allt sparas
automatiskt i webbläsarens lokala lagring.

## Layout

| Område | Innehåll |
| --- | --- |
| Vänster | Stilbibliotek (kan döljas med knappen **Stilar**) |
| Mitten | Redigeringsytan, elementsökväg/inspektör och HTML-koden |
| Höger | Förhandsgranskning av exakt den kod som kopieras |

Dra i avdelarna för att ändra bredd på editor/förhandsgranskning och höjd på kodpanelen.

## Skriva och formatera

- **Enter** delar stycket vid markören. Efter en rubrik blir nästa rad brödtext.
- **Enter på en tom rad** ger en tomrad, som exporteras som `<p …>&nbsp;</p>` så att den syns.
- **Större mellanrum** utan tomrader: välj **↓** (avstånd efter) eller **↑** (avstånd före) i verktygsfältet.
- **Shift+Enter** ger radbrytning (`<br>`) inom stycket.
- **Enter på tom listpunkt** avslutar listan, **Enter på tom sista rad i en ruta** lämnar rutan.
- **Tab / Shift+Tab** ändrar nivå i listor och flyttar mellan tabellceller.
- **Ctrl+Z / Ctrl+Y** ångra och gör om, **Ctrl+K** länk, **Ctrl+Shift+V** klistra in som ren text.

## Stilar

Stilbiblioteket har tre sorters stilar:

- **Block** – gäller hela stycken (t.ex. *Brödtext*, *Rubrik*). Kan byta element, t.ex. `p` → `div`.
- **Ruta** – lägger markerade stycken i en box (t.ex. *Info-ruta*, *Sektion*).
- **Text** – gäller bara markerad text och blir en `<span style="…">`.

Klicka ✎ för att ändra en stil – text som redan använder stilen uppdateras, men
manuella justeringar (t.ex. ändrat avstånd) behålls. **Spara som stil…** under
editorn skapar en ny stil från elementet där markören står. Stilarna kan
exporteras och importeras under ⚙ Inställningar.

Sökvägen under editorn (t.ex. `wrapper › div ▭ Info-ruta › p Brödtext`) visar
elementen där markören står. Klicka på ett element för att se och ändra dess CSS direkt.

## Klistra in

- **Från Word/Google Docs/webben**: formatering rensas men rubriker, listor
  (även Word-listor och punkter som `•`), fet/kursiv, länkar och tabeller behålls.
- **HTML-kod** (t.ex. tidigare utdata) tolkas som HTML och behåller sina stilar.
  Klistra in i editorn eller via **Redigera HTML** i kodpanelen.

## Inställningar (⚙)

Yttre wrapper-div, CSS-variabler (t.ex. `--muBlue`, som målsystemet
definierar), standardstilar för stycken, listor, tabeller och länkar, samt om
sista elementets nedre marginal ska nollställas och om tomma stycken ska
exporteras som `&nbsp;`.

## Utveckling

Testerna kör editorn i en riktig webbläsare (Chromium via Playwright):

```bash
npm install
npm test
```
