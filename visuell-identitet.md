# Visuell identitet – INKA AI Labs

## Färgpalett

Exakta CSS-variabler från varumärket:

| Variabel | Hex | Användning |
|----------|-----|------------|
| `--labs-orange` | `#D95818` | Primär accentfärg – knappar, ikoner, hover-states, CTA |
| `--labs-dark-red` | `#640D02` | Djup varmt rödbrun – kontrast och tyngd |
| `--labs-light-yellow` | `#FDFAF1` | Bakgrundsfärg – varm off-white, ger luftig känsla |
| `--labs-text-yellow` | `#FFF2D9` | Ljusgul – används i textbakgrunder och mjuka sektioner |
| `--labs-offblack` | `#2B2B2B` | Primär textfärg – nästan svart, hög läsbarhet |
| `--labs-white` | `#FFFFFF` | Ren vit – bakgrunder och kontrast |

Känslan är **varm men professionell** – orangen `#D95818` signalerar energi och framåtanda, förankrad av den djupa rödbrun `#640D02` och mjuka gula bakgrundstoner.

## Typsnitt

- **Sans-serif** genomgående – modern, ren och lättläst
- Tydlig typografisk hierarki: stora rubriker, luftig brödtext
- Inget utsmyckat – typografin är ett verktyg för tydlighet, inte dekoration

## Layout och design

- **Minimalistisk** – generöst med whitespace, inget överflöd
- **Sektionsbaserad** – tydliga block med ett budskap per sektion
- **Bildspråk** – professionell fotografi av team och kontor, teknikframåt estetik
- **Ikoner** – enkla och konsekventa, används för att strukturera tjänsteöversikter

## Designkänsla i ord

Ren. Modern. Trovärdig. Varm men seriös. Framåtlutad.

## Navbar / Header

Navbaren använder organiska blob-former som bakgrund – inspirerat av hemsidans hero-sektion.

**Bakgrundsfärg:** `#C94E10` (mörkare orange bas)

**Blob-lager (SVG):**
- Stor gyllengul våg: `#E8950A` med opacity 0.75 – skapar rörelse och värme
- Beige stänk övre vänster: `#FDFAF1` med opacity 0.18 – mjuk koppling till sidans bakgrundsfärg
- Vinröd cirkel nedre höger: `#8B1A1A` med opacity 0.55 – tyngd och kontrast
- Extra ljusorange accent-våg: `#F0A030` med opacity 0.35

**Text och logotyp:** vit (`#FFFFFF`) och halvtransparent vit för undertext (`rgba(255,255,255,0.65)`)

**Känsla:** Organisk, varm, rörlig – signalerar kreativitet och energi utan att vara skrikig.

**Implementation:** Inline SVG med `viewBox="0 0 1200 80"` och `preserveAspectRatio="xMidYMid slice"` placerat som absolut bakgrundslager. Navbar-innehåll får `position: relative; z-index: 1`.

## Riktlinjer för framtida material

- Håll vita ytor – undvik att packa material för tätt
- Använd orange/gul som accentfärg, inte som bakgrund i stora ytor
- Välj professionella bilder med riktiga människor framför stockfoto-känsla
- Konsistens i ikonstil – enkla linjeikoner passar varumärket
- Undvik mörka, tunga bakgrunder – varumärket lever i ljuset
