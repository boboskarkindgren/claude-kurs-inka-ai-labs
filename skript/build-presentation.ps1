# INKA AI Labs - Reinholt Logistik Presentation Builder
# Creates presentation via PowerPoint COM automation

$outputPath = "C:\Users\kindg\OneDrive\Dokument\Claude kurs - INKA AI Labs\output\reinholt-logistik-presentation.pptx"

# Colors: R + G*256 + B*65536
$C = @{
    Orange    = 1595609   # D95818
    DarkRed   = 134500    # 640D02
    OffWhite  = 15858429  # FDFAF1
    Yellow    = 14283519  # FFF2D9
    OffBlack  = 2829099   # 2B2B2B
    White     = 16777215  # FFFFFF
    MidGray   = 9079434   # 8A8A8A
    LightGray = 13424864  # E0D8CC
    Dark4A    = 4868682   # 4A4A4A
    SoftYellow= 15792600  # F0E0B8 (stat box bg)
    Dark3A    = 3815994   # 3A3A3A
    AccentLight = 14215423 # FFE8D8
}

# Helper: add a text box and return the shape
function Add-TextBox($slide, $left, $top, $width, $height, $text, $fontSize, $bold=$false, $italic=$false, $color=$C.OffBlack, $align=2, $vAlign=1) {
    $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
    $tf = $shape.TextFrame
    $tf.WordWrap = -1
    $tf.AutoSize = 0
    $tf.TextRange.Text = $text
    $tf.TextRange.Font.Size = $fontSize
    $tf.TextRange.Font.Bold = if ($bold) { -1 } else { 0 }
    $tf.TextRange.Font.Italic = if ($italic) { -1 } else { 0 }
    $tf.TextRange.Font.Color.RGB = $color
    $tf.TextRange.ParagraphFormat.Alignment = $align  # 1=left, 2=center, 3=right
    $shape.TextFrame.VerticalAnchor = $vAlign  # 1=top, 3=middle
    $shape.Line.Visible = 0
    $shape.Fill.Visible = 0
    return $shape
}

# Helper: add a filled rectangle
function Add-Rect($slide, $left, $top, $width, $height, $fillColor, $lineVisible=$false) {
    $shape = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
    $shape.Fill.Solid()
    $shape.Fill.ForeColor.RGB = $fillColor
    if ($lineVisible) {
        $shape.Line.Visible = -1
        $shape.Line.ForeColor.RGB = $fillColor
    } else {
        $shape.Line.Visible = 0
    }
    return $shape
}

# Helper: set slide background color
function Set-BgColor($slide, $color) {
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = $color
    $slide.FollowMasterBackground = 0
}

# Helper: add notes to slide
function Add-Notes($slide, $text) {
    $slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = $text
}

# Helper: make a card (white box with optional orange top bar)
function Add-Card($slide, $left, $top, $width, $height) {
    $card = Add-Rect $slide $left $top $width $height $C.White
    $card.Shadow.Visible = -1
    $card.Shadow.Type = 1
    $card.Shadow.OffsetX = 3
    $card.Shadow.OffsetY = 3
    $card.Shadow.Blur = 8
    $card.Shadow.ForeColor.RGB = 0
    $card.Shadow.Transparency = 0.88
    # Orange top bar
    $bar = Add-Rect $slide $left $top $width 5 $C.Orange
    return $card
}

# ============================
Write-Host "Starting PowerPoint COM..."
$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.Visible = -1  # msoTrue

$pres = $pptApp.Presentations.Add(-1)

# Set to widescreen 16:9 (960 x 540 points)
$pres.PageSetup.SlideWidth  = 960
$pres.PageSetup.SlideHeight = 540

Write-Host "Building slides..."

# ============================================================
# SLIDE 1: TITLE (dark background)
# ============================================================
$s = $pres.Slides.Add(1, 12)  # ppLayoutBlank
Set-BgColor $s $C.OffBlack

# Left orange bar
Add-Rect $s 0 0 11 540 $C.Orange | Out-Null

# INKA AI Labs label
Add-TextBox $s 30 30 400 30 "INKA AI LABS" 11 $true $false $C.Orange 1 | Out-Null

# Main title
Add-TextBox $s 30 90 900 100 "Reinholt Logistik" 52 $true $false $C.White 1 | Out-Null
Add-TextBox $s 30 185 700 70 "× INKA AI Labs" 38 $false $false $C.Orange 1 | Out-Null

# Tagline
Add-TextBox $s 30 275 700 40 "Från AI-kaos till verklig affärsnytta" 17 $false $true $C.MidGray 1 | Out-Null

# Date
Add-TextBox $s 30 460 500 35 "13 juni 2026  ·  Anna Reinholt, COO" 12 $false $false 7368816 1 | Out-Null

Add-Notes $s "Välkomna! Det är ett privilegium att få träffa er. Vi på INKA AI Labs har lyssnat på vad ni berättat och ser direkt att Reinholt Logistik har rätt förutsättningar att ta AI från testfas till konkret affärsnytta. Idag visar vi hur vi tänker – och hur vi tänker jobba med er."

Write-Host "Slide 1 done"

# ============================================================
# SLIDE 2: VI FÖRSTÅR ER SITUATION
# ============================================================
$s = $pres.Slides.Add(2, 12)
Set-BgColor $s $C.OffWhite

Add-TextBox $s 45 25 870 58 "Vi förstår er situation" 30 $true $false $C.OffBlack 1 | Out-Null

# Pain points (left column)
$bullets = "AI används splittrat – utan röd tråd`r`nMedarbetare provar verktyg på egen hand`r`nSvårt att prioritera vad som ger mest värde`r`nTillväxten kräver mer kapacitet – men resurserna är desamma`r`nNi vet att potentialen finns – men inte hur ni tar den"
$tb = Add-TextBox $s 45 105 490 380 $bullets 15 $false $false $C.Dark4A 1 1
# Add bullets manually
$tf = $tb.TextFrame
for ($i = 1; $i -le $tf.TextRange.Paragraphs().Count; $i++) {
    $para = $tf.TextRange.Paragraphs($i)
    $para.ParagraphFormat.Bullet.Visible = -1
    $para.ParagraphFormat.Bullet.Character = 8226
    $para.ParagraphFormat.Bullet.Font.Color.RGB = $C.Orange
    $para.ParagraphFormat.SpaceAfter = 12
}

# Right: orange stat box
$box = Add-Rect $s 575 100 340 395 $C.Orange
$box.Shadow.Visible = -1; $box.Shadow.Type = 1; $box.Shadow.OffsetX = 4; $box.Shadow.OffsetY = 4; $box.Shadow.Blur = 10; $box.Shadow.ForeColor.RGB = 0; $box.Shadow.Transparency = 0.85

Add-TextBox $s 575 120 340 130 "280" 72 $true $false $C.White 2 3 | Out-Null
Add-TextBox $s 575 250 340 40 "anställda i tillväxt" 14 $false $false $C.White 2 | Out-Null
Add-TextBox $s 575 310 340 40 "2 nya storkundsavtal" 14 $true $false 14215423 2 | Out-Null
Add-TextBox $s 575 350 340 40 "som kräver skalbarhet" 12 $false $false 14215423 2 | Out-Null

Add-Notes $s "Ni berättade att AI-användningen idag är splittrad – och det är precis det vi ser hos de flesta företag i er fas. Det är inte ett tecken på misslyckande, utan ett typiskt stadium. Ni har provat, ni har lärt er – nu handlar det om att sätta riktning. Tillväxten med de nya storkunderna skapar dessutom ett naturligt tryck att göra mer med samma resurser."

Write-Host "Slide 2 done"

# ============================================================
# SLIDE 3: VARFÖR DET ÄR SVÅRT
# ============================================================
$s = $pres.Slides.Add(3, 12)
Set-BgColor $s $C.OffWhite

Add-TextBox $s 45 25 870 58 "Varför det är svårt" 30 $true $false $C.OffBlack 1 | Out-Null
Add-TextBox $s 45 78 870 35 "De flesta företag fastnar i samma tre hinder" 14 $false $true $C.MidGray 1 | Out-Null

$obstacles = @(
    @{ title="För brett"; body="AI kan göra allt – men inte allt på en gång. Utan prioritering sprids energin tunt och ingenting blir färdigt." },
    @{ title="För abstrakt"; body="Det är lätt att prata om AI-potential. Det svåra är att koppla det till era specifika processer och faktiska ROI." },
    @{ title="För ensamt"; body="Att driva AI-implementering internt utan rätt kompetens eller partner tar för lång tid och ger osäkra resultat." }
)

for ($i = 0; $i -lt 3; $i++) {
    $x = 45 + $i * 300
    $card = Add-Rect $s $x 130 275 370 $C.White
    $card.Shadow.Visible = -1; $card.Shadow.Type = 1; $card.Shadow.OffsetX = 3; $card.Shadow.OffsetY = 3; $card.Shadow.Blur = 8; $card.Shadow.ForeColor.RGB = 0; $card.Shadow.Transparency = 0.88
    $card.Line.Visible = -1; $card.Line.ForeColor.RGB = $C.LightGray; $card.Line.Weight = 1
    # Orange top accent
    Add-Rect $s $x 130 275 8 $C.Orange | Out-Null
    Add-TextBox $s ($x+15) 155 245 50 $obstacles[$i].title 20 $true $false $C.OffBlack 1 | Out-Null
    Add-TextBox $s ($x+15) 215 245 270 $obstacles[$i].body 13 $false $false $C.Dark4A 1 1 | Out-Null
}

Add-Notes $s "Det vi ser gång på gång är att det inte handlar om brist på ambition – det handlar om att AI-resan är genuint komplex. Tre hinder återkommer: man försöker göra för mycket på en gång, man pratar abstrakt istället för konkret, och man försöker göra det ensam. Det vi gör hos INKA AI Labs är att ta er igenom dessa tre hinder på ett strukturerat sätt."

Write-Host "Slide 3 done"

# ============================================================
# SLIDE 4: VÅR METODIK
# ============================================================
$s = $pres.Slides.Add(4, 12)
Set-BgColor $s $C.OffWhite

Add-TextBox $s 45 25 870 58 "Vår metodik" 30 $true $false $C.OffBlack 1 | Out-Null
Add-TextBox $s 45 78 870 35 "Tre sammanlänkade pelare – från strategi till skalning" 14 $false $true $C.MidGray 1 | Out-Null

$pillars = @(
    @{ num="01"; title="AI-Strategi"; sub="Kartläggning & riktning"; points="AI-profil och mognadsbedömning`r`nPrioriterade use cases`r`nAI-policy och governance`r`nMätbara mål och KPI:er"; mid=$false },
    @{ num="02"; title="AI-Toolbox"; sub="Implementering & verktyg"; points="Assistenter och chattbotar`r`nAutomatisering av processer`r`nInnehållsproduktion med AI`r`nAnalys och datainsikter"; mid=$true },
    @{ num="03"; title="AI-Konsult"; sub="Partnerskap & skalning"; points="Löpande strategisk rådgivning`r`nUtbildning och förankring`r`nIterativ förbättring`r`nSkalning av vad som fungerar"; mid=$false }
)

for ($i = 0; $i -lt 3; $i++) {
    $x = 45 + $i * 300
    $p = $pillars[$i]
    $bgCol = if ($p.mid) { $C.Orange } else { $C.White }
    $txtCol = if ($p.mid) { $C.White } else { $C.OffBlack }
    $bodyCol = if ($p.mid) { 14215423 } else { $C.Dark4A }
    $subCol = if ($p.mid) { 14737375 } else { $C.MidGray }  # FFD4B8 or MidGray
    $numCol = if ($p.mid) { 14737375 } else { $C.Orange }

    $card = Add-Rect $s $x 130 275 375 $bgCol
    $card.Shadow.Visible = -1; $card.Shadow.Type = 1; $card.Shadow.OffsetX = 3; $card.Shadow.OffsetY = 3; $card.Shadow.Blur = 8; $card.Shadow.ForeColor.RGB = 0; $card.Shadow.Transparency = 0.88
    if (-not $p.mid) { $card.Line.Visible = -1; $card.Line.ForeColor.RGB = $C.LightGray; $card.Line.Weight = 1 }

    Add-TextBox $s ($x+15) 148 100 30 $p.num 11 $true $false $numCol 1 | Out-Null
    Add-TextBox $s ($x+15) 188 245 55 $p.title 22 $true $false $txtCol 1 | Out-Null
    Add-TextBox $s ($x+15) 242 245 30 $p.sub 12 $false $true $subCol 1 | Out-Null

    $tb = Add-TextBox $s ($x+15) 285 245 210 $p.points 12 $false $false $bodyCol 1 1
    $tf = $tb.TextFrame
    for ($j = 1; $j -le $tf.TextRange.Paragraphs().Count; $j++) {
        $para = $tf.TextRange.Paragraphs($j)
        $para.ParagraphFormat.Bullet.Visible = -1
        $para.ParagraphFormat.Bullet.Character = 8226
        $bCol = if ($p.mid) { $C.White } else { $C.Orange }
        $para.ParagraphFormat.Bullet.Font.Color.RGB = $bCol
        $para.ParagraphFormat.SpaceAfter = 6
    }
}

Add-Notes $s "Vår metodik bygger på tre pelare som hänger ihop. Vi börjar alltid med strategin – för utan riktning blir verktyg bara leksaker. AI-Toolbox är där vi implementerar konkreta lösningar, anpassade för er verksamhet. Och AI-Konsult är det löpande partnerskapet som säkerställer att ni faktiskt skalar det som fungerar."

Write-Host "Slide 4 done"

# ============================================================
# SLIDE 5: VAD VI FÖRESLÅR FÖR REINHOLT
# ============================================================
$s = $pres.Slides.Add(5, 12)
Set-BgColor $s $C.OffWhite

Add-TextBox $s 45 25 870 58 "Vad vi föreslår för Reinholt" 30 $true $false $C.OffBlack 1 | Out-Null
Add-TextBox $s 45 78 870 35 "Tre prioriterade use cases – konkreta, lönsamma, genomförbara" 14 $false $true $C.MidGray 1 | Out-Null

$usecases = @(
    @{ num="1"; title="Admin-automatisering"; desc="Automatisera repetitiva uppgifter: orderhantering, fakturering, intern kommunikation och rapportering."; impact="Potential: 15–25% tidsbesparring per medarbetare" },
    @{ num="2"; title="Kundservice-AI"; desc="AI-driven ärendehantering som svarar på vanliga frågor, eskalerar rätt ärenden och minskar svarstider."; impact="Potential: 40–60% av ärenden hanteras automatiskt" },
    @{ num="3"; title="Resursplanering"; desc="AI-stöd för att optimera bemanning och kapacitet baserat på historiska mönster och prognoser."; impact="Potential: Bättre kapacitetsutnyttjande inför tillväxt" }
)

for ($i = 0; $i -lt 3; $i++) {
    $x = 45 + $i * 300
    $uc = $usecases[$i]

    $card = Add-Rect $s $x 130 275 375 $C.White
    $card.Shadow.Visible = -1; $card.Shadow.Type = 1; $card.Shadow.OffsetX = 3; $card.Shadow.OffsetY = 3; $card.Shadow.Blur = 8; $card.Shadow.ForeColor.RGB = 0; $card.Shadow.Transparency = 0.88
    $card.Line.Visible = -1; $card.Line.ForeColor.RGB = $C.LightGray; $card.Line.Weight = 1

    # Orange number box
    $nb = Add-Rect $s ($x+15) 148 48 48 $C.Orange
    Add-TextBox $s ($x+15) 148 48 48 $uc.num 20 $true $false $C.White 2 3 | Out-Null

    Add-TextBox $s ($x+72) 152 185 50 $uc.title 16 $true $false $C.OffBlack 1 | Out-Null
    Add-TextBox $s ($x+15) 215 245 195 $uc.desc 13 $false $false $C.Dark4A 1 1 | Out-Null

    # Impact box (yellow)
    Add-Rect $s ($x+10) 430 255 60 $C.Yellow | Out-Null
    Add-TextBox $s ($x+15) 432 245 55 $uc.impact 11 $true $false $C.DarkRed 1 1 | Out-Null
}

Add-Notes $s "Vi har tittat specifikt på er verksamhet och identifierat tre use cases som passar er fas och era prioriteringar. Admin-automatisering är oftast det snabbaste att implementera och ger direkt synliga resultat. Kundservice-AI kan transformera hur ni hanterar ökad kundvolym utan att anställa lika mycket. Och resursplanering är extra relevant nu när ni skalar med de nya storkunderna."

Write-Host "Slide 5 done"

# ============================================================
# SLIDE 6: REFERENSCASE
# ============================================================
$s = $pres.Slides.Add(6, 12)
Set-BgColor $s $C.OffWhite

Add-TextBox $s 45 25 870 58 "Så här kan det se ut i praktiken" 30 $true $false $C.OffBlack 1 | Out-Null

# Left card
$lcard = Add-Rect $s 45 100 510 405 $C.White
$lcard.Shadow.Visible = -1; $lcard.Shadow.Type = 1; $lcard.Shadow.OffsetX = 3; $lcard.Shadow.OffsetY = 3; $lcard.Shadow.Blur = 8; $lcard.Shadow.ForeColor.RGB = 0; $lcard.Shadow.Transparency = 0.88
$lcard.Line.Visible = -1; $lcard.Line.ForeColor.RGB = $C.LightGray; $lcard.Line.Weight = 1

Add-TextBox $s 60 112 480 28 "Liknande uppdrag: Distributions- och e-handelslogistik" 12 $true $false $C.Orange 1 | Out-Null

$caseText = "Situation: Medelstort logistikbolag med hög administrativ börda och ökande kundvolym.`r`n`r`nUppdrag: AI-strategi + implementation av automatiserad ärendehantering och FAQ-chattbot.`r`n`r`nResultat: 52% av kundärenden hanterades automatiskt inom 3 månader. Frigjorde 2 heltidstjänster för mer kvalificerat arbete."
Add-TextBox $s 60 150 480 335 $caseText 13 $false $false $C.Dark4A 1 1 | Out-Null

# Right: stat boxes
$stats = @(
    @{ num="3 mån"; label="till första resultat"; highlight=$false },
    @{ num="52%"; label="av ärenden automatiserat"; highlight=$true },
    @{ num="2 tjänster"; label="frigjordes för värde"; highlight=$false }
)

for ($i = 0; $i -lt 3; $i++) {
    $y = 100 + $i * 138
    $bgCol = if ($stats[$i].highlight) { $C.Orange } else { $C.White }
    $numCol = if ($stats[$i].highlight) { $C.White } else { $C.Orange }
    $lblCol = if ($stats[$i].highlight) { 14215423 } else { $C.MidGray }

    $box = Add-Rect $s 585 $y 330 118 $bgCol
    $box.Shadow.Visible = -1; $box.Shadow.Type = 1; $box.Shadow.OffsetX = 3; $box.Shadow.OffsetY = 3; $box.Shadow.Blur = 8; $box.Shadow.ForeColor.RGB = 0; $box.Shadow.Transparency = 0.88
    if (-not $stats[$i].highlight) { $box.Line.Visible = -1; $box.Line.ForeColor.RGB = $C.LightGray; $box.Line.Weight = 1 }

    Add-TextBox $s 585 ($y+5) 330 65 $stats[$i].num 30 $true $false $numCol 2 3 | Out-Null
    Add-TextBox $s 585 ($y+68) 330 40 $stats[$i].label 12 $false $false $lblCol 2 | Out-Null
}

Add-Notes $s "Vi vill inte bara presentera teorier – vi vill visa vad som faktiskt är möjligt. Det här exemplet från ett liknande uppdrag visar att vi inom 3 månader kan leverera konkreta, mätbara resultat. 52% av kundärenden automatiserade är inte en siffra vi hittade på – det är vad som faktiskt hände."

Write-Host "Slide 6 done"

# ============================================================
# SLIDE 7: TIDSPLAN
# ============================================================
$s = $pres.Slides.Add(7, 12)
Set-BgColor $s $C.OffWhite

Add-TextBox $s 45 25 870 58 "Tidsplan" 30 $true $false $C.OffBlack 1 | Out-Null
Add-TextBox $s 45 78 870 35 "Realistisk i tre faser – vi börjar smalt, skalar brett" 14 $false $true $C.MidGray 1 | Out-Null

# Timeline line
$line = $s.Shapes.AddShape(1, 120, 205, 720, 6)
$line.Fill.Solid(); $line.Fill.ForeColor.RGB = $C.LightGray
$line.Line.Visible = 0

$phases = @(
    @{ num="Fas 1"; period="Augusti 2026"; title="Kartläggning"; points="AI-mognadsanalys`r`nProcessgenomgång`r`nPrioritering av use cases`r`nBeslutsunderlag"; active=$false },
    @{ num="Fas 2"; period="Sep–Okt 2026"; title="Pilotprojekt"; points="Implementering use case #1`r`nIntern utbildning`r`nMätning och uppföljning`r`nJustering och förbättring"; active=$true },
    @{ num="Fas 3"; period="Nov 2026–"; title="Skalning"; points="Rulla ut fler use cases`r`nBygga intern AI-kompetens`r`nLöpande partnerskap`r`nKontinuerlig optimering"; active=$false }
)

for ($i = 0; $i -lt 3; $i++) {
    $x = 45 + $i * 300
    $ph = $phases[$i]
    $bgCol = if ($ph.active) { $C.Orange } else { $C.White }
    $txtCol = if ($ph.active) { $C.White } else { $C.OffBlack }
    $bodyCol = if ($ph.active) { 14215423 } else { $C.Dark4A }
    $subCol = if ($ph.active) { 14737375 } else { $C.Orange }
    $dotCol = if ($ph.active) { $C.Orange } else { $C.LightGray }

    # Dot on timeline
    $dot = $s.Shapes.AddShape(9, ($x+115), 196, 24, 24)
    $dot.Fill.Solid(); $dot.Fill.ForeColor.RGB = $dotCol; $dot.Line.Visible = 0

    $card = Add-Rect $s $x 228 275 285 $bgCol
    $card.Shadow.Visible = -1; $card.Shadow.Type = 1; $card.Shadow.OffsetX = 3; $card.Shadow.OffsetY = 3; $card.Shadow.Blur = 8; $card.Shadow.ForeColor.RGB = 0; $card.Shadow.Transparency = 0.88
    if (-not $ph.active) { $card.Line.Visible = -1; $card.Line.ForeColor.RGB = $C.LightGray; $card.Line.Weight = 1 }

    Add-TextBox $s ($x+12) 240 250 25 $ph.num 11 $true $false $subCol 1 | Out-Null
    Add-TextBox $s ($x+12) 262 250 22 $ph.period 11 $false $false (if ($ph.active) { 14737375 } else { $C.MidGray }) 1 | Out-Null
    Add-TextBox $s ($x+12) 285 250 48 $ph.title 18 $true $false $txtCol 1 | Out-Null

    $tb = Add-TextBox $s ($x+12) 338 250 165 $ph.points 12 $false $false $bodyCol 1 1
    $tf = $tb.TextFrame
    for ($j = 1; $j -le $tf.TextRange.Paragraphs().Count; $j++) {
        $para = $tf.TextRange.Paragraphs($j)
        $para.ParagraphFormat.Bullet.Visible = -1
        $para.ParagraphFormat.Bullet.Character = 8226
        $bCol = if ($ph.active) { $C.White } else { $C.Orange }
        $para.ParagraphFormat.Bullet.Font.Color.RGB = $bCol
        $para.ParagraphFormat.SpaceAfter = 4
    }
}

Add-Notes $s "Tidsplanen är medvetet realistisk. Vi börjar med kartläggning i augusti – fokuserat arbete där vi förstår er verksamhet på djupet. Sedan kör vi ett pilotprojekt under hösten – ett tydligt avgränsat use case som ger mätbara resultat. Fungerar det, skalar vi brett från november. Vi sätter aldrig igång med allt på en gång."

Write-Host "Slide 7 done"

# ============================================================
# SLIDE 8: VAD NI KAN FÖRVÄNTA ER
# ============================================================
$s = $pres.Slides.Add(8, 12)
Set-BgColor $s $C.OffWhite

Add-TextBox $s 45 25 870 58 "Vad ni kan förvänta er" 30 $true $false $C.OffBlack 1 | Out-Null

$statsTop = @(
    @{ num="3–6 mån"; label="till mätbara resultat från piloten" },
    @{ num="20–40%"; label="effektivitetsökning i automatiserade processer" },
    @{ num="100%"; label="transparens och kontroll under hela resan" }
)

for ($i = 0; $i -lt 3; $i++) {
    $x = 45 + $i * 300
    Add-Rect $s $x 100 275 145 $C.Yellow | Out-Null
    Add-TextBox $s $x 108 275 72 $statsTop[$i].num 26 $true $false $C.Orange 2 | Out-Null
    Add-TextBox $s ($x+10) 178 255 58 $statsTop[$i].label 11 $false $false $C.OffBlack 2 | Out-Null
}

# Left: what's included
Add-TextBox $s 45 268 415 35 "Vad ingår i samarbetet:" 16 $true $false $C.OffBlack 1 | Out-Null
$included = "En dedikerad AI-konsult som känner er verksamhet`r`nTydlig roadmap med prioriterade use cases`r`nImplementering och utbildning i er organisation`r`nLöpande mätning och rapportering av resultat"
$tb = Add-TextBox $s 45 308 415 190 $included 14 $false $false $C.Dark4A 1 1
$tf = $tb.TextFrame
for ($j = 1; $j -le $tf.TextRange.Paragraphs().Count; $j++) {
    $para = $tf.TextRange.Paragraphs($j)
    $para.ParagraphFormat.Bullet.Visible = -1
    $para.ParagraphFormat.Bullet.Character = 8226
    $para.ParagraphFormat.Bullet.Font.Color.RGB = $C.Orange
    $para.ParagraphFormat.SpaceAfter = 10
}

# Right: quote box (dark)
$qbox = Add-Rect $s 495 268 420 240 $C.OffBlack
$qbox.Shadow.Visible = -1; $qbox.Shadow.Type = 1; $qbox.Shadow.OffsetX = 3; $qbox.Shadow.OffsetY = 3; $qbox.Shadow.Blur = 8; $qbox.Shadow.ForeColor.RGB = 0; $qbox.Shadow.Transparency = 0.85

Add-TextBox $s 515 280 380 170 '"De flesta företag har börjat använda AI – men väldigt få har omvandlat det till verklig affärsnytta."' 15 $false $true $C.White 1 1 | Out-Null
Add-TextBox $s 515 458 380 30 "– INKA AI Labs" 12 $true $false $C.Orange 1 | Out-Null

Add-Notes $s "Konkreta förväntningar är viktiga. Vi lovar inte magi – vi lovar ett strukturerat samarbete med tydliga leverabler och mätbara mål. 3-6 månader till resultat är realistiskt för ett pilotprojekt. 20-40% effektivitetsökning är vad vi sett i liknande uppdrag. Och 100% transparens är ett löfte om att ni alltid vet vad vi gör och varför."

Write-Host "Slide 8 done"

# ============================================================
# SLIDE 9: NÄSTA STEG (dark bg)
# ============================================================
$s = $pres.Slides.Add(9, 12)
Set-BgColor $s $C.OffBlack

Add-Rect $s 0 0 11 540 $C.Orange | Out-Null
Add-TextBox $s 30 28 900 65 "Nästa steg" 36 $true $false $C.White 1 | Out-Null
Add-TextBox $s 30 90 750 35 "Vad händer om vi bestämmer oss för att gå vidare?" 15 $false $true 9474192 1 | Out-Null

$steps = @(
    "Ni bekräftar intresset och vi bokar ett uppstartsmöte (ca 2 timmar)",
    "Vi skickar en enkel behovsanalys som ni fyller i inför mötet",
    "Vi presenterar ett konkret uppdragsförslag med scope, budget och tidsplan",
    "Kartläggningsfasen startar – senast 1 augusti 2026"
)

for ($i = 0; $i -lt 4; $i++) {
    $y = 148 + $i * 88
    # Orange number square
    $nb = Add-Rect $s 30 $y 52 52 $C.Orange
    Add-TextBox $s 30 $y 52 52 "0$($i+1)" 16 $true $false $C.White 2 3 | Out-Null
    Add-TextBox $s 95 ($y+6) 830 44 $steps[$i] 15 $false $false $C.White 1 3 | Out-Null
}

Add-Notes $s "Nu vill vi veta: ser ni samma potential som vi ser? Om svaret är ja, är nästa steg väldigt konkret och utan risk. Vi startar med ett uppstartsmöte, ni fyller i en enkel behovsanalys, och vi presenterar ett förslag med tydligt scope och budget. Sedan startar kartläggningen. Ingen lång upphandling, inget abstrakt – bara konkret arbete."

Write-Host "Slide 9 done"

# ============================================================
# SLIDE 10: OM INKA AI LABS
# ============================================================
$s = $pres.Slides.Add(10, 12)
Set-BgColor $s $C.OffWhite

Add-TextBox $s 45 25 870 58 "Om INKA AI Labs" 30 $true $false $C.OffBlack 1 | Out-Null
Add-TextBox $s 45 85 530 30 "AI-byrå i Malmö – del av INKA Interactive" 16 $true $false $C.Orange 1 | Out-Null

$about = "INKA AI Labs hjälper medelstora till stora svenska företag att omsätta AI till verklig affärsnytta – inte bara testa, utan faktiskt implementera och skala.`r`n`r`nVi bygger på INKA Interactives 20+ år av digital erfarenhet kombinerat med AI-spetskompetens. Det ger ett helhetsperspektiv från strategi till implementation som de flesta AI-bolag saknar."
Add-TextBox $s 45 125 530 230 $about 14 $false $false $C.Dark4A 1 1 | Out-Null

Add-TextBox $s 45 375 180 30 "Nuvarande kunder:" 13 $true $false $C.OffBlack 1 | Out-Null
Add-TextBox $s 45 405 530 30 "HydX  ·  Forsen  ·  Jape  ·  Classicum" 13 $false $false $C.MidGray 1 | Out-Null

# Right: fact boxes
$facts = @(
    @{ num="20+"; label="års digital erfarenhet" },
    @{ num="3"; label="tjänstepelare i helhet" },
    @{ num="Malmö"; label="lokal förankring, nationell räckvidd" }
)

for ($i = 0; $i -lt 3; $i++) {
    $y = 100 + $i * 148
    $fbox = Add-Rect $s 615 $y 295 120 $C.White
    $fbox.Shadow.Visible = -1; $fbox.Shadow.Type = 1; $fbox.Shadow.OffsetX = 3; $fbox.Shadow.OffsetY = 3; $fbox.Shadow.Blur = 8; $fbox.Shadow.ForeColor.RGB = 0; $fbox.Shadow.Transparency = 0.88
    $fbox.Line.Visible = -1; $fbox.Line.ForeColor.RGB = $C.LightGray; $fbox.Line.Weight = 1
    # Orange left bar
    Add-Rect $s 615 $y 8 120 $C.Orange | Out-Null
    Add-TextBox $s 632 ($y+10) 268 55 $facts[$i].num 28 $true $false $C.Orange 1 | Out-Null
    Add-TextBox $s 632 ($y+62) 268 45 $facts[$i].label 12 $false $false $C.MidGray 1 | Out-Null
}

Add-Notes $s "INKA AI Labs är inte ett nystartat bolag. Vi bygger på 20 års digital erfarenhet från INKA Interactive, och det märks i hur vi jobbar. Vi är inte här för att sälja er det senaste AI-verktyget – vi är här för att hjälpa er bygga en varaktig AI-förmåga som skapar konkret affärsnytta."

Write-Host "Slide 10 done"

# ============================================================
# SLIDE 11: TACK / KONTAKT (dark bg)
# ============================================================
$s = $pres.Slides.Add(11, 12)
Set-BgColor $s $C.OffBlack

Add-Rect $s 0 0 11 540 $C.Orange | Out-Null
Add-TextBox $s 30 65 900 110 "Tack!" 64 $true $false $C.White 1 | Out-Null
Add-TextBox $s 30 175 750 45 "Vi ser fram emot att hjälpa Reinholt Logistik ta nästa steg." 17 $false $true 9474192 1 | Out-Null

# Contact box
$cbox = Add-Rect $s 30 255 460 235 3157872
$cbox.Line.Visible = 0

Add-TextBox $s 50 270 420 30 "Kontakta oss" 13 $true $false $C.Orange 1 | Out-Null
$contactText = "INKA AI Labs  ·  Malmö`r`ninfo@inkaailabs.se`r`ninkaailabs.se"
Add-TextBox $s 50 308 420 155 $contactText 14 $false $false $C.White 1 1 | Out-Null

Add-TextBox $s 565 310 345 150 "Från tanke till handling." 22 $true $true $C.Orange 2 3 | Out-Null

Add-Notes $s "Tack för er tid idag. Vi hoppas att ni känner er en bit längre på vägen mot att ta AI från testfas till verklig affärsnytta. Vi är redo att starta – frågan är om ni är det. Vilka frågor har ni?"

Write-Host "Slide 11 done"

# ============================================================
# SAVE
# ============================================================
Write-Host "Saving presentation..."
$pres.SaveAs($outputPath)
$pres.Close()
$pptApp.Quit()

Write-Host "✅ Presentation sparad: $outputPath"
