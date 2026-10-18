import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_document():
    doc = docx.Document()

    # Nastavení okrajů: levý 3 cm (vazba), horní/dolní/pravý 2.5 cm
    for section in doc.sections:
        section.top_margin = Inches(0.984)    # 2.5 cm
        section.bottom_margin = Inches(0.984) # 2.5 cm
        section.left_margin = Inches(1.181)   # 3.0 cm
        section.right_margin = Inches(0.984)  # 2.5 cm

    def add_p(text='', align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6, line_spacing=1.15, bold=False, italic=False, size=12):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            run = p.add_run(text)
            run.bold = bold
            run.italic = italic
            run.font.name = 'Calibri'
            run.font.size = Pt(size)
            run.font.color.rgb = RGBColor(30, 41, 59)
        return p

    # 1. Titulní strana (strana 1)
    add_p('Základní škola s rozšířenou výukou jazyků, Husova 142/44, Liberec', WD_ALIGN_PARAGRAPH.CENTER, space_after=120, size=13, bold=True)
    add_p('ZÁVĚREČNÁ PRÁCE', WD_ALIGN_PARAGRAPH.CENTER, space_after=18, size=24, bold=True)
    add_p('Škodlivý software', WD_ALIGN_PARAGRAPH.CENTER, space_after=12, size=20, bold=True)
    add_p('Druhy, způsoby šíření, následky a ochrana uživatele', WD_ALIGN_PARAGRAPH.CENTER, space_after=220, size=14, italic=True)

    p_bottom = doc.add_paragraph()
    p_bottom.paragraph_format.space_after = Pt(4)
    r1 = p_bottom.add_run('Autor: ')
    r1.bold = True
    p_bottom.add_run('Babak Kostiantyn\n')
    r2 = p_bottom.add_run('Třída: ')
    r2.bold = True
    p_bottom.add_run('9. třída\n')
    r3 = p_bottom.add_run('Školní rok: ')
    r3.bold = True
    p_bottom.add_run('2026/2027\n')
    r4 = p_bottom.add_run('Vedoucí práce: ')
    r4.bold = True
    p_bottom.add_run('doplnit jméno učitele')

    doc.add_page_break()

    # 2. Prohlášení (strana 2)
    add_p('Prohlášení', WD_ALIGN_PARAGRAPH.LEFT, space_after=24, size=16, bold=True)
    add_p('Prohlašuji, že jsem závěrečnou práci vypracoval samostatně s použitím uvedené literatury a zdrojů informací a na základě konzultací s vedoucím práce.', space_after=180, line_spacing=1.3)
    add_p('V Liberci dne ...................................                                  Podpis: ...................................')

    doc.add_page_break()

    # 3. Poděkování (strana 3)
    add_p('Poděkování', WD_ALIGN_PARAGRAPH.LEFT, space_after=24, size=16, bold=True)
    add_p('Rád bych poděkoval vedoucímu závěrečné práce za odborné vedení, cenné rady a čas věnovaný konzultacím během zpracování tohoto tématu.', space_after=24, line_spacing=1.3)

    doc.add_page_break()

    # 4. Obsah (strana 4)
    add_p('Obsah (pracovní podoba pro říjnovou odevzdávku)', WD_ALIGN_PARAGRAPH.LEFT, space_after=18, size=16, bold=True)
    add_p('Úvod ................................................................................................................................................ 5')
    add_p('1. Co je malware ............................................................................................................................... 6')
    add_p('   1.1 Původ pojmu a definice ........................................................................................................ 6')
    add_p('   1.2 Rozdíl mezi malwarem a počítačovým virem .......................................................................... 6')
    add_p('   1.3 Cíle škodlivého softwaru ....................................................................................................... 7')
    add_p('   1.4 Význam kybernetické bezpečnosti pro běžného uživatele .................................................... 7')
    add_p('2. Koncept praktické části (metodika a plán) .................................................................................... 8')
    add_p('   2.1 Cíl a zaměření praktické části ............................................................................................. 8')
    add_p('   2.2 Komparativní analýza typů škodlivého kódu .......................................................................... 8')
    add_p('   2.3 Bezpečná modelová architektura trojského koně .................................................................... 8')
    add_p('   2.4 Zpracování a vyhodnocení bezpečnostních statistik ............................................................ 9')
    add_p('   2.5 Interaktivní webová prezentace projektu (GitHub Pages) ..................................................... 9')
    add_p('Seznam použitých zdrojů ................................................................................................................... 10')

    doc.add_page_break()

    # 5. Úvod (strana 5)
    add_p('Úvod', WD_ALIGN_PARAGRAPH.LEFT, space_after=18, size=16, bold=True)
    add_p('Téma škodlivého softwaru bylo vybráno z důvodu jeho vysoké aktuálnosti v dnešní digitální době. Většina lidí denně používá počítače, chytré telefony a internet pro komunikaci, školní úkoly i online platby. Spolu s rostoucím využitím technologií se však neustále zvyšuje také počet kybernetických hrozeb, které mohou ohrozit osobní údaje, soukromí nebo finanční prostředky uživatelů.', space_after=12, line_spacing=1.25)
    add_p('Cílem této práce je srozumitelným způsobem vysvětlit, co je to malware, jaké jsou jeho hlavní typy a jaké techniky útočníci nejčastěji používají k jeho šíření. Práce se zaměřuje také na možnosti účinné prevence a ochrany běžného uživatele.', space_after=12, line_spacing=1.25)
    add_p('Práce je rozdělena na teoretickou a praktickou část. V teoretické části budou podrobně popsány základní kategorie škodlivého softwaru, historie vybraných útoků a doporučená bezpečnostní pravidla. Praktická část se zaměří na vytvoření bezpečné modelové architektury trojského koně (formou logického schématu bez škodlivého kódu), porovnání vlastností jednotlivých hrozeb a vyhodnocení aktuálních statistik z uznávaných bezpečnostních reportů.', space_after=12, line_spacing=1.25)

    doc.add_page_break()

    # 6. Teoretická část - Kapitola 1 (strana 6-7)
    add_p('1. Co je malware', WD_ALIGN_PARAGRAPH.LEFT, space_after=18, size=16, bold=True)

    add_p('1.1 Původ pojmu a definice', WD_ALIGN_PARAGRAPH.LEFT, space_after=12, size=14, bold=True)
    add_p('Slovo malware vzniklo spojením dvou anglických slov: „malicious“ (což znamená škodlivý nebo zlomyslný) a „software“ (programové vybavení počítače). V českém jazyce se pro tento pojem běžně používá označení škodlivý software nebo škodlivý kód.', space_after=12, line_spacing=1.25)
    add_p('V odborné literatuře i mezinárodních bezpečnostních standardech (například podle amerického institutu NIST) je malware definován jako jakýkoliv program, skript nebo kód, který byl vytvořen se záměrem poškodit počítačový systém, narušit jeho běžný provoz, získat neoprávněný přístup k datům nebo sledovat činnost uživatele bez jeho vědomí a souhlasu.', space_after=12, line_spacing=1.25)

    add_p('1.2 Rozdíl mezi malwarem a počítačovým virem', WD_ALIGN_PARAGRAPH.LEFT, space_after=12, size=14, bold=True)
    add_p('Mezi běžnými uživateli se velmi často stává, že jakýkoliv problém s počítačem nebo jakýkoliv podezřelý program označují slovem „virus“. Z technického hlediska je to však nepřesné.', space_after=12, line_spacing=1.25)
    add_p('Malware je nadřazený, obecný pojem pro celou rodinu škodlivých programů. Počítačový virus je pouze jedním z konkrétních typů malware. Hlavní vlastností skutečného počítačového viru je schopnost replikace – to znamená, že se dokáže sám připojit k jinému spustitelnému souboru nebo dokumentu (tzv. hostiteli) a šířit se dál, když uživatel tento soubor otevře nebo přenese.', space_after=12, line_spacing=1.25)
    add_p('Mnoho moderních hrozeb se však chová úplně jinak. Například trojský kůň se sám nekopíruje, ale maskuje se za užitečnou aplikaci (například hru nebo falešný aktualizační balíček). Ransomware zase data zašifruje a požaduje výkupné. Proto je správné používat souhrnné označení malware, které zahrnuje všechny tyto kategorie.', space_after=12, line_spacing=1.25)

    add_p('1.3 Cíle škodlivého softwaru', WD_ALIGN_PARAGRAPH.LEFT, space_after=12, size=14, bold=True)
    add_p('V počátcích počítačové historie (v 80. a 90. letech 20. století) byly škodlivé programy často vytvářeny spíše jako experiment, demonstrace schopností programátora nebo žert. Často pouze zobrazily zprávu na obrazovce nebo přehrály zvuk.', space_after=12, line_spacing=1.25)
    add_p('Dnes je situace zcela odlišná. Většina škodlivého softwaru je vyvíjena s jasným cílem:', space_after=6, line_spacing=1.25)
    add_p('• Finanční zisk: krádež přístupů k bankovním účtům, vydírání pomocí zašifrovaných souborů (ransomware) nebo tajné zneužití výkonu počítače k těžbě kryptoměn.', space_after=4, line_spacing=1.2)
    add_p('• Krádež osobních údajů a identity: získání hesel k sociálním sítím, e-mailům nebo osobních dokumentů.', space_after=4, line_spacing=1.2)
    add_p('• Vytvoření sítě ovládaných počítačů (botnet): napadené zařízení může sloužit k posílání spamu nebo provádění rozsáhlých útoků na servery (DDoS).', space_after=4, line_spacing=1.2)
    add_p('• Špionáž a sledování: tajné odposlouchávání mikrofonu, webkamery nebo zaznamenávání stisknutých kláves (keylogging).', space_after=12, line_spacing=1.2)

    add_p('1.4 Význam kybernetické bezpečnosti pro běžného uživatele', WD_ALIGN_PARAGRAPH.LEFT, space_after=12, size=14, bold=True)
    add_p('Mnoho lidí se domnívá, že pro hackery nejsou zajímaví, protože „nemají co skrývat“. Útočníci však v drtivé většině případů neútočí na konkrétního jednotlivce ručně, ale využívají automatické skenery a hromadné e-maily (phishing). Každý nezabezpečený počítač nebo telefon může být zneužit jako odrazový můstek pro další útoky. Znalost základních principů malware a způsobů ochrany je proto dnes stejně důležitá jako základní pravidla bezpečnosti v reálném světě.', space_after=18, line_spacing=1.25)

    doc.add_page_break()

    # 7. Koncept praktické části (strana 8-9)
    add_p('2. Koncept praktické části (návrh metodiky)', WD_ALIGN_PARAGRAPH.LEFT, space_after=18, size=16, bold=True)
    add_p('Pro splnění požadavků praktické části závěrečné práce byl zvolen analyticko-modelový postup. Praktická část nebude obsahovat tvorbu skutečného škodlivého programu, aby byl dodržen bezpečnostní a etický rámec školní práce. Místo toho bude praktická část postavena na čtyřech propojených pilířích:', space_after=12, line_spacing=1.25)

    add_p('2.1 Komparativní analýza typů malware (Srovnávací matice)', WD_ALIGN_PARAGRAPH.LEFT, space_after=6, size=13, bold=True)
    add_p('Bude vytvořena přehledná srovnávací tabulka hlavních kategorií škodlivého softwaru (virus, červ, trojský kůň, ransomware, spyware, adware, rootkit). Každý typ bude porovnán podle předem daných kritérií: způsob šíření (sebereplikace vs. sociální inženýrství), primární cíl (ničení dat, výkupné, špehování), míra rizika pro běžného uživatele a nejúčinnější metoda detekce.', space_after=12, line_spacing=1.25)

    add_p('2.2 Bezpečná modelová architektura trojského koně', WD_ALIGN_PARAGRAPH.LEFT, space_after=6, size=13, bold=True)
    add_p('Praktická část podrobně rozpracuje modelové logické schéma životního cyklu typického trojského koně v 5 fázích: 1. Návnada (vstupní soubor zamaskovaný za legitimní aplikaci), 2. Spuštění uživatelem, 3. Navázání skryté komunikace s řídicím serverem (C2), 4. Aktivace škodlivé funkce (payload), 5. Možnosti záchytu bezpečnostním softwarem v jednotlivých fázích. Schéma názorně ukáže, ve kterém bodě může uživatel útok včas zastavit.', space_after=12, line_spacing=1.25)

    add_p('2.3 Zpracování a vyhodnocení bezpečnostních statistik', WD_ALIGN_PARAGRAPH.LEFT, space_after=6, size=13, bold=True)
    add_p('Z oficiálních a mezinárodně uznávaných zpráv o kybernetických hrozbách (zejména ENISA Threat Landscape a Verizon Data Breach Investigations Report) budou extrahována aktuální data. Ta budou zpracována do srozumitelných grafů – například zastoupení jednotlivých vektorů útoku (phishing, zranitelnosti systému, kompromitované přihlašovací údaje) a jejich trend v posledních letech.', space_after=12, line_spacing=1.25)

    add_p('2.4 Vzdělávací webová stránka projektu (GitHub Pages)', WD_ALIGN_PARAGRAPH.LEFT, space_after=6, size=13, bold=True)
    add_p('Jako praktický výstup a moderní formu prezentace práce autor vytvoří webovou stránku hostovanou na GitHub Pages. Web bude obsahovat interaktivní verzi schémat, možnost proklikávat jednotlivé fáze útoku, interaktivní srovnávač hrozeb a možnost stažení kompletní dokumentace pro učitele a hodnotící komisi.', space_after=18, line_spacing=1.25)

    doc.add_page_break()

    # 8. Seznam zdrojů (strana 10)
    add_p('Seznam použitých zdrojů', WD_ALIGN_PARAGRAPH.LEFT, space_after=18, size=16, bold=True)
    add_p('[1] SOUČEK, Michal. Úvod do kybernetické bezpečnosti. Praha: Grada Publishing, 2021. ISBN 978-80-271-2940-0.')
    add_p('[2] ENISA (Evropská agentura pro kybernetickou bezpečnost). ENISA Threat Landscape 2024/2025. Dostupné online z: https://www.enisa.europa.eu/publications/enisa-threat-landscape-2024')
    add_p('[3] VERIZON. 2024–2025 Data Breach Investigations Report (DBIR). Dostupné online z: https://www.verizon.com/business/resources/reports/dbir/')
    add_p('[4] NIST (National Institute of Standards and Technology). Special Publication 800-83: Guide to Malware Incident Prevention and Handling for Desktops and Laptops. Dostupné z: https://csrc.nist.gov')
    add_p('[5] NÚKIB (Národní úřad pro kybernetickou a informační bezpečnost). Zpráva o stavu kybernetické bezpečnosti České republiky. Dostupné online z: https://nukib.gov.cz')

    output_path = 'documents/Zaverecna_prace_Skodlivy_software_Babak.docx'
    doc.save(output_path)
    print(f'Document successfully created: {output_path}')

if __name__ == '__main__':
    create_document()
