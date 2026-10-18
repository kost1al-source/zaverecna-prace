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
    add_p('1. Co je malware (teoretická část) ..................................................................................................... 6')
    add_p('   1.1 Původ pojmu a definice ........................................................................................................ 6')
    add_p('   1.2 Rozdíl mezi malwarem a počítačovým virem .......................................................................... 6')
    add_p('   1.3 Cíle škodlivého softwaru ....................................................................................................... 7')
    add_p('   1.4 Význam kybernetické bezpečnosti pro běžného uživatele .................................................... 7')
    add_p('2. Koncept praktické části a metodika ............................................................................................. 8')
    add_p('   2.1 Vlastní šetření: Test rozpoznávání phishingových zpráv u žáků ........................................... 8')
    add_p('   2.2 Zpracování dat a samostatná kapitola Výsledky ................................................................. 8')
    add_p('   2.3 Diskuze a srovnání zjištění se zprávami ENISA a Verizon ................................................... 9')
    add_p('   2.4 Modelová architektura trojského koně a prezentace na GitHub Pages .............................. 9')
    add_p('Seznam použitých zdrojů ................................................................................................................... 10')

    doc.add_page_break()

    # 5. Úvod (strana 5)
    add_p('Úvod', WD_ALIGN_PARAGRAPH.LEFT, space_after=18, size=16, bold=True)
    add_p('Téma škodlivého softwaru bylo vybráno z důvodu jeho vysoké aktuálnosti v dnešní digitální době. Většina lidí denně používá počítače, chytré telefony a internet pro komunikaci, školní úkoly i online platby. Spolu s rostoucím využitím technologií se však neustále zvyšuje také počet kybernetických hrozeb, které mohou ohrozit osobní údaje, soukromí nebo finanční prostředky uživatelů.', space_after=12, line_spacing=1.25)
    add_p('Cílem této práce je srozumitelným způsobem vysvětlit, co je to malware, jaké jsou jeho hlavní typy a jaké techniky útočníci nejčastěji používají k jeho šíření. Práce se zaměřuje také na možnosti účinné prevence a ochrany běžného uživatele.', space_after=12, line_spacing=1.25)
    add_p('Práce je rozdělena na teoretickou a praktickou část. V teoretické části budou podrobně popsány základní kategorie škodlivého softwaru a doporučená bezpečnostní pravidla. Praktická část přinese vlastní šetření mezi žáky školy zaměřené na schopnost rozpoznat podvodný phishingový e-mail, samostatné vyhodnocení výsledků a v diskuzi jejich porovnání se statistikami renomovaných agentur (ENISA, Verizon). Součástí je rovněž bezpečný model fungování trojského koně a interaktivní webová prezentace.', space_after=12, line_spacing=1.25)

    doc.add_page_break()

    # 6. Teoretická část - Kapitola 1 (strana 6-7)
    add_p('1. Co je malware', WD_ALIGN_PARAGRAPH.LEFT, space_after=18, size=16, bold=True)

    add_p('1.1 Původ pojmu a definice', WD_ALIGN_PARAGRAPH.LEFT, space_after=12, size=14, bold=True)
    add_p('Slovo malware vzniklo spojením dvou anglických slov: „malicious“ (což znamená škodlivý nebo zlomyslný) a „software“ (programové vybavení počítače). V českém jazyce se pro tento pojem běžně používá označení škodlivý software nebo škodlivý kód.', space_after=12, line_spacing=1.25)
    add_p('V odborné literatuře i mezinárodních bezpečnostních standardech (například podle amerického institutu NIST či českého NÚKIB) je malware definován jako jakýkoliv program, skript nebo kód, který byl vytvořen se záměrem poškodit počítačový systém, narušit jeho běžný provoz, získat neoprávněný přístup k datům nebo sledovat činnost uživatele bez jeho vědomí a souhlasu.', space_after=12, line_spacing=1.25)

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
    add_p('2. Koncept praktické části (návrh metodiky a vlastního šetření)', WD_ALIGN_PARAGRAPH.LEFT, space_after=18, size=16, bold=True)
    add_p('Na základě doporučení vedoucího práce je praktická část koncipována tak, aby přinesla vlastní originální zjištění a nebyla pouhým přepisem cizích statistik. Praktická část se skládá ze čtyř na sebe navazujících kroků:', space_after=12, line_spacing=1.25)

    add_p('2.1 Vlastní šetření: Test rozpoznávání phishingových zpráv', WD_ALIGN_PARAGRAPH.LEFT, space_after=6, size=13, bold=True)
    add_p('Hlavním jádrem praktické části bude anonymní praktický test určený pro spolužáky z 8. a 9. tříd naší školy (cílová skupina cca 20–30 respondentů). Test nebude teoretický kvíz, ale bude obsahovat 4 konkrétní modelové ukázky e-mailů a zpráv: 1. Skutečný legitimní e-mail ze školního systému, 2. Nápadně podvodný e-mail se špatnou češtinou a slibem výhry, 3. Velmi zdařilý podvod napodobující banku nebo herní platformu (Steam/Discord) s podvrženým odkazem, 4. Zpráva obsahující nebezpečnou přílohu (falešná faktura). Úkolem žáků bude rozhodnout o bezpečnosti zpráv a označit podezřelé znaky.', space_after=12, line_spacing=1.25)

    add_p('2.2 Zpracování výsledků (Samostatná kapitola Výsledky)', WD_ALIGN_PARAGRAPH.LEFT, space_after=6, size=13, bold=True)
    add_p('V samostatné kapitole Výsledky budou odpovědi žáků statisticky zpracovány do přehledných tabulek a grafů. Zjišťovat se bude zejména úspěšnost odhalení podvodu, nejčastější chyby (na jaký typ návnady žáci nejčastěji naletí) a vliv věku či zájmu o technologie na úspěšnost řešení.', space_after=12, line_spacing=1.25)

    add_p('2.3 Diskuze a srovnání se světovými daty (Samostatná kapitola Diskuze)', WD_ALIGN_PARAGRAPH.LEFT, space_after=6, size=13, bold=True)
    add_p('V kapitole Diskuze budou zjištěné výsledky žáků konfrontovány s reálnými údaji z reportů ENISA Threat Landscape a Verizon DBIR (např. údaj, že phishing stojí za 44 % všech úspěšných útoků). Tato diskuze zhodnotí zranitelnost teenagerů na internetu a navrhne konkrétní preventivní opatření pro školu.', space_after=12, line_spacing=1.25)

    add_p('2.4 Modelové schéma trojského koně a prezentace na GitHub Pages', WD_ALIGN_PARAGRAPH.LEFT, space_after=6, size=13, bold=True)
    add_p('Jako doprovodný technický výstup bude vytvořeno bezpečné modelové schéma životního cyklu trojského koně v 5 fázích a interaktivní webová prezentace na GitHub Pages (https://kost1al-source.github.io/zaverecna-prace/), kde budou shrnuty jak teoretické poznatky, tak výsledky šetření.', space_after=18, line_spacing=1.25)

    doc.add_page_break()

    # 8. Seznam zdrojů (strana 10)
    add_p('Seznam použitých zdrojů', WD_ALIGN_PARAGRAPH.LEFT, space_after=18, size=16, bold=True)
    add_p('[1] SOUČEK, Michal. Úvod do kybernetické bezpečnosti. Praha: Grada Publishing, 2021. 168 s. ISBN 978-80-271-2940-0. (Česká odborná kniha)')
    add_p('[2] NÚKIB (Národní úřad pro kybernetickou a informační bezpečnost). Zpráva o stavu kybernetické bezpečnosti České republiky za rok 2024. Brno: NÚKIB, 2025. Dostupné online z: https://nukib.gov.cz')
    add_p('[3] NÚKIB. Vzdělávací portál a metodika prevence sociálního inženýrství. Dostupné z: https://osveta.nukib.gov.cz')
    add_p('[4] ENISA (Agentura Evropské unie pro kybernetickou bezpečnost). ENISA Threat Landscape 2024/2025: Cyber Threat Analysis. Heraklion: ENISA, 2024. Dostupné z: https://www.enisa.europa.eu')
    add_p('[5] VERIZON. 2024–2025 Data Breach Investigations Report (DBIR). Verizon Business, 2025. Dostupné z: https://www.verizon.com/business/resources/reports/dbir/')
    add_p('[6] NIST (National Institute of Standards and Technology). Special Publication 800-83 Rev. 1: Guide to Malware Incident Prevention and Handling. Gaithersburg: NIST, 2020. Dostupné z: https://csrc.nist.gov')

    output_path = 'documents/Zaverecna_prace_Skodlivy_software_Babak.docx'
    doc.save(output_path)
    print(f'Document successfully created: {output_path}')

if __name__ == '__main__':
    create_document()
