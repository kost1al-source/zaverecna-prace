import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def build():
    # Načtení oficiální školní šablony
    template_path = 'zavareckapriklady/1_VZOR_TITULNÍ STRANA.docx'
    doc = docx.Document(template_path)

    # 1. Vyplnění oficiální titulní strany
    doc.paragraphs[9].text = 'Škodlivý software'
    doc.paragraphs[9].style = doc.styles['Nazev']
    
    # Přidáme podnázev
    p_sub = doc.paragraphs[10]
    p_sub.text = 'Druhy, způsoby šíření, následky a ochrana uživatele'
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if p_sub.runs:
        p_sub.runs[0].font.name = 'Times New Roman'
        p_sub.runs[0].font.size = Pt(13)
        p_sub.runs[0].font.italic = True

    doc.paragraphs[11].text = 'Babak Kostiantyn'
    doc.paragraphs[11].style = doc.styles['Jmeno']

    doc.paragraphs[12].text = 'Třída: 9. třída'
    doc.paragraphs[14].text = 'Vedoucí závěrečné práce: '
    doc.paragraphs[15].text = 'Školní rok 2026/2027'

    # 2. Prohlášení a Poděkování
    doc.paragraphs[35].text = 'Prohlášení:'
    doc.paragraphs[36].text = 'Prohlašuji, že jsem závěrečnou práci vypracoval samostatně s použitím uvedené literatury a zdrojů informací a na základě konzultací s vedoucím práce.'
    doc.paragraphs[37].text = 'V Liberci dne ...................................                                                            Podpis: ...................................'

    doc.paragraphs[38].text = 'Poděkování:'
    doc.paragraphs[39].text = 'Rád bych poděkoval vedoucímu závěrečné práce za odborné vedení, cenné rady a čas věnovaný konzultacím během zpracování tohoto tématu.'

    # Odstraníme zbývající prázdné odstavce šablony na konci (pokud nějaké jsou)
    while len(doc.paragraphs) > 40:
        p = doc.paragraphs[-1]._element
        p.getparent().remove(p)

    # Pomocná funkce pro vkládání nových odstavců
    def add_p(text='', style='Normal', space_after=6, line_spacing=1.3, align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=False, italic=False, size=12):
        p = doc.add_paragraph()
        if style in doc.styles:
            p.style = doc.styles[style]
        p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(size)
            run.bold = bold
            run.italic = italic
            run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_heading(text, level=1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        if level == 1:
            p.paragraph_format.space_before = Pt(18)
            p.paragraph_format.space_after = Pt(12)
            run.font.size = Pt(16)
        elif level == 2:
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(8)
            run.font.size = Pt(13)
        return p

    # 3. Obsah práce (Strana 4)
    doc.add_page_break()
    add_heading('Obsah (pracovní podoba pro říjnové odevzdání)', level=1)
    
    toc_items = [
        ('Úvod', '5'),
        ('1. Co je malware (teoretická část)', '6'),
        ('   1.1 Původ pojmu a definice', '6'),
        ('   1.2 Rozdíl mezi malwarem a počítačovým virem', '6'),
        ('   1.3 Cíle škodlivého softwaru', '7'),
        ('   1.4 Význam kybernetické bezpečnosti pro běžného uživatele', '7'),
        ('2. Koncept praktické části a metodika', '8'),
        ('   2.1 Vlastní šetření: Test rozpoznávání phishingových zpráv u žáků', '8'),
        ('   2.2 Zpracování dat a samostatná kapitola Výsledky', '8'),
        ('   2.3 Diskuze a srovnání zjištění se zprávami ENISA a Verizon', '9'),
        ('   2.4 Modelová architektura trojského koně a prezentace na GitHub Pages', '9'),
        ('Seznam použitých zdrojů', '10')
    ]
    for title, page in toc_items:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        dots = '.' * (85 - len(title) - len(page))
        r1 = p.add_run(title + ' ' + dots + ' ')
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(11)
        r2 = p.add_run(page)
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(11)
        r2.bold = True

    # 4. Úvod (Strana 5)
    doc.add_page_break()
    add_heading('Úvod', level=1)
    add_p('Téma škodlivého softwaru bylo vybráno z důvodu jeho vysoké aktuálnosti v dnešní digitální době. Většina lidí denně používá počítače, chytré telefony a internet pro komunikaci, školní úkoly i online platby. Spolu s rostoucím využitím technologií se však neustále zvyšuje také počet kybernetických hrozeb, které mohou ohrozit osobní údaje, soukromí nebo finanční prostředky uživatelů.')
    add_p('Cílem této práce je srozumitelným způsobem vysvětlit, co je to malware, jaké jsou jeho hlavní typy a jaké techniky útočníci nejčastěji používají k jeho šíření. Práce se zaměřuje také na možnosti účinné prevence a ochrany běžného uživatele.')
    add_p('Práce je rozdělena na teoretickou a praktickou část. V teoretické části budou podrobně popsány základní kategorie škodlivého softwaru a doporučená bezpečnostní pravidla. Praktická část přinese vlastní šetření mezi žáky školy zaměřené na schopnost rozpoznat podvodný phishingový e-mail, samostatné vyhodnocení výsledků a v diskuzi jejich porovnání se statistikami renomovaných agentur (ENISA, Verizon). Součástí je rovněž bezpečný model fungování trojského koně a interaktivní webová prezentace.')

    # 5. Teoretická část (Strana 6-7)
    doc.add_page_break()
    add_heading('1. Co je malware', level=1)

    add_heading('1.1 Původ pojmu a definice', level=2)
    add_p('Slovo malware vzniklo spojením dvou anglických slov: „malicious“ (což znamená škodlivý nebo zlomyslný) a „software“ (programové vybavení počítače). V českém jazyce se pro tento pojem běžně používá označení škodlivý software nebo škodlivý kód.')
    add_p('V odborné literatuře i mezinárodních bezpečnostních standardech (například podle amerického institutu NIST či českého NÚKIB) je malware definován jako jakýkoliv program, skript nebo kód, který byl vytvořen se záměrem poškodit počítačový systém, narušit jeho běžný provoz, získat neoprávněný přístup k datům nebo sledovat činnost uživatele bez jeho vědomí a souhlasu.')

    add_heading('1.2 Rozdíl mezi malwarem a počítačovým virem', level=2)
    add_p('Mezi běžnými uživateli se velmi často stává, že jakýkoliv problém s počítačem nebo jakýkoliv podezřelý program označují slovem „virus“. Z technického hlediska je to však nepřesné.')
    add_p('Malware je nadřazený, obecný pojem pro celou rodinu škodlivých programů. Počítačový virus je pouze jedním z konkrétních typů malware. Hlavní vlastností skutečného počítačového viru je schopnost replikace – to znamená, že se dokáže sám připojit k jinému spustitelnému souboru nebo dokumentu (tzv. hostiteli) a šířit se dál, když uživatel tento soubor otevře nebo přenese.')
    add_p('Mnoho moderních hrozeb se však chová úplně jinak. Například trojský kůň se sám nekopíruje, ale maskuje se za užitečnou aplikaci (například hru nebo falešný aktualizační balíček). Ransomware zase data zašifruje a požaduje výkupné. Proto je správné používat souhrnné označení malware, které zahrnuje všechny tyto kategorie.')

    add_heading('1.3 Cíle škodlivého softwaru', level=2)
    add_p('V počátcích počítačové historie (v 80. a 90. letech 20. století) byly škodlivé programy často vytvářeny spíše jako experiment, demonstrace schopností programátora nebo žert. Často pouze zobrazily zprávu na obrazovce nebo přehrály zvuk. Dnes je většina škodlivého softwaru vyvíjena s jasným cílem:')
    add_p('• Finanční zisk: krádež přístupů k bankovním účtům, vydírání pomocí zašifrovaných souborů (ransomware) nebo tajné zneužití výkonu počítače k těžbě kryptoměn.', space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('• Krádež osobních údajů a identity: získání hesel k sociálním sítím, e-mailům nebo osobních dokumentů.', space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('• Vytvoření sítě ovládaných počítačů (botnet): napadené zařízení může sloužit k posílání spamu nebo provádění rozsáhlých útoků na servery (DDoS).', space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('• Špionáž a sledování: tajné odposlouchávání mikrofonu, webkamery nebo zaznamenávání stisknutých kláves (keylogging).', space_after=8, align=WD_ALIGN_PARAGRAPH.LEFT)

    add_heading('1.4 Význam kybernetické bezpečnosti pro běžného uživatele', level=2)
    add_p('Mnoho lidí se domnívá, že pro hackery nejsou zajímaví, protože „nemají co skrývat“. Útočníci však v drtivé většině případů neútočí na konkrétního jednotlivce ručně, ale využívají automatické skenery a hromadné e-maily (phishing). Každý nezabezpečený počítač nebo telefon může být zneužit jako odrazový můstek pro další útoky. Znalost základních principů malware a způsobů ochrany je proto dnes stejně důležitá jako základní pravidla bezpečnosti v reálném světě.')

    # 6. Praktická část (Strana 8-9)
    doc.add_page_break()
    add_heading('2. Koncept praktické části a metodika', level=1)
    add_p('Na základě doporučení vedoucího práce je praktická část koncipována tak, aby přinesla vlastní originální zjištění a nebyla pouhým přepisem cizích statistik. Praktická část se skládá ze čtyř na sebe navazujících kroků:')

    add_heading('2.1 Vlastní šetření: Test rozpoznávání phishingových zpráv', level=2)
    add_p('Hlavním jádrem praktické části bude anonymní praktický test určený pro spolužáky z 8. a 9. tříd naší školy (cílová skupina cca 20–30 respondentů). Test nebude teoretický kvíz, ale bude obsahovat 4 konkrétní modelové ukázky e-mailů a zpráv:')
    add_p('1. Skutečný legitimní e-mail (např. běžné oznámení ze školního systému Bakaláři).', space_after=3, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('2. Nápadně podvodný e-mail (špatná čeština, podezřelá adresa, slib výhry).', space_after=3, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('3. Velmi zdařilý podvod (napodobení zprávy od banky, pošty nebo herní platformy Steam/Discord s podvrženým odkazem).', space_after=3, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('4. Zpráva obsahující nebezpečnou přílohu (falešná faktura s dvojitou příponou).', space_after=6, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('Úkolem žáků bude u každé zprávy rozhodnout, zda je bezpečná, označit podezřelé znaky (chybná URL adresa, nátlak na čas, žádost o heslo) a uvést, jak by na zprávu zareagovali.')

    add_heading('2.2 Zpracování výsledků (Samostatná kapitola Výsledky)', level=2)
    add_p('V samostatné kapitole Výsledky budou odpovědi žáků statisticky zpracovány do přehledných tabulek a grafů. Zjišťovat se bude zejména celková úspěšnost odhalení jednotlivých typů podvodů, nejčastější chyby (která návnada byla pro žáky nejvíce matoucí) a zda hraje roli věk žáků nebo jejich předchozí zájem o informační technologie.')

    add_heading('2.3 Diskuze a srovnání se světovými daty (Samostatná kapitola Diskuze)', level=2)
    add_p('V kapitole Diskuze budou zjištěné výsledky žáků naší školy konfrontovány s reálnými údaji z mezinárodních reportů ENISA Threat Landscape a Verizon DBIR (např. údaj, že phishing stojí za 44 % všech úspěšných útoků na uživatele). Tato diskuze zhodnotí zranitelnost teenagerů na internetu a navrhne konkrétní preventivní opatření pro školní výuku informatiky.')

    add_heading('2.4 Modelová architektura trojského koně a prezentace na GitHub Pages', level=2)
    add_p('Praktickou část doplní bezpečné modelové schéma životního cyklu trojského koně v 5 fázích (vstupní soubor, spuštění, C2 komunikace, aktivace payloadu a body obrany bez nebezpečného kódu) a vzdělávací webová stránka projektu běžící na GitHub Pages (https://kost1al-source.github.io/zaverecna-prace/), kde budou shrnuty jak teoretické poznatky, tak interaktivní přehled výsledků šetření.')

    # 7. Zdroje (Strana 10)
    doc.add_page_break()
    add_heading('Seznam použitých zdrojů', level=1)
    
    sources = [
        '[1] SOUČEK, Michal. Úvod do kybernetické bezpečnosti. Praha: Grada Publishing, 2021. 168 s. ISBN 978-80-271-2940-0. (Česká odborná kniha)',
        '[2] NÚKIB (Národní úřad pro kybernetickou a informační bezpečnost). Zpráva o stavu kybernetické bezpečnosti České republiky za rok 2024. Brno: NÚKIB, 2025. Dostupné online z: https://nukib.gov.cz (Oficiální český zdroj)',
        '[3] NÚKIB. Vzdělávací portál a metodika prevence sociálního inženýrství. Dostupné z: https://osveta.nukib.gov.cz',
        '[4] ENISA (Agentura Evropské unie pro kybernetickou bezpečnost). ENISA Threat Landscape 2024/2025: Cyber Threat Analysis. Heraklion: ENISA, 2024. Dostupné z: https://www.enisa.europa.eu',
        '[5] VERIZON. 2024–2025 Data Breach Investigations Report (DBIR). Verizon Business, 2025. Dostupné z: https://www.verizon.com/business/resources/reports/dbir/',
        '[6] NIST (National Institute of Standards and Technology). Special Publication 800-83 Rev. 1: Guide to Malware Incident Prevention and Handling. Gaithersburg: NIST, 2020. Dostupné z: https://csrc.nist.gov'
    ]
    for src in sources:
        add_p(src, space_after=8, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.15)

    # Uložení do documents i přímo do kořene složky
    out1 = 'documents/Zaverecna_prace_Skodlivy_software_Babak.docx'
    out2 = 'Zaverecna_prace_Skodlivy_software_Babak.docx'
    doc.save(out1)
    doc.save(out2)
    print(f'Files saved successfully:\n- {out1}\n- {out2}')

if __name__ == '__main__':
    build()
