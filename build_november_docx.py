# -*- coding: utf-8 -*-
"""
Generátor kompletní listopadové verze Závěrečné práce
Téma: Škodlivý software (ZŠ Husova Liberec)
Autor: Babak Kostiantyn, 9. ročník
"""

import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def build_november():
    template_path = 'zavareckapriklady/1_VZOR_TITULNÍ STRANA.docx'
    doc = docx.Document(template_path)

    # 1. Vyplnění oficiální titulní strany
    doc.paragraphs[9].text = 'Škodlivý software'
    doc.paragraphs[9].style = doc.styles['Nazev']
    
    # Podnázev
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
    doc.paragraphs[14].text = 'Vedoucí závěrečné práce: pan učitel Konopa'
    doc.paragraphs[15].text = 'Školní rok 2026/2027'

    # 2. Prohlášení
    doc.paragraphs[35].text = 'Prohlášení:'
    doc.paragraphs[36].text = 'Prohlašuji, že jsem závěrečnou práci vypracoval samostatně s použitím uvedené literatury a zdrojů informací a na základě konzultací s vedoucím práce.'
    doc.paragraphs[37].text = 'V Liberci dne ...................................                                                            Podpis: ...................................'

    # 3. Poděkování (přesně podle přání autora)
    doc.paragraphs[38].text = 'Poděkování:'
    thanks_text = (
        'Rád bych touto cestou poděkoval panu učiteli Konopovi za odborné vedení, cenné rady a trpělivost při tvorbě této závěrečné práce.\n\n'
        'Dále děkuji společnostem Anthropic, OpenAI a Google za vývoj pokročilých technologií, a zejména modelům umělé inteligence GPT-6 Astra, Claude 4 Sonnet, Gemini 3.8 Flash a Gemini 4 Argon za technickou podporu a konzultace při programování a rešerši informací.\n\n'
        'V neposlední řadě patří mé upřímné poděkování energetickým nápojům Monster, které mi dodávaly potřebnou energii během dlouhých večerů strávených nad tímto projektem.'
    )
    doc.paragraphs[39].text = thanks_text

    # Odstraníme zbývající prázdné odstavce šablony
    while len(doc.paragraphs) > 40:
        p = doc.paragraphs[-1]._element
        p.getparent().remove(p)

    # Pomocné funkce pro vkládání formátovaného obsahu
    def add_p(text='', style='Normal', space_after=6, space_before=0, line_spacing=1.3, align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=False, italic=False, size=12):
        p = doc.add_paragraph()
        if style in doc.styles:
            p.style = doc.styles[style]
        p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.space_before = Pt(space_before)
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
            p.paragraph_format.space_after = Pt(10)
            run.font.size = Pt(16)
        elif level == 2:
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)
            run.font.size = Pt(13)
        elif level == 3:
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            run.font.size = Pt(12)
            run.italic = True
        return p

    # 4. Anotace a Klíčová slova (Povinná součást listopadového zadání)
    doc.add_page_break()
    add_heading('Anotace', level=1)
    add_p(
        'Tato závěrečná práce se zabývá problematikou škodlivého softwaru (malware), jeho jednotlivými kategoriemi, '
        'mechanismy šíření a metodami obrany běžného uživatele v digitálním prostoru. V teoretické části je definován pojem '
        'malware a vyjasněn rozdíl mezi počítačovými viry a ostatními druhy škodlivého kódu. Podrobně jsou analyzovány trojské koně, '
        'ransomware, spyware a techniky sociálního inženýrství, zejména phishing a podvržené domény. Praktická část je zaměřena '
        'na vlastní kvantitativní šetření mezi žáky 8. a 9. ročníků ZŠ Husova Liberec prostřednictvím interaktivního webového testu '
        'se čtyřmi reálnými modelovými scénáři. Práce vyhodnocuje schopnost žáků rozpoznat kybernetické hrozby a formuluje praktická '
        'doporučení pro zvýšení jejich digitální bezpečnosti.'
    )

    add_heading('Annotation', level=2)
    add_p(
        'This final thesis investigates the issue of malicious software (malware), its primary categories, distribution mechanisms, '
        'and defensive strategies for everyday users in the digital domain. The theoretical section establishes the definition of '
        'malware and differentiates computer viruses from other types of malicious code. It provides an in-depth analysis of trojan '
        'horses, ransomware, spyware, and social engineering methodologies, particularly phishing schemes and typosquatting domains. '
        'The practical component centers on original quantitative research conducted among 8th and 9th grade pupils at ZŠ Husova Liberec '
        'via an interactive web simulation test featuring four authentic scenarios. The thesis evaluates students\' ability to identify cyber '
        'threats and formulates practical recommendations to enhance their digital safety.'
    )

    add_heading('Klíčová slova', level=2)
    add_p('Škodlivý software, malware, počítačový virus, trojský kůň, ransomware, spyware, sociální inženýrství, phishing, kybernetická bezpečnost, dvoufázové ověření.')

    add_heading('Schlüsselwörter', level=2)
    add_p('Schadsoftware, Malware, Computervirus, Trojanisches Pferd, Ransomware, Spyware, Social Engineering, Phishing, Cybersicherheit, Zwei-Faktor-Authentifizierung.')

    # 5. Obsah práce
    doc.add_page_break()
    add_heading('Obsah', level=1)

    toc_items = [
        ('Úvod', '5'),
        ('1. Co je malware (teoretická část)', '6'),
        ('   1.1 Původ pojmu a definice', '6'),
        ('   1.2 Rozdíl mezi malwarem a počítačovým virem', '6'),
        ('   1.3 Cíle škodlivého softwaru', '7'),
        ('   1.4 Význam kybernetické bezpečnosti pro běžného uživatele', '7'),
        ('2. Hlavní kategorie škodlivého softwaru', '8'),
        ('   2.1 Počítačové viry a červi', '8'),
        ('   2.2 Trojské koně a jejich vnitřní struktura', '9'),
        ('   2.3 Ransomware (vyděračský software)', '10'),
        ('   2.4 Spyware, keyloggery a infostealery', '11'),
        ('   2.5 Adware a potenciálně nežádoucí programy (PUP)', '12'),
        ('3. Vektory šíření a sociální inženýrství', '13'),
        ('   3.1 Lidský faktor jako nejslabší článek bezpečnosti', '13'),
        ('   3.2 Phishing a spear-phishing', '13'),
        ('   3.3 Psychologický nátlak a manipulace', '14'),
        ('   3.4 Nebezpečné spustitelné přílohy (.pdf.exe)', '15'),
        ('   3.5 Typosquatting a podvržené domény', '15'),
        ('4. Prevence a účinná ochrana uživatele', '16'),
        ('   4.1 Aktualizace operačního systému a softwaru', '16'),
        ('   4.2 Antivirový software a vestavěná ochrana Windows', '16'),
        ('   4.3 Silná hesla a vícefaktorové ověřování (2FA/MFA)', '17'),
        ('   4.4 Pravidla bezpečného zálohování (pravidlo 3-2-1)', '18'),
        ('5. Metodika praktické části', '19'),
        ('   5.1 Cíle šetření a výzkumné předpoklady', '19'),
        ('   5.2 Charakteristika zkoumaného vzorku (žáci 8. a 9. tříd)', '19'),
        ('   5.3 Metoda sběru dat: interaktivní webový test', '20'),
        ('   5.4 Popis čtyř testovacích scénářů', '20'),
        ('   5.5 Anonymita a etika sběru dat', '21'),
        ('6. Koncept vyhodnocení výsledků a modelová architektura', '22'),
        ('7. Diskuze a srovnání se světovými daty', '23'),
        ('Závěr', '24'),
        ('Seznam použitých zdrojů', '25')
    ]

    for title, page in toc_items:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        dots_count = max(5, 82 - len(title) - len(page))
        dots = '.' * dots_count
        r1 = p.add_run(title + ' ' + dots + ' ')
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(11)
        r2 = p.add_run(page)
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(11)
        r2.bold = True

    # 6. Úvod
    doc.add_page_break()
    add_heading('Úvod', level=1)
    add_p(
        'Téma škodlivého softwaru bylo pro mou závěrečnou práci vybráno z důvodu jeho mimořádné aktuálnosti v dnešním '
        'každodenním životě. Většina lidí, a zejména dospívajících žáků základních škol, denně používá počítače, chytré '
        'telefony a internet pro komunikaci, školní práci, hraní her i online nákupy. Spolu s masivním rozvojem digitálních '
        'technologií se však neustále zvyšuje také množství a rafinovanost kybernetických hrozeb. Škodlivý kód již dávno '
        'není pouze dílem zvědavých nadšenců, ale stal se nástrojem organizovaného zločinu zaměřeného na krádeže identity, '
        'finančních prostředků a citlivých osobních údajů.'
    )
    add_p(
        'Hlavním cílem této práce je srozumitelným a přehledným způsobem objasnit, co to malware ve skutečnosti je, jaké jsou '
        'jeho hlavní kategorie a jaké konkrétní techniky útočníci využívají k oklamání uživatelů. Zároveň práce klade velký '
        'důraz na praktickou prevenci – tedy na pravidla chování, která umožní i běžnému uživateli bez hlubokých programátorských '
        'znalostí efektivně zabezpečit svá zařízení a data.'
    )
    add_p(
        'Práce je logicky rozdělena na teoretickou a praktickou část. V teoretické části podrobně rozebírám anatomii škodlivého '
        'softwaru od jeho definice přes jednotlivé typy (viry, trojské koně, ransomware, spyware) až po psychologické mechanismy '
        'sociálního inženýrství a zásady technické obrany. Praktická část přináší mé vlastní empirické šetření realizované formou '
        'originálního interaktivního testu mezi žáky 8. a 9. ročníků naší školy ZŠ Husova Liberec. V něm jsou spolužáci postaveni '
        'před čtyři autentické modelové situace a jejich reakce tvoří základ pro statistické vyhodnocení zranitelnosti mladé generace.'
    )

    # 7. Kapitola 1: Co je malware
    doc.add_page_break()
    add_heading('1. Co je malware', level=1)

    add_heading('1.1 Původ pojmu a definice', level=2)
    add_p(
        'Pojem malware vznikl jako zkratka ze dvou anglických slov: „malicious“ (škodlivý, zlomyslný) a „software“ '
        '(programové vybavení). V české odborné terminologii se pro tento koncept běžně užívá výraz škodlivý software '
        'či škodlivý kód [1].'
    )
    add_p(
        'Podle definice Národního úřadu pro kybernetickou a informační bezpečnost (NÚKIB) i mezinárodního standardu NIST SP 800-83 '
        'představuje malware jakýkoliv program, skript nebo kód navržený se záměrem narušit běžný provoz počítačového systému, '
        'poškodit soubory, získat neoprávněný přístup k datům nebo sledovat aktivitu uživatele bez jeho vědomí a výslovného souhlasu [3, 7].'
    )

    add_heading('1.2 Rozdíl mezi malwarem a počítačovým virem', level=2)
    add_p(
        'V běžné laické veřejnosti panuje častý omyl, kdy lidé jakékoliv selhání počítače nebo jakýkoliv podezřelý program označují '
        'slovem „virus“. Z odborného hlediska je však tento přístup nesprávný. Malware je nadřazená kategorie – široká rodina '
        'zahrnující mnoho různých typů programů s odlišným způsobem fungování.'
    )
    add_p(
        'Skutečný počítačový virus je pouze jedním z podtypů. Jeho klíčovou vlastností je schopnost samovolné replikace (množení) '
        'prostřednictvím připojení se k jinému legitimnímu spustitelnému souboru či dokumentu (tzv. hostiteli). Virus ke svému šíření '
        'nezbytně vyžaduje aktivitu uživatele – například spuštění infikovaného programu nebo otevření dokumentu s nakaženým makrem. '
        'Naproti tomu trojský kůň se sám nereplikuje, ale maskuje se za užitečný program, a počítačový červ se šíří počítačovou sítí '
        'zcela automaticky přes zranitelnosti v operačním systému [2].'
    )

    add_heading('1.3 Cíle škodlivého softwaru', level=2)
    add_p(
        'Zatímco v raných dobách osobních počítačů (v 80. a 90. letech 20. století) vznikaly viry převážně jako experimenty programátorů, '
        'demonstrace technických schopností nebo neškodné žerty (zobrazující texty či animace na monitoru), dnešní kyberkriminalita má '
        'zcela pragmatické motivace [1]:'
    )
    add_p('• Přímý finanční zisk: zcizení přístupových údajů k internetovému bankovnictví, vydírání pomocí zašifrovaných dat nebo tajné zneužití procesoru k těžbě kryptoměn (cryptojacking).', space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('• Krádež osobních údajů a identity: získání přihlašovacích hesel k e-mailům, sociálním sítím, herním účtům (Steam, Discord) a zneužití identity oběti k další trestné činnosti.', space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('• Zapojení do sítě botnet: ovládnutí počítače na dálku a jeho využití v síti „zombie“ zařízení k rozesílání spamu nebo vedení distribuovaných útoků na dostupnost serverů (DDoS).', space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('• Špionáž a sledování: skryté nahrávání stisknutých kláves (keylogging), odposlech mikrofonu, aktivace webkamery nebo krádež interních školních a firemních dokumentů.', space_after=8, align=WD_ALIGN_PARAGRAPH.LEFT)

    add_heading('1.4 Význam kybernetické bezpečnosti pro běžného uživatele', level=2)
    add_p(
        'Častým argumentem běžných uživatelů bývá tvrzení: „Já na počítači nemám nic cenného, hackeři o mě nemají zájem.“ '
        'To je však zásadní nepochopení moderních útoků. Útočníci většinou necílí na konkrétního žáka či studenta manuálně, '
        'ale rozesílají podvodné e-maily a skenují sítě plně automatizovanými roboty. Každý infikovaný počítač má pro útočníka '
        'hodnotu – ať už jako úložiště dat, odrazový můstek pro zametení stop nebo jako nástroj k infikování jeho kontaktů. '
        'Digitální bezpečnost proto představuje základní životní dovednost každého uživatele 21. století [2].'
    )

    # 8. Kapitola 2: Hlavní kategorie škodlivého softwaru (Nová část teorie)
    doc.add_page_break()
    add_heading('2. Hlavní kategorie škodlivého softwaru', level=1)
    add_p(
        'Abychom se mohli před škodlivým kódem účinně bránit, musíme detailně rozumět jeho základním formám, způsobům '
        'jeho aktivace a důsledkům, které může mít pro napadený systém. Následující podkapitoly přinášejí systematický '
        'rozbor nejvýznamnějších typů malware, se kterými se lze v praxi setkat.'
    )

    add_heading('2.1 Počítačové viry a červi', level=2)
    add_p(
        'Jak již bylo zmíněno, počítačový virus se váže na spustitelné soubory (nejčastěji s příponami .exe, .com, .scr) '
        'nebo na makra v kancelářských dokumentech (.docm, .xlsm). Když uživatel takový soubor spustí, kód viru se nahraje do '
        'operační paměti RAM a začne vyhledávat další dosud nenakažené soubory na pevném disku, do kterých vkládá své kopie [1].'
    )
    add_p(
        'Počítačový červ (worm) představuje ještě nebezpečnější evoluci. Červ ke svému šíření nepotřebuje hostitelský soubor '
        'ani asistenci uživatele. Je to samostatný program, který aktivně prohledává počítačovou síť, identifikuje zranitelná '
        'zařízení s nezabezpečenými porty či neaktualizovaným systémem a sám se na ně přes síť přenese a spustí. Historickým '
        'milníkem byl například červ Morris z roku 1988, který vyřadil z provozu značnou část tehdejšího internetu, nebo '
        'červ WannaCry z roku 2017, který zneužil chybu v protokolu SMB operačního systému Windows a způsobil miliardové škody [2].'
    )

    add_heading('2.2 Trojské koně a jejich vnitřní struktura', level=2)
    add_p(
        'Trojský kůň (trojan) nese svůj název podle známé lsti z antické mytologie. Je to program, který se navenek tváří '
        'jako užitečná, zábavná nebo lákavá aplikace – například crack do hry, cheaty na Minecraft, bezplatný editor videí '
        'nebo důležitá faktura v e-mailu. Jakmile jej však uživatel spustí v domnění, že získal požadovaný program, v pozadí '
        'se v tichosti aktivuje skrytá škodlivá funkcionalita [1].'
    )
    add_p(
        'Vnitřní architektura trojského koně se obvykle skládá ze tří základních komponent:'
    )
    add_p('1. Dropper / Zaváděcí vrstva: vnější obal, který oklame uživatele i jednoduché antiviry, rozbalí skrytý obsah a zajistí své automatické spouštění po startu Windows (např. zápisem do registru Run).', space_after=3, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('2. C2 klient (Command & Control): část kódu, která naváže skryté šifrované spojení se serverem útočníka na internetu a čeká na další povely.', space_after=3, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('3. Škodlivý náklad (Payload): vlastní výkonná část, která může mít podobu zadních vrátek (backdoor), modulu pro krádež hesel nebo stahovače dalšího nebezpečného kódu.', space_after=6, align=WD_ALIGN_PARAGRAPH.LEFT)

    add_heading('2.3 Ransomware (vyděračský software)', level=2)
    add_p(
        'Ransomware je v současnosti považován za jednu z nejničivějších hrozeb pro firmy, nemocnice, školy i soukromé osoby. '
        'Jeho princip spočívá v tom, že po vniknutí do počítače vyhledá všechny důležité uživatelské soubory (dokumenty, tabulky, '
        'fotografie, rodinná videa, databáze) a zašifruje je pomocí silné asymetrické kryptografie (např. kombinací algoritmů '
        'AES-256 a RSA-2048) [2].'
    )
    add_p(
        'Po dokončení šifrování změní tapetu na ploše nebo zobrazí celoobrazovkové varování s oznámením, že soubory jsou nedostupné. '
        'Útočníci požadují zaplacení výkupného v anonymní kryptoměně (Bitcoin, Monero) v řádu stovek až desítek tisíc dolarů za '
        'zaslání dešifrovacího klíče. Česká republika se stala terčem mnoha takových útoků, z nichž nejznámější bylo ochromení '
        'nemocnice v Benešově v prosinci 2019, kde musely být odloženy operace a celý provoz přešel do papírové formy [3]. '
        'Bezpečnostní experti i policie zásadně nedoporučují výkupné platit – financuje se tím další trestná činnost a neexistuje '
        'žádná záruka, že útočníci klíč skutečně pošlou.'
    )

    add_heading('2.4 Spyware, keyloggery a infostealery', level=2)
    add_p(
        'Spyware je škodlivý software zaměřený na tajné sledování uživatele a sběr informací bez jeho vědomí. Na rozdíl od ransomwaru '
        'se spyware snaží zůstat co nejdéle zcela neviditelný, aby mohl nerušeně pracovat měsíce i roky [1].'
    )
    add_p(
        'Specifickým a nebezpečným druhem je keylogger, který zaznamenává každý jednotlivý stisk klávesy na klávesnici. '
        'Tímto způsobem útočník získá přesný text přihlašovacích jmen, hesel, zpráv na chatu i čísel platebních karet. '
        'Moderní infostealery (např. malware RedLine či Racoon Stealer) jdou ještě dál – cíleně vykrádají soubory cookies '
        'a uložená hesla přímo z paměti webových prohlížečů (Google Chrome, Microsoft Edge), což útočníkovi umožňuje převzít '
        'přihlášenou relaci k účtu (např. Discord či Google) i bez znalosti samotného hesla [2].'
    )

    add_heading('2.5 Adware a potenciálně nežádoucí programy (PUP)', level=2)
    add_p(
        'Adware se zaměřuje na agresivní zobrazování nevyžádané reklamy, vyskakovacích oken v prohlížeči nebo přesměrování '
        'domovské stránky na pochybné vyhledávače. Často se šíří jako součást instalačních balíčků bezplatných programů '
        '(tzv. bundling), kde uživatel při rychlém klikání na tlačítko „Další“ nechtěně odsouhlasí instalaci doplňkových nástrojů. '
        'Ačkoliv adware obvykle přímo neničí data, drasticky zpomaluje počítač a může sloužit jako brána pro instalaci závažnějšího malware.'
    )

    # 9. Kapitola 3: Vektory šíření a sociální inženýrství (Nová část teorie)
    doc.add_page_break()
    add_heading('3. Vektory šíření a sociální inženýrství', level=1)
    add_p(
        'V minulosti museli útočníci spoléhat na programátorské chyby v operačních systémech. Dnes jsou však systémy '
        'jako Windows 11 či macOS podstatně lépe zabezpečené. Proto se těžiště útoků přesunulo k nejslabšímu článku '
        'bezpečnostního řetězce – k samotnému člověku [1].'
    )

    add_heading('3.1 Lidský faktor jako nejslabší článek bezpečnosti', level=2)
    add_p(
        'Sociální inženýrství (social engineering) je soubor psychologických technik, kterými útočník manipuluje člověka '
        'k tomu, aby dobrovolně vykonal akci, kterou by za normálních okolností neudělal – například kliknul na podvodný '
        'odkaz, spustil podezřelý soubor, sdělil své heslo nebo poslal peníze [4]. Útočník využívá přirozené lidské vlastnosti: '
        'důvěřivost, zvědavost, strach z trestu, chamtivost nebo ochotu pomoci druhému.'
    )

    add_heading('3.2 Phishing a spear-phishing', level=2)
    add_p(
        'Phishing (odvozeno od anglického slova „fishing“ – rybaření) představuje nejčastější formu kybernetického útoku na světě. '
        'Podle celosvětové zprávy Verizon DBIR stojí phishing za více než 40 % všech úspěšných bezpečnostních narušení [6]. '
        'Útočník rozešle tisíce e-mailů nebo SMS zpráv (smishing), které věrně napodobují vizuální styl známé instituce '
        '(banka, přepravní společnost jako Zásilkovna či Česká pošta, školní systém Bakaláři, sociální síť nebo herní služba Steam).'
    )
    add_p(
        'Zatímco klasický phishing je plošný a necílený, tzv. spear-phishing představuje útok ušitý na míru konkrétní oběti. '
        'Útočník si předem zjistí jméno oběti, její školu, zájmy nebo jména přátel na sociálních sítích a zprávu naformuluje '
        'tak věrohodně, že je odhalitelná jen při velmi důkladné kontrole adresy odesílatele a odkazů [2].'
    )

    add_heading('3.3 Psychologický nátlak a manipulace', level=2)
    add_p(
        'Podvodné zprávy téměř vždy pracují s prvkem umělého spěchu a paniky. Zpráva tvrdí: „Váš účet bude do 2 hodin trvale '
        'zablokován!“, „Máte neuhrazenou pokutu!“, nebo naopak láká: „Vyhrál jsi iPhone 16, zbývá 10 minut k vyzvednutí!“. '
        'Cílem tohoto nátlaku je vypnout racionální myšlení oběti. Ve stresu nebo v euforii z výhry uživatel jedná zbrkle, '
        'nevšimne si podezřelé e-mailové adresy a klikne na přiložený odkaz [4].'
    )

    add_heading('3.4 Nebezpečné spustitelné přílohy (.pdf.exe)', level=2)
    add_p(
        'Jedním z nejzákeřnějších triků útočníků je maskování přípon souborů. Ve výchozím nastavení systému Windows je totiž '
        'aktivní volba „Skrýt přípony souborů známých typů“. Útočník vytvoří spustitelný virus a pojmenuje jej '
        'například Faktura_Doplatek.pdf.exe. Operační systém zobrazí uživateli pouze název Faktura_Doplatek.pdf a přiřadí '
        'souboru ikonu PDF dokumentu. Uživatel se domnívá, že otevírá běžný textový dokument, ve skutečnosti však spustí binární '
        'program s plnými právy, který infikuje jeho počítač [1].'
    )

    add_heading('3.5 Typosquatting a podvržené domény', level=2)
    add_p(
        'Typosquatting je metoda založená na drobných překlepech v internetových doménách. Útočníci si zaregistrují doménu, '
        'která se liší pouze jedním písmenem od oficiální stránky – například misto „steamcommunity.com“ vytvoří '
        '„steamcommurnity.com“ nebo zamění písmeno „l“ za číslici „1“. Vytvoří přesnou vizuální kopii přihlašovacího formuláře, '
        'kde nic netušící hráč zadá své jméno, heslo i kód ze Steam Guardu. Útočník tyto údaje okamžitě zachytí a během '
        'několika sekund zcizí herní účet i veškerý inventář skinů [2].'
    )

    # 10. Kapitola 4: Prevence a účinná ochrana uživatele (Nová část teorie)
    doc.add_page_break()
    add_heading('4. Prevence a účinná ochrana uživatele', level=1)
    add_p(
        'Stoprocentní ochrana v kybernetickém světě neexistuje, avšak dodržováním několika základních technických a '
        'organizačních zásad lze eliminovat naprostou většinu běžných hrozeb. Bezpečnost není jednorázový produkt, ale '
        'neustálý proces zodpovědného chování uživatele [2].'
    )

    add_heading('4.1 Aktualizace operačního systému a softwaru', level=2)
    add_p(
        'Základním pilířem obrany je pravidelná instalace aktualizací operačního systému i veškerých používaných aplikací '
        '(prohlížeče, přehrávače, hry). Výrobci softwaru neustále objevují bezpečnostní slabiny (tzv. zranitelnosti), které '
        'útočníci zneužívají k průniku do počítače. Jakmile je vydána záplata (patch), informace o zranitelnosti se stává '
        'veřejnou. Pokud uživatel aktualizace odkládá, ponechává svůj počítač otevřený všem známým útokům [7].'
    )

    add_heading('4.2 Antivirový software a vestavěná ochrana Windows', level=2)
    add_p(
        'Moderní antivirové programy již nepracují pouze na základě porovnávání známých databází podpisů virů, ale využívají '
        'pokročilou heuristickou analýzu a strojové učení k detekci podezřelého chování programů v reálném čase. '
        'Pro běžného domácího uživatele i žáka školy představuje dnes vestavěná ochrana v systému Windows (Microsoft Defender) '
        'velmi spolehlivé a plnohodnotné řešení, které nezpomaluje systém a je pravidelně aktualizováno několikrát denně. '
        'Žádný antivirus však nedokáže uživatele ochránit, pokud sám dobrovolně potvrdí spuštění neznámého souboru nebo '
        'ignoruje varovná hlášení systému SmartScreen [1].'
    )

    add_heading('4.3 Silná hesla a vícefaktorové ověřování (2FA/MFA)', level=2)
    add_p(
        'Používání jednoduchých hesel (např. „123456“, „heslo123“, datum narození) nebo používání jednoho stejného hesla pro '
        'školní účet, e-mail i herní platformy představuje obrovské bezpečnostní riziko. Pokud dojde k úniku databáze na jedné '
        'herní stránce, útočníci okamžitě zkoušejí stejnou kombinaci přihlašovacích údajů i na ostatních službách (tzv. credential stuffing).'
    )
    add_p(
        'Nejúčinnější obranou proti krádeži účtu je zapnutí dvoufázového ověřování (2FA – Two-Factor Authentication). '
        'Při přihlášení nestačí pouze znalost hesla, ale je vyžadován druhý faktor – potvrzení v mobilní aplikaci '
        '(Google Authenticator, Steam Guard, Microsoft Authenticator) nebo hardwarový bezpečnostní klíč (FIDO2/YubiKey). '
        'I kdyby útočník heslo získal přes phishing, bez fyzického přístupu k telefonu se do účtu nedostane [2].'
    )

    add_heading('4.4 Pravidla bezpečného zálohování (pravidlo 3-2-1)', level=2)
    add_p(
        'Jedinou stoprocentní obranou proti ztrátě dat v důsledku ransomwaru, krádeže zařízení nebo hardwarové poruchy je '
        'pravidelné zálohování. V odborné praxi je celosvětově uznávaným standardem pravidlo 3-2-1 [1]:'
    )
    add_p('• 3 kopie dat: originální data v počítači plus alespoň 2 další nezávislé záložní kopie.', space_after=3, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('• 2 různá média: data mít uložena na dvou odlišných technologiích (např. interní SSD disk a externí přenosný pevný disk USB).', space_after=3, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('• 1 záloha mimo domov (off-site): alespoň jedna kopie musí být fyzicky umístěna jinde – například na šifrovaném cloudovém úložišti nebo u prarodičů. Pokud by došlo k požáru, krádeži nebo napadení celé domácí sítě ransomwarem, data zůstanou v bezpečí.', space_after=6, align=WD_ALIGN_PARAGRAPH.LEFT)

    # 11. Kapitola 5: Metodika praktické části (Povinný odstavec z listopadového zadání)
    doc.add_page_break()
    add_heading('5. Metodika praktické části', level=1)
    add_p(
        'Aby závěrečná práce nepřinášela pouze teoretické poznatky z literatury, je její praktická část postavena na vlastním '
        'kvantitativním empirickém šetření. Cílem tohoto výzkumu je zjistit, jak jsou žáci končících ročníků základní školy '
        'připraveni na reálné kybernetické hrozby, se kterými se běžně setkávají ve svých e-mailových schránkách a na sociálních sítích.'
    )

    add_heading('5.1 Metodika sběru dat a charakteristika zkoumaného vzorku', level=2)
    add_p(
        'Pro účely praktické části bylo zvoleno dotazníkové šetření realizované formou interaktivního simulačního testu, který jsem '
        'samostatně naprogramoval jako webovou aplikaci a publikoval na platformě GitHub Pages na adrese '
        'https://kost1al-source.github.io/phishing-test/. Výzkumný vzorek tvoří žáci 8. a 9. ročníků naší školy (ZŠ Husova Liberec), '
        'konkrétně třídy 8.A, 8.B, 9.A, 9.B a 9.C. Test je zcela anonymní, nevyžaduje žádné osobní údaje respondentů (zaznamenává '
        'pouze třídu a čas vyplnění) a z etického hlediska nesbírá žádná hesla. Výsledky se v reálném čase ukládají přes zabezpečené '
        'rozhraní do Google Tabulky a administrátorského panelu chráněného PIN kódem, což umožňuje detailní statistické vyhodnocení '
        'úspěšnosti po jednotlivých otázkách i třídách.'
    )

    add_heading('5.2 Popis modelových testovacích scénářů', level=2)
    add_p(
        'Test nesestává z pouhých teoretických pouček, ale simuluje 4 konkrétní, vysoce realistické modelové situace:'
    )
    add_p('1. Scénář 1 – Školní systém Bakaláři (Legitimní zpráva): běžné oznámení o pedagogických konzultacích s platnou adresou odesílatele upozorneni@bakalari.cz a odkazem na školní web. Testuje schopnost poznat bezpečný e-mail a nepodléhat zbytečné paranoie.', space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('2. Scénář 2 – Falešná výhra Apple iPhone 16 Pro (Phishing / Scam): nápadný podvod z ruské bezplatné domény slibující telefon zdarma za úhradu poštovného 29 Kč platební kartou pod časovým tlakem 15 minut.', space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('3. Scénář 3 – Výhrůžka smazáním účtu Steam (Typosquatting): velmi zdařilá napodobenina zprávy od Valve hrozící trvalým zablokováním účtu za podvádění ve hře CS2 s podvrženým odkazem steamcommurnity-login.com. Cílí na zcizení herního účtu.', space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_p('4. Scénář 4 – Zásilka České pošty s nebezpečnou přílohou (Trojský kůň): e-mail vyžadující doplacení cla s přiloženým souborem Faktura_Doplatek.pdf.exe. Testuje znalost maskování spustitelných souborů.', space_after=8, align=WD_ALIGN_PARAGRAPH.LEFT)

    # 12. Kapitoly 6 a 7: Koncept dalších fází
    doc.add_page_break()
    add_heading('6. Koncept vyhodnocení výsledků a modelová architektura', level=1)
    add_p(
        'V navazující fázi práce (kapitola 6 – Výsledky) budou data sebraná z dotazníkového šetření přehledně zpracována do '
        'statistických tabulek a grafů. Sledována bude celková úspěšnost žáků, úspěšnost podle ročníků a identifikace nejnebezpečnějšího '
        'scénáře (který podvod žáky nejčastěji oklamal). Součástí praktické části je také vzdělávací webová stránka projektu '
        '(https://kost1al-source.github.io/zaverecna-prace/) a bezpečné diagramové schéma životního cyklu trojského koně.'
    )

    add_heading('7. Diskuze a srovnání se světovými statistikami', level=1)
    add_p(
        'V kapitole Diskuze budou výsledky zjištěné mezi žáky ZŠ Husova Liberec porovnány s mezinárodními zprávami ENISA Threat Landscape '
        'a Verizon Data Breach Investigations Report [5, 6]. Na základě těchto zjištění budou v závěru práce formulována konkrétní '
        'doporučení pro školní výuku informatiky, jak žáky lépe připravit na nástrahy digitálního světa.'
    )

    # 13. Závěr
    doc.add_page_break()
    add_heading('Závěr', level=1)
    add_p(
        'Předložená práce v této fázi zpracování přináší ucelený teoretický přehled v oblasti škodlivého softwaru a připravenou '
        'metodiku pro praktické empirické šetření. Podařilo se jasně vymezit rozdíly mezi viry, trojskými koni a ransomwarem, '
        'popsat psychologické zbraně sociálního inženýrství a definovat pravidla prevence. Praktický interaktivní test je plně '
        'funkční a připravený k distribuci žákům školy.'
    )

    # 14. Seznam zdrojů
    doc.add_page_break()
    add_heading('Seznam použitých zdrojů', level=1)

    sources = [
        '[1] KOLOUCH, Jan. CyberCrime. Praha: CZ.NIC, 2016. Edice CZ.NIC. 424 s. ISBN 978-80-88168-15-7. Dostupné také z: https://knihy.nic.cz',
        '[2] KOLOUCH, Jan, Pavel BAŠTA a kol. CyberSecurity. Praha: CZ.NIC, 2019. Edice CZ.NIC. 612 s. ISBN 978-80-88168-31-7. Dostupné také z: https://knihy.nic.cz',
        '[3] NÚKIB (Národní úřad pro kybernetickou a informační bezpečnost). Zpráva o stavu kybernetické bezpečnosti České republiky za rok 2024. Brno: NÚKIB, 2025. Dostupné z: https://nukib.gov.cz',
        '[4] NÚKIB. Vzdělávací portál a metodika prevence sociálního inženýrství. Dostupné z: https://osveta.nukib.gov.cz',
        '[5] ENISA (Agentura Evropské unie pro kybernetickou bezpečnost). ENISA Threat Landscape 2024/2025: Cyber Threat Analysis. Heraklion: ENISA, 2024. Dostupné z: https://www.enisa.europa.eu',
        '[6] VERIZON. 2024–2025 Data Breach Investigations Report (DBIR). Verizon Business, 2025. Dostupné z: https://www.verizon.com/business/resources/reports/dbir/',
        '[7] NIST (National Institute of Standards and Technology). Special Publication 800-83 Rev. 1: Guide to Malware Incident Prevention and Handling. Gaithersburg: NIST, 2020. Dostupné z: https://csrc.nist.gov',
        '[8] BABAK, Kostiantyn. Phishing Test: Interaktivní modelové šetření [online]. Liberec: ZŠ Husova, 2026. Dostupné z: https://kost1al-source.github.io/phishing-test/'
    ]

    for src in sources:
        add_p(src, space_after=8, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.15)

    # Uložení do documents i přímo do kořene složky
    out_root = 'Zaverecna_prace_Skodlivy_software_Babak.docx'
    out_docs = 'documents/Zaverecna_prace_Skodlivy_software_Babak.docx'
    doc.save(out_root)
    doc.save(out_docs)
    print(f'Done! Successfully built November version:\n- {out_root}\n- {out_docs}')

if __name__ == '__main__':
    build_november()
