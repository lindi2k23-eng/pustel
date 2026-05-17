import os
from docx import Document
from docx.shared import Pt, Inches
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt

def create_final_ultimate_docx():
    docx_filename = "Analiza_Historyczno-Badawcza_Najnowszego Pustelnika170520266.docx"

    doc = Document()
    doc.add_heading('DOGŁĘBNA ANALIZA HISTORYCZNO-TOPOGRAFICZNA PUSTELNIKA', 0)

    # 1. Chronologia i Eremita
    doc.add_heading('1. Początki: Eremita Bartłomiej i Zagadka Chronologii', level=1)
    doc.add_paragraph(
        "• Rok 1440: Pierwotny zamysł fundacyjny (Bolesław IV). Ziemia zarezerwowana w puszczy.\n"
        "• Rok 1465: Eremita Bartłomiej ze Słuńczewa zasiedla teren. Buduje kaplicę św. Katarzyny.\n"
        "• KIM BYŁ BARTŁOMIEJ? Pochodził ze Słuńczewa (dzisiejszy warszawski SŁUŻEW). Służew był centrum kultu św. Katarzyny (parafia od 1238 r.), skąd Bartłomiej przeniósł wezwanie do Pustelnika."
    )

    # 2. Kontekst Administracyjny: Dlaczego Księgi Zakroczymskie?
    doc.add_heading('2. Ślady Biurokratyczne: Dlaczego Służew jest w Księgach Zakroczymskich?', level=1)
    doc.add_paragraph(
        "Pojawienie się zapisów o Słuńczewie/Służewie (Ziemia Warszawska) w Księgach Zakroczymskich wynika z:\n"
        "1. MOBILNOŚCI SĄDÓW: Książęta mazowieccy podróżowali z sądem objazdowym. Akty wpisywano do ksiąg miejsca, w którym akurat przebywał władca (np. w Zakroczymiu).\n"
        "2. RANGI WYDARZENIA: Lokacja Pustelnika była sprawą państwową, rejestrowaną w najważniejszych księgach książęcych, niezależnie od granic powiatów."
    )

    # 3. Topografia: Podział przez rzekę Czarną
    doc.add_heading('3. Układ Przestrzenny: Podział przez rzekę Czarną', level=1)
    p = doc.add_paragraph()
    p.add_run('Rzeka Czarna stanowi historyczną i topograficzną oś Pustelnika:').bold = True

    doc.add_paragraph(
        "• WSCHODNI BRZEG (Rejon dzisiejszej ul. Kopernika/Nadrzecznej):\n"
        "Tu prawdopodobnie stała PIERWOTNA KAPLICA Bartłomieja (1465 r.). Było to miejsce najbliżej wody, serce pustelni.\n\n"
        "• ZACHODNI BRZEG (Miejsce dzisiejszej Plebanii):\n"
        "To jest teren nadany 'NAPRZECIW KAPLICY' (ex opposito) w 1473 r. Pleban Maciej (1482 r.) wybudował tu swój dom na suchym wzniesieniu, oddzielając się rzeką od reszty wsi.\n\n"
        "• OGRÓD PLEBAŃSKI (Zachodni brzeg):\n"
        "Miejsce lokalizacji starego kościoła (sprzed 1840 r.) oraz NAJSTARSZEGO CMENTARZA (XV-XVIII w.)."
    )

    # 4. Podział Nadań (1473 i 1477)
    doc.add_heading('4. Struktura Gruntowa: 2 włóki vs 12 włók', level=1)
    doc.add_paragraph(
        "• 1473 r. (2 włóki): Majątek kościelny (Sanctuarium) nadany 'naprzeciw kaplicy' – dzisiejsza strona zachodnia (Plebania).\n"
        "• 1477 r. (12 włók): Majątek wiejski nadany Andrzejowi z Chylina – teren dzisiejszej wsi otaczający kaplicę."
    )

    # 5. Słowiańskie Korzenie
    doc.add_heading('5. Dziedzictwo Słowiańskie', level=1)
    doc.add_paragraph(
        "1. DZIADOSZ: Związek z obrzędem 'Dziadów' i kultem przodków.\n"
        "2. CISEK: Cis jako święte drzewo Słowian (symbol wieczności i zaświatów).\n"
        "3. RZEKA CZARNA: Mitologiczna granica między światami."
    )

    # 6. Tłumaczenie Aktu 1473 (Słowo w Słowo)
    doc.add_heading('6. Akt Fundacyjny 1473 r. - Pełne Tłumaczenie', level=1)
    translations = [
        ("Fundus Ecclesiae in Czyszek de cruda radice", "Fundacja Kościoła w Cisku na surowym korzeniu."),
        ("ex opposito oraculi sanctae Katherine", "naprzeciw kaplicy (miejsca modlitwy) świętej Katarzyny."),
        ("penes fluvium Czarna iacentes", "przy rzece Czarnej leżące."),
        ("ratione sanctuarii", "z tytułu uposażenia (poświętnego).")
    ]
    for lat, pol in translations:
        p = doc.add_paragraph()
        p.add_run(f"{lat}: ").bold = True
        p.add_run(pol)

    # 7. Instrukcja Badawcza
    doc.add_heading('7. Przewodnik dla Badacza', level=1)
    doc.add_paragraph(
        "• LIDAR: Szukać fundamentów 'w ogrodzie' na zachodnim brzegu.\n"
        "• GEORADAR: Zalecany przy ul. Nadrzecznej (wschodni brzeg - stara kaplica) oraz w ogrodzie plebańskim (zachodni brzeg - kościół i cmentarz).\n"
        "• TRAKT: Pod asfaltem 'kocie łby' (bruk XIX-wieczny)."
    )

    doc.save(docx_filename)

def create_final_ultimate_pptx():
    pptx_filename = "Analiza_Historyczno-Badawcza_Najnowszego Pustelnika170520266.pptx"
    prs = Presentation()

    slides = [
        ("Pustelnik: Synteza Badawcza", "Topografia, Administracja i Kultura"),
        ("Słuńczew w Zakroczymiu", "• Słuńczew = Służew warszawski.\n• Zapis w Księgach Zakroczymskich ze względu na mobilność sądów książęcych.\n• Bartłomiej (1465) – kapłan ze Służewa."),
        ("Podział przez Rzekę", "• Wschód (ul. Kopernika): Pierwotna kaplica (1465).\n• Zachód (Plebania): Teren nadany 'naprzeciw' (1473).\n• Rzeka Czarna jako oś graniczna."),
        ("Najstarszy Cmentarz", "• Lokalizacja: Ogród plebański (zachodni brzeg).\n• Czas: XV w. – 1840 r.\n• Pochówki wokół starego kościoła."),
        ("Struktura Nadań", "• 2 włóki kościelne (1473): Siedlisko Plebanii.\n• 12 włók wiejskich (1477): Budowa wsi dookoła."),
        ("Słowiańskie Toponimy", "• DZIADOSZ: Kult przodków / Dziady.\n• CISEK: Święte drzewo Cis.\n• CZARNA: Rzeka graniczna zaświatów.")
    ]

    for t, b in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = t
        slide.placeholders[1].text = b

    prs.save(pptx_filename)

if __name__ == "__main__":
    create_final_ultimate_docx()
    create_final_ultimate_pptx()
    print("Dokumentacja zaktualizowana o podział rzeki i kontekst Zakroczymia.")
