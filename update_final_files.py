import os
from docx import Document
from docx.shared import Pt, Inches
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt

def update_final_files():
    docx_filename = "Analiza_Historyczno-Badawcza_Najnowszego Pustelnika170520266.docx"
    pptx_filename = "Analiza_Historyczno-Badawcza_Najnowszego Pustelnika170520266.pptx"

    # --- DOCX UPDATE ---
    doc = Document()
    doc.add_heading('DOGŁĘBNA ANALIZA HISTORYCZNO-TOPOGRAFICZNA PUSTELNIKA', 0)

    # 1. Chronologia i Pustelnia
    doc.add_heading('1. Fundamenty: 1440, 1465 i Kaplica św. Katarzyny', level=1)
    doc.add_paragraph(
        "Kluczowe dla zrozumienia Pustelnika jest rozróżnienie między tradycją a dokumentami:\n"
        "• Rok 1440: Moment 'papierowej' fundacji (Bolesław IV). Ziemia wyznaczona w puszczy, ale bez osadników.\n"
        "• Rok 1465: Eremita Bartłomiej ze Słuńczewa zasiedla las Cisek. Buduje kaplicę, która staje się 'Punktem Zero'.\n"
        "• Topografia: Kaplica stała na północnym brzegu rzeki Czarnej (rejon dzisiejszej ul. Nadrzecznej)."
    )

    # 2. PRZEŁOMOWY DOKUMENT: Akt z 1473 r. (Słowo w słowo)
    doc.add_heading('2. PEŁNE TŁUMACZENIE AKTU FUNDACYJNEGO (1473 r.)', level=1)

    latin_text = (
        "Fundus Ecclesiae in Czyszek de cruda radice Anno domini millesimo quadringentesimo [septuagesimo tertio] etc. "
        "In Camyenicza feria quarta proxima post Jacobi apostoli... "
        "Illustrissimus princeps dominus Boleslaus dux... zelo devotionis accensus volens suae et predecessorum suorum sexus utriusque saluti succurrere... "
        "duos mansos mensurae Chulmensis ex opposito oraculi sanctae Katherine penes fluvium Czarna iacentes de cruda radice in Czyszek..."
    )

    doc.add_heading('Rekonstrukcja Łacińska:', level=2)
    doc.add_paragraph(latin_text).italic = True

    doc.add_heading('Tłumaczenie słowo w słowo:', level=2)
    translations = [
        ("Fundus Ecclesiae in Czyszek de cruda radice", "Fundacja Kościoła w Cisku na surowym korzeniu (wykarczowanym)."),
        ("Anno domini millesimo quadringentesimo [1473]", "Roku Pańskiego tysiącznego czterechsetnego (siedemdziesiątego trzeciego)."),
        ("Illustrissimus princeps dominus Boleslaus dux", "Najjaśniejszy książę pan Bolesław (V), książę (Mazowsza)."),
        ("duos mansos mensurae Chulmensis", "dwie włóki miary chełmińskiej."),
        ("ex opposito oraculi sanctae Katherine", "naprzeciw kaplicy (oraculum - miejsce modlitwy) świętej Katarzyny."),
        ("penes fluvium Czarna iacentes", "przy rzece Czarnej leżące."),
        ("ratione sanctuarii", "z tytułu uposażenia (sanctuarium - poświętnego)."),
        ("dedit donavit et largitus est", "dał, darował i hojnie ofiarował."),
        ("absolvit ab omnibus solutionibus", "uwalnia od wszystkich opłat, czynszów św. Marcina i pańszczyzny.")
    ]

    for lat, pol in translations:
        p = doc.add_paragraph()
        p.add_run(f"{lat}: ").bold = True
        p.add_run(pol)

    doc.add_paragraph(
        "Kluczowe wnioski z aktu:\n"
        "1. Użycie słowa 'ORACULUM' zamiast 'Ecclesia' dla kaplicy św. Katarzyny świadczy o jej pierwotnym, pustelniczym charakterze.\n"
        "2. Lokalizacja 'PENES FLUVIUM CZARNA' (Przy rzece Czarnej) ostatecznie potwierdza nadrzeczne położenie pierwszej kaplicy."
    )

    # 3. Podział Nadań (1473 i 1477)
    doc.add_heading('3. Podział nadań: Kościół (1473) vs Wieś (1477)', level=1)
    doc.add_paragraph(
        "• 1473 r. – 2 WŁÓKI DLA KOŚCIOŁA: Nadane 'naprzeciw kaplicy'. To dzisiejsze ogrody plebańskie i teren plebanii.\n"
        "• 1477 r. – 12 WŁÓK DLA WSI: Nadane Andrzejowi z Chylina 'koło kaplicy'. To dzisiejsza zabudowa miejscowości."
    )

    # 4. Inne Akty (Skrót)
    doc.add_heading('4. Pozostałe dokumenty (Chronologia)', level=1)
    other_acts = [
        ("1482", "ks. Konrad III daje Maciejowi z Nieksyna karczmę przy gościńcu do Liwa."),
        ("1504", "Pojawienie się nazwy Pustelnik (P.) jako synonimu Czarnej Woli."),
        ("1540", "Erekcja parafii Pustelnik przez biskupa Jakuba Buczackiego."),
        ("1775", "Wizytacja: kościół zniszczony ze starości, kult w drewnianej kaplicy.")
    ]
    for d, t in other_acts:
        p = doc.add_paragraph()
        p.add_run(f"{d}: ").bold = True
        p.add_run(t)

    # 5. Topografia i Metodyka
    doc.add_heading('5. Topografia i Badania', level=1)
    doc.add_paragraph(
        "• Trakt Liwski: Pod asfaltem znajdują się historyczne 'kocie łby'.\n"
        "• Metoda LIDAR: Kluczowa do znalezienia fundamentów w ogrodzie plebańskim.\n"
        "• Lokalizacja: Pierwotna pustelnia stała przy dzisiejszej ul. Nadrzecznej."
    )

    doc.save(docx_filename)

    # --- PPTX UPDATE ---
    prs = Presentation()

    slides = [
        ("Pustelnik: Analiza Archiwalna", "Dokument Fundacyjny 1473 r.\ni Monografia Topograficzna"),
        ("Akt Fundacyjny 1473 r.", "• 'duos mansos ex opposito oraculi' - 2 włóki naprzeciw kaplicy.\n• 'penes fluvium Czarna' - przy rzece Czarnej.\n• 'de cruda radice' - na surowym korzeniu (karczunek)."),
        ("Tłumaczenie Słowo w Słowo", "• oraculum = kaplica / miejsce modlitwy.\n• sanctuarium = poświętne (uposażenie).\n• absolvit = uwalnia od podatków i pańszczyzny."),
        ("Podział Nadań 1473/1477", "• 1473: Jądro kościelne (2 włóki, dzisiejsza plebania).\n• 1477: Lokacja wsi (12 włók, otoczenie kaplicy).\n• Oś obu nadań: Rzeka Czarna i Trakt Liwski."),
        ("Lokalizacja i 'Kocie Łby'", "• Kaplica (1465): Przy dzisiejszej ul. Nadrzecznej.\n• Plebania: Stoi na gruncie nadanym 'naprzeciw kaplicy'.\n• Trakt: Historyczny bruk pod współczesnym asfaltem."),
        ("Dalsze Badania", "• LIDAR: Analiza rzeźby ogrodu plebańskiego.\n• Georadar: Prześwietlanie terenów nad rzeką.\n• Archiwa: Wizytacja 1775 w Płocku (opis kościoła).")
    ]

    for t, b in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = t
        slide.placeholders[1].text = b

    prs.save(pptx_filename)

if __name__ == "__main__":
    update_final_files()
    print("Pliki zaktualizowane o akt z 1473 r.")
