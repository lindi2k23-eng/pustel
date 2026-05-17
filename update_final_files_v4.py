import os
from docx import Document
from docx.shared import Pt, Inches
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt

def update_comprehensive_files_v4():
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
        "• SŁUŃCZEW = SŁUŻEW (Warszawa): Bartłomiej prawdopodobnie przeniósł kult św. Katarzyny z jej najstarszego mazowieckiego ośrodka."
    )

    # 2. Orientacja i Najstarszy Cmentarz
    doc.add_heading('2. Układ Przestrzenny i Dawne Cmentarzysko', level=1)
    doc.add_paragraph(
        "Analiza wskazuje na rygorystyczny układ symboliczny (Wschód-Zachód):\n"
        "• PLEBANIA (ZACHÓD): Miejsce życia, dom plebana nadany 'naprzeciw' kaplicy w 1473 r.\n"
        "• OGRÓD PLEBAŃSKI (ŚRODEK): Tu znajdował się NAJSTARSZY CMENTARZ (XV–XVIII w.). Zanim w 1840 r. założono obecny cmentarz, mieszkańców chowano wokół kościoła, czyli właśnie w pasie ogrodu między plebanią a rzeką.\n"
        "• KOŚCIÓŁ/KAPLICA (WSCHÓD): Świątynia zorientowana na wschód (symbol zmartwychwstania), położona bliżej rzeki Czarnej."
    )

    # 3. Powiązania Słowiańskie (Etymologia i Mitologia)
    doc.add_heading('3. Korzenie Słowiańskie i Mitologia Miejsca', level=1)
    doc.add_paragraph(
        "Pustelnik i jego okolice kryją nazwy o głębokim rodowodzie słowiańskim:\n"
        "1. DZIADOSZ (Dziadoszu): Etymologia bezpośrednio związana ze słowiańskim 'Dziadem' (przodkiem, duchem opiekuńczym) oraz obrzędem DZIADÓW. Może sugerować dawne miejsce kultu przodków.\n"
        "2. CISEK (Czyszek): Od CISA – drzewa świętego i magicznego w wierzeniach Słowian, symbolu nieśmiertelności i związku ze światem zmarłych.\n"
        "3. RZEKA CZARNA: W mitologii słowiańskiej nazwa ta często oznaczała granicę między światem żywych a krainą cieni (zaświatami)."
    )

    # 4. Akt Fundacyjny (1473 r.) - Tłumaczenie
    doc.add_heading('4. Pełne Tłumaczenie Aktu Fundacyjnego (1473 r.)', level=1)
    translations = [
        ("Fundus Ecclesiae in Czyszek de cruda radice", "Fundacja Kościoła w Cisku na surowym korzeniu (wykarczowanym)."),
        ("ex opposito oraculi sanctae Katherine", "naprzeciw kaplicy (oraculum - miejsce modlitwy) świętej Katarzyny."),
        ("penes fluvium Czarna iacentes", "przy rzece Czarnej leżące (potwierdzenie lokalizacji kaplicy)."),
        ("ratione sanctuarii", "z tytułu uposażenia (sanctuarium - poświętnego).")
    ]
    for lat, pol in translations:
        p = doc.add_paragraph()
        p.add_run(f"{lat}: ").bold = True
        p.add_run(pol)

    # 5. Podział Nadań: 2 włóki vs 12 włók
    doc.add_heading('5. Struktura Gruntowa: Podział 1473 / 1477', level=1)
    doc.add_paragraph(
        "• 1473 r. (2 włóki): Jądro kościelne 'naprzeciw kaplicy' – dzisiejsze siedlisko plebańskie.\n"
        "• 1477 r. (12 włók): Lokacja wsi 'koło kaplicy' – teren zabudowy wiejskiej (kmiecej)."
    )

    # 6. Instrukcja Badawcza
    doc.add_heading('6. Jak szukać fundamentów i grobów?', level=1)
    doc.add_paragraph(
        "1. LIDAR (Geoportal): Szukać śladów fundamentów starego kościoła 'w ogrodzie'.\n"
        "2. GEORADAR: Zalecany w ogrodzie plebańskim w celu lokalizacji najstarszych pochówków (XV–XVIII w.) oraz fundamentów pierwszej kaplicy.\n"
        "3. ARCHEOLOGIA: Wschodnia część ogrodu (bliżej rzeki) jest najbardziej obiecująca pod względem znalezisk słowiańskich i wczesnochrześcijańskich."
    )

    doc.save(docx_filename)

    # --- PPTX UPDATE ---
    prs = Presentation()
    slides = [
        ("Pustelnik: Słowiańskie Korzenie", "Analiza Mitologii i Układu Przestrzennego"),
        ("Układ Symboliczny (Wschód-Zachód)", "• Zachód: Plebania (Świat żywych).\n• Środek: Ogród / Najstarszy Cmentarz (Granica).\n• Wschód: Kościół (Świat boski/zmartwychwstanie)."),
        ("Najstarszy Cmentarz", "• Przed 1840 r.: Pochówki wokół kościoła w ogrodzie plebańskim.\n• Ogrody plebańskie kryją groby z lat 1465–1840.\n• Nowy cmentarz (z XIX w.) jest oddalony od świątyni."),
        ("Etymologia Słowiańska", "• DZIADOSZ: Związek z obrzędem 'Dziadów' i kultem przodków.\n• CISEK: Cis jako drzewo święte i magiczne.\n• CZARNA: Rzeka jako granica zaświatów."),
        ("Akt 1473: Topografia", "• 'penes fluvium Czarna' - kaplica przy rzece.\n• 'ex opposito oraculi' - plebania naprzeciw kaplicy.\n• Teren dzisiejszej plebanii to historyczne uposażenie."),
        ("Podział Nadań", "• 1473: 2 włóki dla kościoła (Sanctuarium).\n• 1477: 12 włók dla osadników (Wieś).\n• Kaplica Bartłomieja jako punkt centralny osadnictwa.")
    ]
    for t, b in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = t
        slide.placeholders[1].text = b

    prs.save(pptx_filename)

if __name__ == "__main__":
    update_comprehensive_files_v4()
    print("Pliki zaktualizowane o wątki słowiańskie i cmentarne.")
