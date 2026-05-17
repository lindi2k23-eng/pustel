import os
from docx import Document
from docx.shared import Pt, Inches
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt

def update_comprehensive_files_v3():
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

    # 2. Tożsamość Eremity i Kult św. Katarzyny
    doc.add_heading('2. Tożsamość Eremity: Bartłomiej de Slunczew (Służew)', level=1)
    doc.add_paragraph(
        "• SŁUŃCZEW = SŁUŻEW (Warszawa): Bartłomiej pochodził z najstarszej mazowieckiej parafii św. Katarzyny na Służewie.\n"
        "• PRZENIESIENIE KULTU: To on zaszczepił wezwanie św. Katarzyny w lesie Cisek.\n"
        "• KONTEKST: Zbieżność z klasztorem w Górach Świętokrzyskich (fundowanym w tym samym czasie) wynika z ogólnopolskiego renesansu fundacji eremickich."
    )

    # 3. Relacja Przestrzenna: Kaplica a Kościół
    doc.add_heading('3. Układ Przestrzenny: Czy stały koło siebie?', level=1)
    doc.add_paragraph(
        "Analiza wskazuje na ewoluujący kompleks sakralny (Sanctuarium):\n"
        "• ETAP I (1465): Pierwotna kaplica (Oraculum) Bartłomieja – stała najniżej, przy samej wodzie (ul. Nadrzeczna).\n"
        "• ETAP II (1540): Kościół parafialny – wzniesiony 'obok' kaplicy, na nieco wyższym terenie (dzisiejsze ogrody plebańskie). Często opisywany jako kościół 'w ogrodzie'.\n"
        "• ETAP III (1843): Przesunięcie jądra sakralnego na wzgórze (miejsce obecnego kościoła).\n"
        "Konkluzja: W ogrodach plebańskich znajdują się nakładające się na siebie fundamenty różnych budowli stojących niegdyś obok siebie."
    )

    # 4. Akt Fundacyjny (1473 r.) - Tłumaczenie
    doc.add_heading('4. Pełne Tłumaczenie Aktu Fundacyjnego (1473 r.)', level=1)
    translations = [
        ("Fundus Ecclesiae in Czyszek de cruda radice", "Fundacja Kościoła w Cisku na surowym korzeniu (wykarczowanym)."),
        ("ex opposito oraculi sanctae Katherine", "naprzeciw kaplicy (oraculum) świętej Katarzyny."),
        ("penes fluvium Czarna iacentes", "przy rzece Czarnej leżące."),
        ("ratione sanctuarii", "z tytułu uposażenia (poświętnego).")
    ]
    for lat, pol in translations:
        p = doc.add_paragraph()
        p.add_run(f"{lat}: ").bold = True
        p.add_run(pol)

    # 5. Podział Nadań: 2 włóki vs 12 włók
    doc.add_heading('5. Podział nadań: Kościół (1473) vs Wieś (1477)', level=1)
    doc.add_paragraph(
        "• 1473 r. (2 włóki): Nadane 'NAPRZECIW' kaplicy. Teren dzisiejszej plebanii (ul. Kopernika).\n"
        "• 1477 r. (12 włók): Nadane Andrzejowi z Chylina 'KOŁO' kaplicy. Teren zabudowy wiejskiej."
    )

    # 6. Trakt, "Kocie Łby" i Metodyka
    doc.add_heading('6. Topografia i Badania', level=1)
    doc.add_paragraph(
        "• Trakt Liwski: Historyczne 'kocie łby' pod asfaltem.\n"
        "• LIDAR: Szukać fundamentów w pasie od plebanii do rzeki.\n"
        "• Archiwa: Wizytacja 1775 (Płock) – opis kościoła i kaplicy stojących obok siebie."
    )

    doc.save(docx_filename)

    # --- PPTX UPDATE ---
    prs = Presentation()
    slides = [
        ("Pustelnik: Ewolucja Sanktuarium", "Czy Kaplica i Kościół stały koło siebie?"),
        ("Bartłomiej ze Służewa", "• Pochodzenie: Słuńczew = Służew warszawski.\n• Kult św. Katarzyny przeniesiony do puszczy.\n• Model eremity świętokrzyskiego."),
        ("Układ: Kaplica - Kościół", "• Kaplica (1465): Przy rzece (ul. Nadrzeczna).\n• Kościół (1540): 'W ogrodzie' obok kaplicy.\n• Plebania: Nadana 'naprzeciw' kaplicy."),
        ("Akt 1473: Fundament Prawny", "• 'penes fluvium Czarna' - przy rzece.\n• 'ex opposito oraculi' - naprzeciw kaplicy.\n• 2 włóki poświętnego (dzisiejsza plebania)."),
        ("Podział nadań", "• 1473 (2 włóki): Jądro kościelne.\n• 1477 (12 włók): Osadnictwo wiejskie (kmiecie).\n• Centralny punkt: Kaplica św. Katarzyny."),
        ("Badania w terenie", "• LIDAR: Analiza fundamentów w ogrodach.\n• Georadar: Prześwietlanie pasa nad rzeką.\n• Trakt: Historyczne 'kocie łby' pod drogą.")
    ]
    for t, b in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = t
        slide.placeholders[1].text = b

    prs.save(pptx_filename)

if __name__ == "__main__":
    update_comprehensive_files_v3()
    print("Pliki zaktualizowane o relację kościoła i kaplicy.")
