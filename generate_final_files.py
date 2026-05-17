import os
from docx import Document
from docx.shared import Pt, Inches
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt

def create_final_comprehensive_docx():
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

    # 2. Podział Nadań: Kościół vs Wieś (1473 i 1477)
    doc.add_heading('2. Struktura Własności: Podział nadań z 1473 i 1477 r.', level=1)
    p = doc.add_paragraph()
    p.add_run('W latach 70. XV wieku książę Bolesław V dokonał precyzyjnego podziału ziemi wokół kaplicy:').bold = True

    doc.add_paragraph(
        "A. 1473 r. – 2 WŁÓKI DLA KOŚCIOŁA (Poświętne):\n"
        "Lokalizacja: 'NAPRZECIW KAPLICY' (ex opposito capelle).\n"
        "Analiza: Są to grunty dzisiejszej plebanii i ogrodów plebańskich. Jeśli kaplica stała nad rzeką, to ziemia 'naprzeciw' (idąc w górę terenu) to właśnie dzisiejsze siedlisko parafialne.\n\n"
        "B. 1477 r. – 12 WŁÓK DLA WSI (Lokacja Andrzeja z Chylina):\n"
        "Lokalizacja: 'KOŁO KAPLICY' i 'KOŁO RZEKI CZARNEJ'.\n"
        "Analiza: Tereny otaczające jądro kościelne, przeznaczone pod zabudowę wiejską i pola uprawne kmieci. To dzisiejszy obszar całej miejscowości Pustelnik."
    )

    # 3. Tłumaczenie Aktów (Słowo w Słowo)
    doc.add_heading('3. Tłumaczenie Aktów Archiwalnych', level=1)
    acts = [
        ("1465", "ksiądz Bartłomiej ze Słuńczewa eremita in C. (Zakr. 5, 443v)", "Bartłomiej, pustelnik ze Służewa, zamieszkuje las Cisek."),
        ("1473", "ks. Bolesław V daje 2 wł. na nowym korzeniu naprzeciw kaplicy Ś. Katarzyny...", "Fundacja uposażenia plebana (poświętne) tuż przy pustelni."),
        ("1477", "ks. Bolesław V daje Andrzejowi z Chylina nowo lokowaną wieś Cz.W. k. kaplicy eremity...", "Lokacja wsi Czarna Wola na 12 włókach otaczających kaplicę."),
        ("1482", "ks. Konrad III daje Maciejowi z Nieksyna 4 wł. i karczmę przy gościńcu do Liwa.", "Ustanowienie karczmy i infrastruktury handlowej na trakcie."),
        ("1526", "ks. Anna daje Dorocie brzeg rz. Czarnej z 1/2 stawu i młyna naprzeciw wsi Cz.W.", "Potwierdzenie rzeki Czarnej jako granicy i osi gospodarczej."),
        ("1540", "Jakub bp płoc. eryguje par. Pustelnik przy istniejącym kościele filialnym Ś. Katarzyny.", "Podniesienie kaplicy do rangi samodzielnej parafii."),
        ("1775", "kośc. zniszczony ze starości, zastępuje go kaplica drew. pw. Ś. Katarzyny.", "Kryzys budowlany i tymczasowa kaplica przed fundacją Grabowskiego.")
    ]
    for date, orig, transl in acts:
        p = doc.add_paragraph()
        p.add_run(f"{date}: ").bold = True
        p.add_run(orig + "\n")
        p.add_run(f"ANALIZA: {transl}").italic = True

    # 4. Trakt, "Kocie Łby" i Metodyka
    doc.add_heading('4. Topografia i Badania', level=1)
    doc.add_paragraph(
        "• Trakt Liwski: Pod asfaltem znajdują się historyczne 'kocie łby' (bruk kamienny).\n"
        "• Lokalizacja kaplicy (1465): Ok. 30-70m od rzeki Czarnej (północny brzeg).\n"
        "• Metoda LIDAR: Kluczowa do znalezienia fundamentów w ogrodzie plebańskim.\n"
        "• Metoda GPR (Georadar): Zalecana wzdłuż ul. Nadrzecznej i Kopernika."
    )

    doc.save('Pustelnik_Analiza_Finalna.docx')

def create_final_pptx():
    prs = Presentation()
    data = [
        ("Pustelnik: Analiza Nadań", "XV-wieczna struktura wsi i kościoła"),
        ("1473 vs 1477", "• 1473: 2 włóki 'NAPRZECIW' kaplicy (Dla Kościoła - Plebania).\n• 1477: 12 włók 'KOŁO' kaplicy (Dla Wsi - Andrzej z Chylina).\n• Kaplica św. Katarzyny jako centrum obu nadań."),
        ("Kaplica nad Rzeką", "• Lokalizacja: Północny brzeg rzeki Czarnej.\n• Teren: Dzisiejsza ul. Nadrzeczna / Kopernika.\n• Fundamenty: Szukaj w niższej części ogrodu plebańskiego."),
        ("Gospodarka: Karczma i Młyn", "• 1482: Karczma przy gościńcu liwskim (Trakt Liwski).\n• 1525: Udziały w młynie i stawie na Czarnej.\n• Nawierzchnia: Historyczne 'kocie łby' pod asfaltem."),
        ("Instrukcja Badawcza", "• Geoportal/LIDAR: Warstwa cieniowania.\n• Archiwa: Płock (Wizytacja 1775), AGAD (Metryka Koronna).\n• Mapy: Karol de Perthees (1791).")
    ]
    for t, b in data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = t
        slide.placeholders[1].text = b
    prs.save('Pustelnik_Prezentacja_Finalna.pptx')

if __name__ == "__main__":
    create_final_comprehensive_docx()
    create_final_pptx()
    print("Pliki wygenerowane pomyślnie.")
