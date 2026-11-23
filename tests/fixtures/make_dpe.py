"""DPE de test FICTIFS pour tests/dpe-scan.*.js : DPE 2021 texte, DPE 2021 scanné, ancien DPE 2014, etc.
   python3 tests/fixtures/make_dpe.py <dossier>   (reportlab, Pillow, pdftoppm, img2pdf)"""
import io, os, random, subprocess, sys
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageDraw, ImageFont, ImageFilter

OUT = sys.argv[1] if len(sys.argv) > 1 else "/home/claude/gen/DPE_SIGEO"
os.makedirs(OUT, exist_ok=True)
W, H = A4
FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONTB = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont("DV", FONT)); pdfmetrics.registerFont(TTFont("DVB", FONTB))

COLORS = ["#009c6d", "#52b153", "#78bd76", "#f4e70f", "#f0b40f", "#eb8235", "#d7221f"]

def label_png(ep, ges, letter):
    """Étiquette énergie (image, comme dans les vrais DPE)."""
    img = Image.new("RGB", (900, 640), "white")
    d = ImageDraw.Draw(img)
    f = ImageFont.truetype(FONTB, 34); fs = ImageFont.truetype(FONT, 22); fb = ImageFont.truetype(FONTB, 52)
    for i, L in enumerate("ABCDEFG"):
        y = 40 + i * 80
        w = 200 + i * 55
        if L == letter:
            d.rectangle([150, y - 12, 150 + w + 330, y + 72], outline="black", width=4)
            d.rectangle([160, y, 160 + w, y + 60], fill=COLORS[i])
            d.text((175, y + 2), L, fill="white", font=fb)
            d.text((180 + w, y - 4), f"{ep}", fill="black", font=fb)
            d.text((180 + w + 120, y - 4), f"{ges}*", fill="black", font=fb)
            d.text((180 + w, y + 50), "kWh/m²/an", fill="black", font=fs)
            d.text((180 + w + 120, y + 50), "kg CO₂/m²/an", fill="black", font=fs)
        else:
            d.rectangle([160, y, 160 + w, y + 60], fill=COLORS[i])
            d.text((175, y + 8), L, fill="white", font=f)
    b = io.BytesIO(); img.save(b, "PNG"); b.seek(0)
    return ImageReader(b)

def text_lines(c, x, y, lines, size=9, font="DV", lead=12.5):
    c.setFont(font, size)
    for l in lines:
        c.drawString(x, y, l); y -= lead
    return y

def dpe2021(path, v):
    c = canvas.Canvas(path, pagesize=A4)
    # p.1
    c.setFont("DVB", 15); c.drawString(40, H - 50, "DPE   diagnostic de performance énergétique (logement)")
    text_lines(c, 380, H - 80, [f"n° : {v['ademe']}", f"établi le : {v['etabli']}", f"valable jusqu’au : {v['valable']}"], 9)
    text_lines(c, 40, H - 80, [f"adresse : {v['adresse']}", f"type de bien : Appartement", "année de construction : 1948 - 1974",
                               f"surface habitable : {v['surface']} m²", "propriétaire : SCI EXEMPLE", "adresse : 1 place du Marché 35000 RENNES"], 9)
    c.setFont("DVB", 12); c.drawString(40, H - 175, "Performance énergétique et climatique")
    c.drawImage(label_png(v['ep'], v['ges'], v['letter']), 40, H - 470, width=330, height=235)
    text_lines(c, 390, H - 260, ["Ce logement émet " + v['ges_tot'] + " kg de CO₂ par", "an, soit l’équivalent de 4 456 km", "parcourus en voiture."], 9)
    c.setFont("DVB", 12); c.drawString(40, H - 500, "Estimation des coûts annuels d’énergie du logement")
    text_lines(c, 40, H - 518, ["Les coûts sont estimés en fonction des caractéristiques de votre logement et pour une utilisation standard sur 5 usages",
                               "(chauffage, eau chaude sanitaire, climatisation, éclairage, auxiliaires) voir p.3 pour voir les détails par poste."], 7.5, lead=10)
    c.setFont("DVB", 13); c.drawString(150, H - 550, f"entre {v['cmin']} €   et   {v['cmax']} €   par an")
    c.setFont("DV", 7); c.drawString(40, H - 570, "Prix moyens des énergies indexés au 1er janvier 2021 (abonnements compris)")
    c.setFont("DVB", 11); c.drawString(40, H - 600, "Informations diagnostiqueur")
    text_lines(c, 40, H - 618, v['diag_lines'], 9)
    c.showPage()
    # p.2
    c.setFont("DVB", 15); c.drawString(40, H - 50, "DPE   Schéma des déperditions de chaleur")
    text_lines(c, 40, H - 90, ["Performance de l’isolation", "toiture ou plafond 20 %", "murs 35 %", "portes et fenêtres 15 %",
                               "Système de ventilation en place : VMC simple flux autoréglable",
                               "Confort d’été (hors climatisation) : bon"], 9)
    c.showPage()
    # p.3
    c.setFont("DVB", 15); c.drawString(40, H - 50, "DPE   Montants et consommations annuels d’énergie")
    y = text_lines(c, 40, H - 90, ["usage          consommation d’énergie (en kWh énergie primaire)          frais annuels d’énergie (fourchette d’estimation*)"], 8)
    for u, e, ef, a, b in v['usages']:
        c.setFont("DV", 9); c.drawString(40, y, u); c.drawString(150, y, f"{e} ({ef} é.f.)"); c.drawString(380, y, f"entre {a} € et {b} €"); y -= 16
    c.setFont("DVB", 9)
    c.drawString(40, y - 6, "énergie totale pour les usages recensés :"); c.drawString(260, y - 6, f"{v['ep_tot']} kWh ({v['ef_tot']} kWh é.f.)")
    c.drawString(400, y - 6, f"entre {v['cmin']} € et {v['cmax']} € par an")
    c.showPage()
    c.save()

def old2014(path, v):
    c = canvas.Canvas(path, pagesize=A4)
    c.setFont("DVB", 13); c.drawString(40, H - 50, "Diagnostic de performance énergétique - logement (6.A)")
    y = text_lines(c, 40, H - 80, [f"N° : {v['ademe']}", f"Valable jusqu’au : {v['valable']}", "Type de bâtiment : appartement",
        "Année de construction : 1930", f"Surface habitable : {v['surface']} m²", f"Adresse : {v['adresse']}",
        "Date de visite : 05/05/2014", f"Date d’établissement : {v['etabli']}", f"Diagnostiqueur : {v['diag']}", f"Cabinet : {v['cabinet']}"], 9)
    y = text_lines(c, 40, y - 10, ["Consommations annuelles par énergie",
        "obtenues par la méthode 3CL-DPE, version 2012, estimées à l’immeuble / au logement, prix moyens des énergies indexés au 15/08/2015",
        f"Consommations d’énergie pour les usages recensés   {v['ef_tot']}   {v['ep_tot']}   {v['cout']} €"], 8)
    y = text_lines(c, 40, y - 10, [f"Consommation conventionnelle : {v['ep']} kWhEP/m².an", f"Estimation des émissions : {v['ges']} kg éqCO2/m².an",
        "Logement économe                                   Logement énergivore"], 10)
    c.showPage(); c.save()

def scan(src, dst, angle=0.6, dpi=150):
    import tempfile; tmp = tempfile.mkdtemp(prefix="scan_")
    subprocess.run(["pdftoppm", "-r", str(dpi), "-gray", "-png", src, os.path.join(tmp, "p")], check=True)
    pages = []
    for fn in sorted(os.listdir(tmp)):
        im = Image.open(os.path.join(tmp, fn)).convert("L").rotate(angle, expand=False, fillcolor=255, resample=Image.BICUBIC)
        im = im.filter(ImageFilter.GaussianBlur(0.6))
        random.seed(1)
        px = im.load()
        for _ in range(im.width * im.height // 400):
            x, y = random.randrange(im.width), random.randrange(im.height); px[x, y] = random.choice([0, 90, 200])
        p = os.path.join(tmp, fn.replace(".png", ".jpg")); im.save(p, "JPEG", quality=70); pages.append(p)
    with open(dst, "wb") as f:
        subprocess.run(["img2pdf", *pages], check=True, stdout=f)

A = dict(ademe="2275E1234567X", etabli="19/10/2022", valable="18/10/2032", adresse="12 rue Lamarck 75018 PARIS", surface="45,20",
         ep=208, ges=19, letter="D", ges_tot="860", cmin="1 230", cmax="1 700", ep_tot="9 410", ef_tot="4 091",
         usages=[("chauffage", "7 200", "3 130", "960", "1 300"), ("eau chaude", "1 800", "783", "230", "320"), ("éclairage", "410", "178", "40", "80")],
         diag_lines=["ACADE", "37 avenue de la République 75011 PARIS", "diagnostiqueur : Jean DUPONT", "tel : 01 23 45 67 89",
                     "email : contact@acade.fr", "n° de certification : C2019-SE01-045", "organisme de certification : I.Cert"])
B = dict(ademe="2575E0456789B", etabli="15/03/2025", valable="14/03/2035", adresse="8 boulevard Voltaire 75011 PARIS", surface="28,50",
         ep=340, ges=21, letter="E", ges_tot="600", cmin="1 080", cmax="1 470", ep_tot="9 690", ef_tot="4 213",
         usages=[("chauffage", "7 900", "3 435", "880", "1 200"), ("eau chaude", "1 790", "778", "200", "270")],
         diag_lines=["SOCOTEC Diagnostic", "5 place des Frères Montgolfier 78280 GUYANCOURT", "diagnostiqueur : Marie LEROY", "tel : 01 30 12 80 00",
                     "email : diag@socotec.com", "n° de certification : 8524-A", "organisme de certification : Bureau Veritas Certification"])
C = dict(ademe="1475V1001234X", valable="11/05/2024", surface="52", adresse="3 rue de Rivoli 75004 PARIS", etabli="12/05/2014",
         diag="Paul MARTIN", cabinet="DIAG 75", ef_tot="9 500,00", ep_tot="16 588,00", cout="1 120,50", ep="318,99", ges="74,64")
E = dict(A, ademe="2335E0001112K", etabli="02/02/2023", valable="01/02/2033", adresse="40 rue de la Paix 33000 BORDEAUX", surface="70,00",
         ep=150, ges=4, letter="C", ges_tot="300", cmin="700", cmax="1 000", ep_tot="10 500", ef_tot="4 565",
         diag_lines=["DIAG EXPRESS SUD-OUEST", "2 quai des Chartrons 33000 BORDEAUX", "diagnostiqueur : Luc BERNARD", "email : luc@diagexpress.fr",
                     "n° de certification : DTI-2020-12", "organisme de certification : Dekra"])
D = dict(A, ademe="2475E0009998T", etabli="01/01/2024", valable="31/12/2033", adresse="5 rue Haute 69001 LYON", surface="60,00",
         ep=95, ges=3, letter="B", ges_tot="180", cmin="500", cmax="720", ep_tot="5 700", ef_tot="2 478",
         diag_lines=["AC ENVIRONNEMENT", "Agence de Lyon", "diagnostiqueur : Sophie GIRARD", "email : lyon@ac-environnement.com",
                     "n° de certification : B2C-0456", "organisme de certification : B2C"])

dpe2021(os.path.join(OUT, "IM0393664_10_DPE_2022-10-19.pdf"), A)
import tempfile; btxt = os.path.join(tempfile.mkdtemp(), "B_text.pdf"); dpe2021(btxt, B); scan(btxt, os.path.join(OUT, "IM0393664_11_DPE_2025-03-15.pdf"))
old2014(os.path.join(OUT, "IM0393665_3_DPE_2014-05-12.pdf"), C)
dpe2021(os.path.join(OUT, "IM0393667_5_DPE_2023-02-02.pdf"), E)
dpe2021(os.path.join(OUT, "IM0400000_7_DPE_2024-01-01.pdf"), D)
open(os.path.join(OUT, "notes.txt"), "w").write("pas un PDF\n")
print("ok", sorted(os.listdir(OUT)))
