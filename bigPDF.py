# Import Libraries
import io
import os
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import messagebox

import fitz  # PyMuPDF
from fitz import Page
from pdf2docx import Converter
from PIL import Image
from PyPDF4 import PdfFileMerger, PdfFileReader, PdfFileWriter, utils


def convert_pdf_2_docx():
    dateien = fd.askopenfilenames(filetypes=(("PDF", "*.pdf"), ("Alle", "*.*")))

    for datei in dateien:
        output = datei.replace(".pdf", "") + "_converted.docx"
        try:
            cv = Converter(datei)
            cv.convert(output)
            cv.close()
        except ValueError:
            message = "Datei verschlüsselt" + datei
            messagebox.showwarning(message=message)
            return


def merge_pdfs(page_range: tuple | None = None, bookmark: bool = True):
    dateien = fd.askopenfilenames(filetypes=(("PDF", "*.pdf"), ("Alle", "*.*")))
    try:
        output = dateien[0].replace(".pdf", "") + "_combined.pdf"
    except IndexError:
        message = "Bitte eine Datei auswählen"
        messagebox.showwarning(message=message)
        return

    merger = PdfFileMerger(strict=False)
    for datei in dateien:
        bookmark_name = (
            os.path.splitext(os.path.basename(datei))[0] if bookmark else None
        )
        merger.append(
            fileobj=open(datei, "rb"),
            pages=page_range,
            import_bookmarks=False,
            bookmark=bookmark_name,
        )
    # Insert the pdf at specific page
    merger.write(fileobj=open(output, "wb"))
    merger.close()


def get_images():
    # file path you want to extract images from
    dateien = fd.askopenfilenames(filetypes=(("PDF", "*.pdf"), ("Alle", "*.*")))
    try:
        pfad = dateien[0]
    except IndexError:
        message = "Bitte eine Datei auswählen"
        messagebox.showwarning(message=message)
        return
    ordner_pfad = ""
    for txt in pfad.split(sep="/")[:-1]:
        ordner_pfad += txt + "/"

    os.chdir(ordner_pfad)
    os.mkdir("Gefundene_Bilder")
    bilder_pfad = ordner_pfad + "Gefundene_Bilder"
    os.chdir(bilder_pfad)
    # open the file
    for datei in dateien:
        pdf_file = fitz.open(datei)
        # iterate over PDF pages
        for page_index in range(len(pdf_file)):
            # get the page itself
            page = pdf_file[page_index]
            image_list = page.get_images()
            # printing number of images found in this page
            if image_list:
                print(
                    f"[+] Found a total of {len(image_list)} images in page {page_index}"
                )
            else:
                print("[!] No images found on page", page_index)
            for image_index, img in enumerate(page.get_images(), start=1):
                # get the XREF of the image
                xref = img[0]
                # extract the image bytes
                base_image = pdf_file.extract_image(xref)
                image_bytes = base_image["image"]
                # get the image extension
                image_ext = base_image["ext"]
                # load it to PIL
                image = Image.open(io.BytesIO(image_bytes))
                # save it to local disk
                with open(
                    f"image{page_index+1}_{image_index}.{image_ext}", "wb"
                ) as bild:
                    image.save(bild)


def rotate_pages(eingabe_text):
    """Dreht Seiten einer PDF und speichert das neue Dokument ab.
    Seitendrehen in entry bspw:: 0:r, 1:r, 2:l, 5:r, alle:l"""
    datei = fd.askopenfilename(filetypes=(("PDF", "*.pdf"), ("Alle", "*.*")))
    zu_drehende_seiten_raw = drehen_eingabe.get()

    ordner_pfad = ""
    for txt in datei.split(sep="/")[:-1]:
        ordner_pfad += txt + "/"

    zu_drehende_seiten = {}
    try:
        pdf_writer = PdfFileWriter()
        pdf_reader = PdfFileReader(datei)
        anzahl_seiten = pdf_reader.getNumPages()
        for txt in zu_drehende_seiten_raw.split(","):
            seiten_zahl, richtung = txt.split(":")
            if seiten_zahl == "alle":
                for seite in range(anzahl_seiten):
                    zu_drehende_seiten.update({seite: richtung})
                break
            else:
                zu_drehende_seiten.update({int(seiten_zahl): richtung})
    except FileNotFoundError:
        message = "Bitte eine Datei auswählen"
        messagebox.showwarning(message=message)
        return
    except ValueError:
        message = "Bitte zu drehende Seiten eingeben"
        messagebox.showwarning(message=message)
        return

    for key in zu_drehende_seiten.keys():
        if zu_drehende_seiten[key] == "r":
            zu_drehende_seiten[key] = 90
        elif zu_drehende_seiten[key] == "l":
            zu_drehende_seiten[key] = -90
        elif zu_drehende_seiten[key] == "u":
            zu_drehende_seiten[key] = 180
    # seiten drehen
    for seite in range(anzahl_seiten):
        if seite in zu_drehende_seiten.keys():
            page = pdf_reader.getPage(seite).rotateClockwise(zu_drehende_seiten[seite])
        else:
            page = pdf_reader.getPage(seite)
        pdf_writer.addPage(page)

    os.chdir(ordner_pfad)
    with open("gedrehtesDokument.pdf", "wb") as fh:
        pdf_writer.write(fh)

    eingabe_text.set("Fertig")


def decrypt(passwort_text):
    dateien = fd.askopenfilenames(filetypes=(("PDF", "*.pdf"), ("Alle", "*.*")))
    passwort = passwort_text.get()

    try:
        pfad = dateien[0]
    except IndexError:
        message = "Bitte eine Datei auswählen"
        messagebox.showwarning(message=message)
        return
    except KeyError:
        message = "Bitte Passwort eingeben"
        messagebox.showwarning(message=message)
        return

    ordner_pfad = ""
    for txt in pfad.split(sep="/")[:-1]:
        ordner_pfad += txt + "/"

    datei_name = ""
    for txt in pfad.split(sep="/")[-1]:
        datei_name += txt

    for datei in dateien:
        output = datei.replace(".pdf", "") + "_decrypted.pdf"
        pdf_reader = PdfFileReader(datei)

        try:
            pdf_reader.decrypt(password=passwort)
            pdf_writer = PdfFileWriter()

            os.chdir(ordner_pfad)
            for p in range(pdf_reader.getNumPages()):
                pdf_writer.addPage(pdf_reader.getPage(p))

            with open(output, "wb") as file:
                pdf_writer.write(file)
        except utils.PdfReadError:
            passwort_text.set("Falsches Passwort für: " + datei_name)
            return
        passwort_text.set("FERTIG")


root = tk.Tk()
root.title("bigPDF")
root.geometry("800x300")

frame = tk.Frame(root, bd=5)
frame.pack(fill="both")

convert_button = tk.Button(frame, text="Zu Word umwandeln", command=convert_pdf_2_docx)
convert_button.pack(fill="x")

merge_button = tk.Button(frame, text="Zusammenfügen", command=merge_pdfs)
merge_button.pack(fill="x")

bild_button = tk.Button(frame, text="Bilder extrahieren", command=get_images)
bild_button.pack(fill="x")

eingabe_text = tk.StringVar(value="")
drehen_button = tk.Button(
    frame, text="Seiten drehen", command=lambda: rotate_pages(eingabe_text)
)
drehen_button.pack(fill="x")
drehen_eingabe = tk.Entry(frame, textvariable=eingabe_text)
drehen_eingabe.pack(fill="x")

passwort_text = tk.StringVar(value="")
decrypt_button = tk.Button(
    frame, text="Entschlüsseln", command=lambda: decrypt(passwort_text)
)
decrypt_button.pack(fill="x")
passwort_eingabe = tk.Entry(frame, textvariable=passwort_text)
passwort_eingabe.pack(fill="x")


root.mainloop()
