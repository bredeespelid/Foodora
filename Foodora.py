import tkinter as tk
from tkinter import filedialog, messagebox
from langchain_community.document_loaders import PyPDFLoader
import pandas as pd
import re
import os
from decimal import Decimal, ROUND_HALF_UP

def extract_amount(text):
    match = re.search(r'(-?\d{1,3}(?:[\s\xa0]\d{3})*(?:,\d{2})?)\s*NOK', text)
    if match:
        cleaned = match.group(1).replace('\xa0', '').replace(' ', '').replace(',', '.')
        return float(cleaned)
    return None

def calculate_difference(data):
    amount_99999905 = sum(Decimal(str(x['Beløp'])) for x in data if x['Konto'] == 99999905)
    amount_7210 = sum(Decimal(str(x['Beløp'])) for x in data if x['Konto'] == 7210)
    amount_6551 = sum(Decimal(str(x['Beløp'])) for x in data if x['Konto'] == 6551)
    amount_3066 = sum(Decimal(str(x['Beløp'])) for x in data if x['Konto'] == 3066)
    
    difference = amount_99999905 + (amount_7210 + amount_6551) - amount_3066
    return difference.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

def process_pdf(file_path):
    loader = PyPDFLoader(file_path)
    pages = loader.load_and_split()[:1]
    content = pages[0].page_content

    selger = re.search(r'Selger: Godt Brød - (.+)', content)
    dato = re.search(r'Fakturadato: (\d{2}\.\d{2}\.\d{4})', content)
    totalsalg = re.search(r'Ditt totalsalg inkl\. MVA\nTotalt \( 1 \) (.+)', content)
    vi_betaler = re.search(r'Vi betaler til deg \( 1 \) \+ \( 2 \) (.+)', content)
    sanctions = re.search(r'Sanctions\s+\d+\s+(-?\d{1,3}(?:[\s\xa0]\d{3})*(?:,\d{2})?)\s*NOK', content)
    hardware = re.search(r'Hardware\s+\d+\s+(-?\d{1,3}(?:[\s\xa0]\d{3})*(?:,\d{2})?)\s*NOK', content)

    data = []

    if vi_betaler:
        data.append({
            'Selger': selger.group(1) if selger else None,
            'Dato': dato.group(1) if dato else None,
            'Konto': 99999905,
            'Beløp': extract_amount(vi_betaler.group(0))
        })
    if totalsalg:
        data.append({
            'Selger': selger.group(1) if selger else None,
            'Dato': dato.group(1) if dato else None,
            'Konto': 3066,
            'Beløp': extract_amount(totalsalg.group(0))
        })

    if sanctions:
        sanctions_amount = extract_amount(sanctions.group(0))
        if sanctions_amount is not None:
            data.append({
                'Selger': selger.group(1) if selger else None,
                'Dato': dato.group(1) if dato else None,
                'Konto': 7210,
                'Beløp': abs(sanctions_amount)
            })
    if hardware:
        hardware_amount = extract_amount(hardware.group(0))
        if hardware_amount is not None:
            data.append({
                'Selger': selger.group(1) if selger else None,
                'Dato': dato.group(1) if dato else None,
                'Konto': 6551,
                'Beløp': abs(hardware_amount)
            })

    difference = calculate_difference(data)
    if abs(difference) >= Decimal('0.01'):
        data.append({
            'Selger': selger.group(1) if selger else None,
            'Dato': dato.group(1) if dato else None,
            'Konto': 7740,
            'Beløp': float(difference)
        })

    return data

def main():
    root = tk.Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(
        title="Velg PDF-filer",
        filetypes=[("PDF files", "*.pdf")]
    )

    if not file_paths:
        messagebox.showinfo("Informasjon", "Ingen filer valgt. Programmet avsluttes.")
        return

    all_data = []
    for file_path in file_paths:
        try:
            file_data = process_pdf(file_path)
            all_data.extend(file_data)
            print(f"Prosessert {os.path.basename(file_path)}: {len(file_data)} rader funnet.")
        except Exception as e:
            messagebox.showerror("Feil", f"Feil ved prosessering av {os.path.basename(file_path)}: {str(e)}")

    if not all_data:
        messagebox.showinfo("Informasjon", "Ingen data funnet i de valgte filene.")
        return

    output_file = filedialog.asksaveasfilename(
        title="Lagre Excel-fil",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not output_file:
        messagebox.showinfo("Informasjon", "Ingen lagringsdestinasjon valgt. Programmet avsluttes.")
        return

    df = pd.DataFrame(all_data)
    df.to_excel(output_file, index=False)
    messagebox.showinfo("Suksess", f"Data er lagret i {output_file}")

    print("\nOppsummering av innsamlede data:")
    print(df)

if __name__ == "__main__":
    main()
