#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import sys
from functions import read_excel_file, generate_statistics


def main():
    """
    Hauptfunktion: Liest eine Excel-Datei ein und generiert Statistiken.
    Andere Funktionen wie PDF-Generierung sind nicht enthalten.
    """
    # Kommandozeilenargumente parsen
    parser = argparse.ArgumentParser(description='Excel Chat Export Statistik Generator')
    parser.add_argument('excel_file', help='Pfad zur Excel-Datei')
    parser.add_argument('-v', '--verbose', action='store_true', help='Ausführliche Ausgabe')
    args = parser.parse_args()
    
    # Überprüfe, ob die Excel-Datei angegeben wurde
    if not args.excel_file:
        print("Fehler: Keine Excel-Datei angegeben.")
        parser.print_help()
        sys.exit(1)
    
    # Lese die Excel-Datei mit der Funktion aus excel_reader
    df, metadata = read_excel_file(args.excel_file)
    
    if df is None:
        print("Fehler: Die Excel-Datei konnte nicht gelesen werden.")
        sys.exit(1)
    
    # Zeige Statistiken an
    stats = generate_statistics(df, metadata, args.excel_file, args.verbose)
    print(stats)

if __name__ == "__main__":
    main()
