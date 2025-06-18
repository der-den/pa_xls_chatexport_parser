import pandas as pd
import sys
from pathlib import Path

def analyze_excel(excel_file):
    """
    Analysiert eine Excel-Datei und gibt alle Blätter und Spalten zurück
    """
    try:
        # Excel-Datei einlesen
        print(f"Lese Excel-Datei: {excel_file}")
        excel = pd.ExcelFile(excel_file)
        
        # Alle Blätter (Sheets) auflisten
        sheets = excel.sheet_names
        print(f"Gefundene Blätter: {sheets}")
        
        result = []
        result.append(f"Excel-Datei: {Path(excel_file).name}")
        result.append("=" * 50)
        
        # Für jedes Blatt die Spalten auflisten
        for sheet in sheets:
            print(f"Analysiere Blatt: {sheet}")
            # Lese die ersten beiden Zeilen separat
            df_header = pd.read_excel(excel_file, sheet_name=sheet, nrows=2)
            
            # Lese die gesamte Tabelle für die Datenanalyse
            df = pd.read_excel(excel_file, sheet_name=sheet)
            
            # Prüfe, ob die zweite Zeile die tatsächlichen Spaltenbezeichner enthält
            has_header_in_second_row = False
            if len(df) >= 2:
                # Extrahiere die zweite Zeile als potenzielle Spaltenbezeichner
                second_row = df.iloc[1].tolist()
                if any(isinstance(val, str) for val in second_row if pd.notna(val)):
                    has_header_in_second_row = True
            
            # Spaltenüberschriften aus der ersten Zeile
            columns = df.columns.tolist()
            
            result.append(f"\nBlatt: {sheet}")
            result.append("-" * 30)
            result.append(f"Anzahl Zeilen: {len(df)}")
            result.append(f"Anzahl Spalten: {len(columns)}")
            
            # Wenn Spaltenbezeichner in der zweiten Zeile sind
            if has_header_in_second_row:
                result.append("\nHinweis: Spaltenbezeichner wurden in der zweiten Zeile gefunden!")
                result.append("\nSpaltenbezeichner (aus zweiter Zeile):")
                second_row_headers = df.iloc[0].tolist()
                for i, header in enumerate(second_row_headers):
                    if pd.isna(header):
                        header = f"[Leer {i+1}]"
                    result.append(f"{i+1}. {header}")
            
            result.append("\nSpalten (mit Original-Namen):")
            
            # Füge jede Spalte mit Beispielwerten hinzu
            for i, col in enumerate(columns):
                # Versuche, einen nicht-leeren Beispielwert zu finden
                # Überspringe die ersten beiden Zeilen, da diese Header sein könnten
                if len(df) > 2:
                    sample_values = df.iloc[2:][col].dropna().head(3).tolist()
                else:
                    sample_values = df[col].dropna().head(3).tolist()
                    
                sample_str = ", ".join([str(val) for val in sample_values])
                if len(sample_str) > 100:
                    sample_str = sample_str[:97] + "..."
                
                result.append(f"{i+1}. {col} - Beispielwerte: {sample_str}")
            
            # Füge eine Leerzeile für bessere Lesbarkeit hinzu
            result.append("")
        
        return "\n".join(result)
        
    except Exception as e:
        return f"Fehler beim Analysieren der Excel-Datei: {str(e)}"

if __name__ == "__main__":
    # Prüfe, ob ein Dateiname als Argument übergeben wurde
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        # Standardmäßig "Report.xlsx" im aktuellen Verzeichnis verwenden
        excel_file = "Report.xlsx"
    
    # Analysiere die Excel-Datei
    result = analyze_excel(excel_file)
    
    # Speichere das Ergebnis in einer Textdatei
    output_file = "excel_struktur.txt"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(result)
    
    print(f"Analyse abgeschlossen. Ergebnisse wurden in '{output_file}' gespeichert.")
