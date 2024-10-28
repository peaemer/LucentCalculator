import requests
import pandas as pd
import json

def update_xlsx_file():
    # API-Abfrage senden
    url = 'https://questlog.gg/throne-and-liberty/api/trpc/actionHouse.getAuctionHouse?input={"language":"en-nc","regionId":"eu-e"}'
    response = requests.get(url)

    # Überprüfen, ob die Anfrage erfolgreich war
    if response.status_code == 200:
        data = response.json()
        print("Raw API response:", json.dumps(data, indent=4))  # JSON-Antwort für Debugging ausgeben

        # Prüfen, ob die erwarteten Schlüssel existieren
        if 'result' in data and 'data' in data['result'] and len(data['result']['data']) > 0:
            items = data['result']['data']  # Liste der Items
            rows = []  # Sammlung der Einträge für das DataFrame

            for item in items:
                # Item-Details extrahieren
                item_details = {
                    'id': item['id'],
                    'name': item['name'],
                    'icon': item['icon'],
                    'grade': item['grade'],
                    'mainCategory': item['mainCategory'],
                    'subCategory': item['subCategory'],
                    'minPrice': item['minPrice'],
                    'inStock': item['inStock'],
                }

                # Letzter Handel, falls vorhanden
                if 'traitItems' in item and len(item['traitItems']) > 0:
                    latest_trade = item['traitItems'][0]  # Annahme: neuester Handel ist der erste in der Liste
                    item_details.update({
                        'latest_trade_traitId': latest_trade['traitId'],
                        'latest_trade_minPrice': latest_trade['minPrice'],
                        'latest_trade_inStock': latest_trade['inStock']
                    })

                # Fügen Sie die Item-Details zu den Zeilen hinzu
                rows.append(item_details)

            # DataFrame erstellen und in Excel-Datei speichern
            df = pd.DataFrame(rows)
            df.to_excel('latest_item_histories_output.xlsx', index=False)  # Exportieren in Excel-Datei
            print("Daten erfolgreich in 'latest_item_histories_output.xlsx' exportiert")
        else:
            print("Keine Items in der Antwort verfügbar.")
    else:
        print(f"Fehler beim Abrufen der Daten: {response.status_code}")

def get_item_price(item_name):
    # Excel-Datei laden
    try:
        df = pd.read_excel('latest_item_histories_output.xlsx')

        # Nach Item-Name filtern und den Preis abrufen
        item = df[df['name'] == item_name]
        if not item.empty:
            min_price = item['minPrice'].values[0]
            print(f"Der Mindestpreis für {item_name} ist {min_price}")
            return min_price
        else:
            print(f"Das Item '{item_name}' wurde nicht gefunden.")
            return None
    except FileNotFoundError:
        print("Excel-Datei wurde nicht gefunden. Bitte führen Sie zuerst das Update durch.")
        return None

def main():
    update_xlsx_file()  # Aktualisiert die Excel-Datei
    get_item_price(input("Wie ist der name des Items?/Whats the name of the item?: "))  # Beispielabfrage nach dem Preis eines spezifischen Items

if __name__ == "__main__":
    main()
