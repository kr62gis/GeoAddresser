import pandas as pd
import re
from difflib import get_close_matches

streets = []
number = []
number_add = []
coord_east = []
coord_north = []
note_of_change = []
addresses_no_hits = []
data_split2=False

def user_input_boolean(text):
    cache = input(text)
    while cache != "y" and cache != "n":
        print("Die Eingabe entspricht nicht der Vorgabe. Versuchen Sie es erneut!")
        cache = input(text)
    if cache == "y":
        cache = True
    else:
        cache = False
    return cache

def user_input_int(text):
    cache = input(text)
    while cache.isdigit() is False:
        print("Die Eingabe entspricht nicht der Vorgabe. Versuchen Sie es erneut!")
        cache = input(text)
    cache = int(cache) - 1
    return cache

header = user_input_boolean("Haben die Spalten Überschriften (y/n): ")
data_split = user_input_boolean("Sind Straße und Hausnummer getrennt? (y/n): ")
if data_split:
    data_split2 = user_input_boolean("Sind Hausnummer und Hausnummer Zusatz getrennt? (y/n): ")
column_street = user_input_int("Spaltennummer der Straßennamen: (z.B.: 2): ")
if data_split:
    column_number = user_input_int("Spaltennummer der Hausnummern: (z.B.: 3): ")
if data_split and data_split2:
    column_number_add = user_input_int("Spaltennummer der Hausnummern Zusätze: (z.B.: 4): ")
get_best_match = user_input_boolean("Soll für den Straßennamen der beste Treffer aus dem Straßenverzeichnis verwendet werden,\nwenn der Straßenname nicht im Straßenverzeichnis enthalten ist? (y/n): ")
if get_best_match:
    print("\nBitte beachten Sie:")
    print("Der beste übereinstimmende Wert muss nicht der gesuchte Wert sein.\nBitte prüfen Sie die erstellte Liste auf Richtigkeit!")

skiprows=0
if header:
    skiprows=1

df = pd.read_excel("data.xlsx", header=None)
df1 = df.dropna(subset=column_street)
df2 = df1.reset_index(drop=True)
df3 = df2[skiprows:]
streets = df3[column_street].tolist()

if data_split:
    number = df3[column_number].astype(str).tolist()
if data_split and data_split2:
    number_add = df3[column_number_add].astype(str).tolist()

df_addresses = pd.read_csv("https://gisdata.krzn.de/files/opendatagis/Stadt_Krefeld/Navi_Geb/Adressen_Stadt_Krefeld_CSV.zip", header=0, sep=";", usecols=["Koordinate_Rechtswert", "Koordinate_Hochwert", "Strassenname", "Hausnummer", "Hausnummer_Zusatz"], encoding='latin-1')
df_addresses_number_add = df_addresses["Hausnummer_Zusatz"].fillna(value="")
df_addresses = df_addresses.astype(str)
df_addresses["Adresse"] = df_addresses["Strassenname"] + " " + df_addresses["Hausnummer"] + " " + df_addresses_number_add.astype(str)
streets_gdw = df_addresses["Strassenname"].unique()
addresses_gdw = df_addresses["Adresse"].tolist()
coord_east_gdw = df_addresses["Koordinate_Rechtswert"].tolist()
coord_north_gdw = df_addresses["Koordinate_Hochwert"].tolist()

#Seperate data
def seperate(expression, param_reference, param_extract):
    for i in range(0, len(param_reference)):
        re_result = re.search(expression, param_reference[i])
        if re_result:
            param_extract.append(param_reference[i][re_result.start():].strip())
            param_reference[i] = param_reference[i][:re_result.start()].strip()
        else:
            param_extract.append("")

if data_split == False and data_split2 == False:
    seperate("\D+$", streets, number_add)
    seperate("\d", streets, number)
elif data_split and data_split2 == False:
    seperate("\D+$", number, number_add)

#Improve number_add especially to ignore long texts
for i in range(0, len(number_add)):
    number_add[i] = number_add[i].lower()
    if len(number_add[i]) < 3:
        continue
    re_result = re.search("[^a-zA-Z][a-zA-Z][^a-zA-Z]", number_add[i])
    if re_result:
        number_add[i] = number_add[i][re_result.start():re_result.end()]
    else:
        number_add[i] = ""

#Extract first match
def first_match(expression, param):
    for i in range(0, len(param)):
        re_result = re.search(expression, param[i])
        if re_result:
            param[i] = re_result[0]
        else:
            param[i] = ""

first_match("[a-z]", number_add)
first_match("\d{1,3}", number)

#Get best close match for streets in streets_gdw
if get_best_match:
    for i in range(0, len(streets)):
        if streets[i] not in streets_gdw:
            best_hit = get_close_matches(streets[i], streets_gdw, cutoff=0.7)
            note_of_change.append("**")
            if best_hit:
                streets[i] = best_hit[0]
        else:
            note_of_change.append("")

#Test if address in addresses_gdw
for i in range(0, len(streets)):
    address = streets[i] + " " + number[i] + " " + number_add[i]
    if address not in addresses_gdw:
        addresses_no_hits.append("***")
        coord_east.append("")
        coord_north.append("")
    else:
        addresses_no_hits.append("")
        j=addresses_gdw.index(address)
        coord_east.append(coord_east_gdw[j])
        coord_north.append(coord_north_gdw[j])

if header:
    streets.insert(0, "STRASSENNAME")
    number.insert(0, "HAUS_NR")
    number_add.insert(0, "HAUS_NR_ZUSATZ")
    coord_east.insert(0, "Koord_Rechtswert")
    coord_north.insert(0, "Koord_Hochwert")
    if get_best_match:
        note_of_change.insert(0, "Änderungshinweis")
    addresses_no_hits.insert(0, "Adresse_nicht_gefunden")

df2[len(df2.columns)] = streets
df2[len(df2.columns)] = number
df2[len(df2.columns)] = number_add
df2[len(df2.columns)] = coord_east
df2[len(df2.columns)] = coord_north
if get_best_match:
    df2[len(df2.columns)] = note_of_change
df2[len(df2.columns)] = addresses_no_hits

df2.to_csv("data_geo.csv", index=False, header=None)

input("\nDrücken Sie eine beliebige Taste zum Beenden!")