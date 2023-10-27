import pandas as pd
import folium
import requests
import asyncio
from datetime import datetime, timedelta
from shapely.geometry import Polygon, MultiPolygon
from shapely.ops import unary_union
from geopy.geocoders import Nominatim
from traveltimepy import PublicTransport, Coordinates, TravelTimeSdk

def create_legend():
    legend_html = '''
    <div style="position: fixed; bottom: 50px; left: 50px; z-index:9999; background-color: rgba(255, 255, 255, 0.8);
     border-radius:6px; padding: 10px; font-size:14px; width: 200px;">
    &nbsp;<b>Légende</b><br>
    &nbsp;<i class="fa fa-square fa-1x" style="color:blue"></i>&nbsp;Transports publics<br>
    &nbsp;<i class="fa fa-square fa-1x" style="color:white"></i>&nbsp;Marche<br>
    &nbsp;<i class="fa fa-square fa-1x" style="color:yellow"></i>&nbsp;Vélo<br>
    &nbsp;<i class="fa fa-square fa-1x" style="color:orange"></i>&nbsp;VAE (25 km/h)<br>
    &nbsp;<i class="fa fa-square fa-1x" style="color:red"></i>&nbsp;Speedelec (45 km/h)<br>
    &nbsp;<i class="fa fa-square fa-1x" style="color:black"></i>&nbsp;Voiture / Motocycle<br>
    </div>
    '''
    return legend_html

# Charger le fichier "Input_2.xlsx" pour les adresses des lieux de travail
df_travail = pd.read_excel("/Users/quentinschneiter/Desktop/Code/Mobility Plan API/Carte/Input_2.xlsx", usecols=["Adresse"])
df_travail.rename(columns={"Adresse": "adresse_travail"}, inplace=True)

# Charger le fichier "Input_1.xlsx" pour les adresses des employés
df_employe = pd.read_excel("/Users/quentinschneiter/Desktop/Code/Mobility Plan API/Tableau/Input_1.xlsx", usecols=["Adresse"])
df_employe.rename(columns={"Adresse": "adresse_employe"}, inplace=True)

API_KEY_GEOCODE = "YOUR_GOOGLE_GEOCODING_API_KEY"
API_KEY_MAPBOX = "YOUR_MAPBOX_API_KEY"

def geocode(address):
    url = f"https://maps.googleapis.com/maps/api/geocode/json?address={address}&key=AIzaSyBc5vBnBylv7zDiQyVqrF5j1nQOXEVWupE"
    response = requests.get(url).json()
    if response['status'] == 'OK':
        location = response['results'][0]['geometry']['location']
        return location['lat'], location['lng']
    return None, None

# Extraire les coordonnées géographiques pour les adresses des employés
df_employe['lat_employe'], df_employe['lng_employe'] = zip(*df_employe['adresse_employe'].apply(geocode))

# Coordonnées géographiques pour le lieu de travail
df_travail['lat_travail'], df_travail['lng_travail'] = zip(*df_travail['adresse_travail'].apply(geocode))

# Création des isochrones
def get_isochrone(lat, lng, mode, minutes):
    if mode == "VAE":
        mode = "cycling"
        minutes = int(minutes * (22/16))
    elif mode == "Speedelec":
        mode = "cycling"
        minutes = int(minutes * (32/16))
    url = f"https://api.mapbox.com/isochrone/v1/mapbox/{mode}/{lng},{lat}?contours_minutes={minutes}&access_token=pk.eyJ1IjoicXVlbnRpbnNjaG5laXRlciIsImEiOiJjbGxxa29xNDQwZ2R6M2xucXlhbDJpZDFoIn0.MzoZyf-e4PuuHrDulWpPPA"
    response = requests.get(url).json()
    if 'features' in response:
        return response['features'][0]['geometry']['coordinates']
    else:
        print(f"Error with {mode} for {lat}, {lng}: {response.get('message', 'Unknown error')}")
        return []

color_map = {
    "driving": "black",
    "cycling": "yellow",
    "VAE": "orange",
    "Speedelec": "red",
    "walking": "white"
}

modes = ["driving", "cycling", "VAE", "Speedelec", "walking"]

async def travel_time_isochrones(address):
    sdk = TravelTimeSdk("1a7fcb1c", "9b1251ac76e41ca60f37b752c992f204")
    
    start_time = datetime(year=2023, month=11, day=13, hour=7, minute=30)
    end_time = datetime(year=2023, month=11, day=13, hour=8, minute=30)

    geolocator = Nominatim(user_agent="quentin.schneiter@gmail.com")
    location = geolocator.geocode(address)
    if location:
        lat, lng = location.latitude, location.longitude
        polygons = []

        current_time = start_time
        while current_time <= end_time:
            results = await sdk.union_async(
                coordinates=[Coordinates(lat=lat, lng=lng)],
                arrival_time=current_time,
                transportation=PublicTransport(walking_time=1800),
                travel_time=1800
            )
            if results and results.shapes:
                for shape in results.shapes:
                    coords = [(c.lng, c.lat) for c in shape.shell]
                    polygons.append(Polygon(coords))

            current_time += timedelta(minutes=15)

        return unary_union(polygons)

    return None

async def main():

    # Création de la carte
    for _, row in df_travail.iterrows():
        m = folium.Map(location=[row['lat_travail'], row['lng_travail']], zoom_start=13)

     # 2. Ajoutez la légende à la carte
        legend = folium.Element(create_legend())
        m.get_root().html.add_child(legend)

        # Traitement des isochrones du premier script
        for mode in modes:
            coords = get_isochrone(row['lat_travail'], row['lng_travail'], mode, 30)
            if coords:
                folium.Polygon(locations=[(coord[1], coord[0]) for coord in coords], color=color_map[mode], fill=True, fill_opacity=0.2).add_to(m)

        # Traitement des isochrones du deuxième script
        merged_polygon = await travel_time_isochrones(row['adresse_travail'])
        if merged_polygon:
            if isinstance(merged_polygon, Polygon):
                folium.Polygon([(y, x) for x, y in merged_polygon.exterior.coords], color="blue", fill=True, fill_color="blue").add_to(m)
            elif isinstance(merged_polygon, MultiPolygon):
                for poly in merged_polygon.geoms:
                    folium.Polygon([(y, x) for x, y in poly.exterior.coords], color="blue", fill=True, fill_color="blue").add_to(m)
                    
        # Ajouter une punaise noire pour le lieu de travail
        folium.Marker([row['lat_travail'], row['lng_travail']], 
                  icon=folium.Icon(color='black'),
                  popup=row['adresse_travail'],
                  tooltip=row['adresse_travail']).add_to(m)

        # Ajouter des punaises vertes pour les adresses des employés
        for _, emp_row in df_employe.iterrows():
            if not pd.isna(emp_row['lat_employe']) and not pd.isna(emp_row['lng_employe']):
                folium.Marker([emp_row['lat_employe'], emp_row['lng_employe']], 
                              icon=folium.Icon(color='green'),
                              popup=emp_row['adresse_employe'],
                              tooltip=emp_row['adresse_employe']).add_to(m)

        # Sauvegardez la carte au format .html
        safe_filename = "".join([c if c.isalnum() or c.isspace() else "_" for c in row['adresse_travail']])
        m.save(f"/Users/quentinschneiter/Desktop/Code/Mobility Plan API/Carte/{safe_filename}.html")

asyncio.run(main())
