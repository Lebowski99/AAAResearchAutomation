import os
import zipfile
import tkinter as tk
from tkinter import filedialog
import xml.etree.ElementTree as ET

import simplekml
from shapely.geometry import Point, Polygon
from tqdm import tqdm

def select_kmz_file():
    """Opens a file dialog to select a KMZ file."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title="Select KMZ File",
        filetypes=[("KMZ Files", "*.kmz")]
    )
    return file_path

def parse_kml_coordinates(coord_string):
    """
    Parse a KML-style <coordinates> string of:
      lon,lat[,alt] lon,lat[,alt] ...
    Return a list of (lon, lat) float tuples (ignoring altitude).
    """
    coords = []
    coord_string = coord_string.strip()
    if not coord_string:
        return coords
    for chunk in coord_string.split():
        parts = chunk.split(',')
        lon = float(parts[0])
        lat = float(parts[1])
        coords.append((lon, lat))
    return coords

def extract_polygons_from_kmz(kmz_file_path):
    """
    Extract polygons from a KMZ by:
      1. Unzipping the file to read the KML.
      2. Using xml.etree.ElementTree to find <Placemark><Polygon> elements.
      3. Parsing <coordinates> into Shapely polygons.

    Returns a list of (polygon_name, shapely Polygon).
    """
    polygons = []
    try:
        # 1) Unzip the KMZ and read the .kml
        with zipfile.ZipFile(kmz_file_path, "r") as kmz:
            # Typically, there's a doc.kml or similar
            kml_files = [f for f in kmz.namelist() if f.endswith(".kml")]
            if not kml_files:
                print("❌ No .kml file found inside the KMZ.")
                return polygons

            kml_name = kml_files[0]  # just pick the first .kml
            with kmz.open(kml_name) as kml_f:
                tree = ET.parse(kml_f)
                root = tree.getroot()

        # 2) Our KML is in the standard namespace
        ns = {"kml": "http://www.opengis.net/kml/2.2"}

        # Find all placemarks
        placemarks = root.findall(".//kml:Placemark", ns)
        if not placemarks:
            print("❌ Found no <Placemark> elements in KML.")
            return polygons

        # 3) For each Placemark, see if there's a <Polygon>
        for pm in placemarks:
            name_elem = pm.find("kml:name", ns)
            placemark_name = name_elem.text if name_elem is not None else "Unnamed Polygon"

            # A single Placemark might contain multiple <Polygon> elements
            poly_elems = pm.findall(".//kml:Polygon", ns)
            for poly_elem in poly_elems:
                # Typically coordinates appear under <outerBoundaryIs><LinearRing><coordinates>
                coords_elem = poly_elem.find(".//kml:outerBoundaryIs//kml:coordinates", ns)
                if coords_elem is not None and coords_elem.text:
                    outer_coords = parse_kml_coordinates(coords_elem.text)
                    if outer_coords:
                        polygon = Polygon(outer_coords)
                        polygons.append((placemark_name, polygon))
    except Exception as e:
        print(f"❌ Error extracting polygons: {e}")

    return polygons

def create_filtered_kmz(polygon, polygon_name, output_filename="filtered_degree_squares.kmz"):
    """
    Creates a KMZ file with markers only inside the selected polygon.
    Loops over every integer lat/lon worldwide to see if it's contained by the polygon.
    """
    if not polygon:
        print("❌ No valid polygon selected.")
        return

    print(f"🚀 Generating markers inside: {polygon_name}")

    kml_obj = simplekml.Kml()
    icon_url = "http://maps.google.com/mapfiles/kml/shapes/donut.png"
    unique_points = set()

    # Loop through integer degree coordinates globally
    for lat in range(-90, 91):   # Latitude from -90 to 90
        for lon in range(-180, 181):  # Longitude from -180 to 180
            point = Point(lon, lat)
            if polygon.contains(point):
                unique_points.add((lat, lon))

    progress = tqdm(total=len(unique_points), desc="Generating KMZ", unit="marker")

    # Add markers inside the polygon
    for lat, lon in unique_points:
        pnt = kml_obj.newpoint(
            name=f"{lat}N {lon}E",
            coords=[(lon, lat)]
        )
        pnt.style.iconstyle.icon.href = icon_url
        progress.update(1)

    progress.close()
    print("✅ Marker generation complete! Now saving to KML file...")

    # Save to KML
    temp_kml = "temp.kml"
    kml_obj.save(temp_kml)
    print("✅ KML file saved successfully! Now converting to KMZ...")

    # Convert the KML to KMZ
    try:
        with zipfile.ZipFile(output_filename, 'w', zipfile.ZIP_DEFLATED) as kmz:
            kmz.write(temp_kml, arcname="doc.kml")
        os.remove(temp_kml)
        print(f"🎉 KMZ file successfully created: {os.path.abspath(output_filename)}")
    except Exception as e:
        print(f"❌ Error saving KMZ file: {e}")

if __name__ == "__main__":
    # Ask user for a KMZ file
    kmz_path = select_kmz_file()
    if kmz_path:
        # Extract polygons via ElementTree
        polygons = extract_polygons_from_kmz(kmz_path)
        if polygons:
            print("\nAvailable Polygons:")
            for i, (name, _) in enumerate(polygons, start=1):
                print(f"{i}. {name}")

            # Ask user to choose which polygon to use
            choice = input("\nEnter the number of the polygon to use: ")
            if choice.isdigit():
                choice_idx = int(choice) - 1
                if 0 <= choice_idx < len(polygons):
                    selected_name, selected_polygon = polygons[choice_idx]
                    create_filtered_kmz(selected_polygon, selected_name)
                else:
                    print("❌ Invalid choice. Exiting.")
            else:
                print("❌ Please enter a valid number.")
        else:
            print("❌ No polygons found in the KMZ file.")
    else:
        print("❌ No file selected. Exiting.")

    input("\nPress Enter to exit...")
