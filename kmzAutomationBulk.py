import os
import re
import zipfile
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from xml.etree import ElementTree as ET

# ======================================================
# KMZ OVERLAY BULK AUTOMATION
# Create <Document> if missing, & alphabetical .kmz order
# ======================================================
#
# This script processes multiple subfolders, each containing:
#   - Exactly one .kmz ("original overlay" with a single GroundOverlay),
#   - Several .png overlays.
# Produces a new .kmz named "<SubfolderName> - transparent.kmz", with:
#   - All overlays duplicated from the single original overlay,
#   - Overlays referencing <Icon><href> in a subfolder named <SubfolderName>,
#   - <Document> (or <Folder>) named <SubfolderName>,
#   - Overlays sorted by trailing numeric portion in <name>,
#   - Final .kmz files stored in alphabetical order.
#
# Additionally, if there's no <Document> or <Folder> in the original .kml,
# we create one automatically.

def extract_kmz(kmz_path, extract_to):
    """Extracts the .kmz file into a temporary directory."""
    with zipfile.ZipFile(kmz_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)

def create_kmz(source_dir, kmz_path):
    """
    Zips the modified directory back into a .kmz file,
    adding files in alphabetical order.
    """
    with zipfile.ZipFile(kmz_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        # Walk the tree in alphabetical order
        for root, dirs, files in os.walk(source_dir):
            dirs.sort()
            files.sort()
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, start=source_dir)
                zip_ref.write(file_path, arcname)

def parse_trailing_number(name_str: str) -> int:
    """Extract trailing number from something like 'N56E29-001'."""
    if not name_str:
        return 999999999
    parts = name_str.rsplit('-', 1)
    if len(parts) < 2:
        return 999999999
    suffix = parts[1]
    match = re.match(r"^(\d+)", suffix)
    if match:
        return int(match.group(1))
    return 999999999

def reorder_overlays(document_node, namespace):
    """
    After we've created all overlays, reorder them by
    the trailing digits in <name>.
    """
    overlays = document_node.findall("kml:GroundOverlay", namespace)

    def get_overlay_key(overlay):
        name_el = overlay.find(".//kml:name", namespace)
        if name_el is not None and name_el.text:
            return parse_trailing_number(name_el.text)
        return 999999999

    sorted_overlays = sorted(overlays, key=get_overlay_key)

    # Remove them all first
    for ov in overlays:
        document_node.remove(ov)

    # Re-append in sorted order
    for ov in sorted_overlays:
        document_node.append(ov)

def modify_kml(kml_path, png_filepaths, subfolder_name):
    """
    - 'png_filepaths' is a list of absolute file paths to .png images.
    - We copy them into a subfolder within the extracted .kmz named `subfolder_name`.
    - The Icon href for each <GroundOverlay> is set to 'subfolder_name/<filename>.png'.
    - Then reorder them by trailing numeric portion in <name>.
    - Also rename the <Document> or <Folder> node to match 'subfolder_name'.
    - If there's no <Document> or <Folder>, we create a <Document>.
    """
    namespace = {'kml': 'http://www.opengis.net/kml/2.2'}
    tree = ET.parse(kml_path)
    root = tree.getroot()

    # Find all existing ground overlays
    overlays = root.findall(".//kml:GroundOverlay", namespace)
    if not overlays:
        print("No GroundOverlay found in the KML file.")
        return False
    first_overlay = overlays[0]

    # Try to find <Document> or <Folder>
    document_node = root.find(".//kml:Document", namespace)
    if document_node is None:
        document_node = root.find(".//kml:Folder", namespace)

    # If still none, create a <Document> ourselves and move the overlays under it.
    if document_node is None:
        print("No <Document> or <Folder> found in KML; creating one ourselves.")
        document_node = ET.SubElement(root, "{http://www.opengis.net/kml/2.2}Document")
        # Move existing GroundOverlay(s) under the new <Document>
        for ov in overlays:
            root.remove(ov)
            document_node.append(ov)

    # Rename <Document> or <Folder> to match subfolder name
    doc_name_el = document_node.find("kml:name", namespace)
    if doc_name_el is not None:
        doc_name_el.text = subfolder_name
    else:
        doc_name_el = ET.SubElement(document_node, "{http://www.opengis.net/kml/2.2}name")
        doc_name_el.text = subfolder_name

    # Identify the PNG used by the original overlay
    original_png = None
    icon_href_el = first_overlay.find(".//kml:Icon/kml:href", namespace)
    if icon_href_el is not None and icon_href_el.text:
        original_png = os.path.basename(icon_href_el.text.strip())

    # Build (absolute_path, filename)
    all_pngs = []
    for abs_path in png_filepaths:
        fname = os.path.basename(abs_path)
        all_pngs.append((abs_path, fname))

    # If the original PNG is in the list, skip duplicating it
    duplicates_skipped = False
    create_list = []
    for (abs_path, fname) in all_pngs:
        if original_png and fname == original_png:
            duplicates_skipped = True
        else:
            create_list.append((abs_path, fname))

    print(f" Template overlay uses: {original_png if original_png else '(No PNG reference)'}")
    if duplicates_skipped:
        print(f" Skipped duplicating original PNG: {original_png}")

    # Create overlays for each new PNG
    for (abs_path, fname) in create_list:
        # Clone the first overlay as a template
        new_overlay = ET.fromstring(ET.tostring(first_overlay))

        # <Icon><href> => subfolder_name/filename
        icon_href = new_overlay.find(".//kml:Icon/kml:href", namespace)
        if icon_href is not None:
            icon_href.text = f"{subfolder_name}/{fname}"

        # <name> => filename without extension
        overlay_name = new_overlay.find(".//kml:name", namespace)
        if overlay_name is not None:
            overlay_name.text = os.path.splitext(fname)[0]

        # Append the new overlay
        document_node.append(new_overlay)

    # Create a subfolder for images inside the extracted KMZ
    kml_dir = os.path.dirname(kml_path)
    subfolder_fullpath = os.path.join(kml_dir, subfolder_name)
    if not os.path.exists(subfolder_fullpath):
        os.makedirs(subfolder_fullpath)

    # Copy all PNGs (including original) into that subfolder
    for (abs_path, fname) in all_pngs:
        dst_path = os.path.join(subfolder_fullpath, fname)
        if not os.path.isfile(abs_path):
            print(f"WARNING: PNG file does not exist: {abs_path}")
            continue
        shutil.copy2(abs_path, dst_path)

    # Reorder overlays by trailing numeric portion
    reorder_overlays(document_node, namespace)

    # Save changes
    tree.write(kml_path, xml_declaration=True, encoding='utf-8')
    return True

def process_subfolder(subfolder_path, output_folder):
    """
    For a given subfolder (e.g. 'N56E29'), find the single .kmz, gather .png files,
    create a new .kmz named 'N56E29 - transparent.kmz' in 'output_folder'.
    """
    subfolder_name = os.path.basename(subfolder_path)
    kmz_files = [f for f in os.listdir(subfolder_path) if f.lower().endswith('.kmz')]
    if not kmz_files:
        print(f"No .kmz found in {subfolder_path}. Skipping.")
        return False
    if len(kmz_files) > 1:
        print(f"Multiple .kmz files found in {subfolder_path}, using the first: {kmz_files[0]}")
    original_kmz = kmz_files[0]
    original_kmz_path = os.path.join(subfolder_path, original_kmz)

    # Build a list of absolute PNG paths
    all_files = os.listdir(subfolder_path)
    png_files_abs = [
        os.path.join(subfolder_path, f)
        for f in all_files
        if f.lower().endswith('.png')
    ]
    if not png_files_abs:
        print(f"No PNG files in {subfolder_path}. Skipping.")
        return False

    new_kmz_name = f"{subfolder_name} - transparent.kmz"
    new_kmz_path = os.path.join(output_folder, new_kmz_name)

    # Create a temp folder for extraction
    temp_dir = os.path.join(output_folder, f"temp_extract_{subfolder_name}")
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    # Extract original .kmz
    extract_kmz(original_kmz_path, temp_dir)

    # Find the .kml
    kml_file = None
    for root_dir, _, files in os.walk(temp_dir):
        for file in files:
            if file.lower().endswith(".kml"):
                kml_file = os.path.join(root_dir, file)
                break
        if kml_file:
            break

    if not kml_file:
        print(f"No .kml found in {original_kmz_path}. Skipping.")
        shutil.rmtree(temp_dir)
        return False

    print(f"Modifying KML in subfolder: {subfolder_name}")
    success = modify_kml(kml_file, png_files_abs, subfolder_name)
    if not success:
        print("modify_kml() failed.")
        shutil.rmtree(temp_dir)
        return False

    # Repackage as new .kmz (alphabetical file ordering)
    create_kmz(temp_dir, new_kmz_path)

    # Cleanup
    shutil.rmtree(temp_dir)
    print(f"Created: {new_kmz_path}")
    return True

def main():
    root_tk = tk.Tk()
    root_tk.withdraw()
    try:
        main_folder = filedialog.askdirectory(
            title="Select the MAIN folder with subfolders"
        )
        if not main_folder:
            messagebox.showwarning("Warning", "No main folder selected.")
            return

        output_folder = filedialog.askdirectory(
            title="Select the OUTPUT folder for new .kmz files"
        )
        if not output_folder:
            messagebox.showwarning("Warning", "No output folder selected.")
            return

        # Each subdir is something like "N56E29"
        all_subfolders = [
            os.path.join(main_folder, d)
            for d in os.listdir(main_folder)
            if os.path.isdir(os.path.join(main_folder, d))
        ]
        if not all_subfolders:
            messagebox.showinfo("Info", "No subfolders found.")
            return

        success_count = 0
        for subf in all_subfolders:
            print("=" * 60)
            print(f"Processing subfolder: {subf}")
            ok = process_subfolder(subf, output_folder)
            if ok:
                success_count += 1
            print()

        messagebox.showinfo(
            "Done",
            f"Processed {success_count} subfolders.\n"
            f"KMZ files are in: {output_folder}"
        )

    except Exception as e:
        messagebox.showerror("Unexpected Error", str(e))

if __name__ == "__main__":
    main()
