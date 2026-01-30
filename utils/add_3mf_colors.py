"""
Lumina Studio - 3MF Color Module
Color addition functions for 3MF files using colorgroup + triangle.pid/p1 method

Author: Lumina Studio
Version: 1.0.0
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any


# ============ Color Mode Definitions ============

# RYBW Color Mode (Red/Yellow/Blue/White)
RYBW_COLORS: Dict[str, str] = {
    "Red": "#FF0000FF",
    "Yellow": "#FFFF00FF",
    "Blue": "#0000FFFF",
    "White": "#FFFFFFFF"
}

# CMYW Color Mode (Cyan/Magenta/Yellow/White)
CMYW_COLORS: Dict[str, str] = {
    "Cyan": "#00FFFFFF",
    "Magenta": "#FF00FFFF",
    "Yellow": "#FFFF00FF",
    "White": "#FFFFFFFF"
}

# 3MF Material Namespaces
MATERIAL_NAMESPACE = "http://schemas.microsoft.com/3dmanufacturing/material/2015/02"
MODEL_NAMESPACE = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02"

# XML Namespace Mapping
NAMESPACES = {
    'm': MATERIAL_NAMESPACE,
    'model': MODEL_NAMESPACE
}


# ============ Helper Functions ============

def register_namespaces():
    """
    Register XML namespaces

    Registers all 3MF standard namespaces to ensure generated XML is compatible with BambuStudio
    """
    # Register all 3MF standard namespaces
    ET.register_namespace('', MODEL_NAMESPACE)  # Default namespace
    ET.register_namespace('b', 'http://schemas.microsoft.com/3dmanufacturing/beamlattice/2017/02')
    ET.register_namespace('m', MATERIAL_NAMESPACE)
    ET.register_namespace('p', 'http://schemas.microsoft.com/3dmanufacturing/production/2015/06')
    ET.register_namespace('s', 'http://schemas.microsoft.com/3dmanufacturing/slice/2015/07')
    ET.register_namespace('sc', 'http://schemas.microsoft.com/3dmanufacturing/securecontent/2019/04')


def get_colorgroup_id(color_name: str, color_mode: str = 'rybw') -> str:
    """
    Generate colorgroup ID based on color name and color mode

    Args:
        color_name: Color name (e.g., "Red", "Cyan")
        color_mode: Color mode ("rybw" or "cmyw")

    Returns:
        colorgroup ID
    """
    # RYBW color mode mapping
    rybw_map = {
        'Red': '1',
        'Yellow': '2',
        'Blue': '3',
        'White': '4'
    }

    # CMYW color mode mapping
    cmyw_map = {
        'Cyan': '1',
        'Magenta': '2',
        'Yellow': '3',
        'White': '4'
    }

    # Select appropriate mapping based on color mode
    if color_mode == 'cmyw':
        color_to_id_map = cmyw_map
    else:  # rybw
        color_to_id_map = rybw_map

    # Return '1' (first color) if color name is not in mapping
    return color_to_id_map.get(color_name, '1')


def detect_color_mode(object_names: List[str]) -> str:
    """
    Automatically detect color mode based on object name list

    Args:
        object_names: List of object names

    Returns:
        "rybw" or "cmyw"
    """
    # Count matching names
    rybw_count = sum(1 for name in object_names if name in RYBW_COLORS)
    cmyw_count = sum(1 for name in object_names if name in CMYW_COLORS)

    # Prefer mode with more matches
    if cmyw_count > rybw_count:
        return "cmyw"
    elif rybw_count > cmyw_count:
        return "rybw"
    else:
        # Default to RYBW if unable to determine
        return "rybw"


def is_assembly(object_element: ET.Element) -> bool:
    """
    Check if object is an assembly object

    Args:
        object_element: object element

    Returns:
        True if assembly, False otherwise
    """
    # Check for <components> child element (handle namespaces)
    # First try without namespace
    components = object_element.find('components')
    if components is not None:
        return True

    # If not found, try with namespace
    for child in object_element:
        if 'components' in child.tag.split('}'):
            return True

    return False


# ============ 3MF File Reading Functions ============

def read_3mf(input_path: Path) -> Tuple[ET.Element, Dict[str, bytes]]:
    """
    Read 3MF file and parse XML

    Args:
        input_path: 3MF file path

    Returns:
        (root_element, other_files) tuple
        - root_element: XML root element of 3D/3dmodel.model
        - other_files: Dictionary of other files {file_path: file_content}

    Raises:
        ValueError: If file is not a valid 3MF file
    """
    # Check if it's a ZIP file
    if not zipfile.is_zipfile(input_path):
        raise ValueError(f"File is not a valid ZIP file: {input_path}")

    other_files = {}

    with zipfile.ZipFile(input_path, 'r') as zip_ref:
        # List all files
        file_list = zip_ref.namelist()
        print(f"3MF file contains {len(file_list)} files")

        # Read 3D/3dmodel.model
        model_path = '3D/3dmodel.model'
        if model_path not in file_list:
            raise ValueError(f"3MF file does not contain {model_path}")

        with zip_ref.open(model_path) as model_file:
            # Read XML content
            xml_content = model_file.read()
            # Parse XML
            root = ET.fromstring(xml_content)
            print(f"Successfully parsed XML, root element: {root.tag}")

        # Read other files (keep [Content_Types].xml and _rels/.rels etc.)
        for file_path in file_list:
            if file_path != model_path:
                with zip_ref.open(file_path) as f:
                    other_files[file_path] = f.read()

        print(f"Preserved {len(other_files)} auxiliary files")

    return root, other_files


def get_objects_info(root: ET.Element) -> List[Dict[str, Any]]:
    """
    Get information about all objects

    Args:
        root: XML root element

    Returns:
        List of object info, each element contains id, name, type, is_assembly, etc.
    """
    objects = []

    # Find resources element (with namespace)
    # resources can be default namespace (empty string) or with namespace
    resources = None

    # Try to find resources (without namespace)
    resources = root.find('resources')

    # If not found, try all possible namespaces
    if resources is None:
        # Check root element's namespace
        tag_namespace = root.tag.split('}')[0].strip('{') if '}' in root.tag else ''

        # Try to use the same namespace
        if tag_namespace:
            resources = root.find(f'{{{tag_namespace}}}resources')

    if resources is None:
        # Last resort: iterate all child elements directly
        for child in root:
            if 'resources' in child.tag.split('}'):
                resources = child
                break

    if resources is None:
        raise ValueError("<resources> element missing in XML")

    # Find all object elements
    for obj in resources:
        if 'object' in obj.tag.split('}'):
            obj_id = obj.get('id')
            obj_name = obj.get('name', '')
            obj_type = obj.get('type', 'model')

            objects.append({
                'element': obj,
                'id': obj_id,
                'name': obj_name,
                'type': obj_type,
                'is_assembly': is_assembly(obj)
            })

    print(f"\nFound {len(objects)} objects:")
    for obj in objects:
        assembly_str = " (assembly)" if obj['is_assembly'] else ""
        print(f"  - ID: {obj['id']}, Name: {obj['name']}, Type: {obj['type']}{assembly_str}")

    return objects


def get_triangles_info(objects: List[Dict[str, Any]]) -> Dict[str, int]:
    """
    Count triangles for each object

    Args:
        objects: Object info list

    Returns:
        Dictionary {object_name: triangle_count}
    """
    triangles_count = {}

    for obj in objects:
        if obj['is_assembly']:
            continue

        # Find mesh element (handle namespaces)
        mesh = None
        for child in obj['element']:
            if 'mesh' in child.tag.split('}'):
                mesh = child
                break

        if mesh is None:
            continue

        # Find triangles element (handle namespaces)
        triangles = None
        for child in mesh:
            if 'triangles' in child.tag.split('}'):
                triangles = child
                break

        if triangles is None:
            continue

        # Count triangle elements
        count = 0
        for child in triangles:
            if 'triangle' in child.tag.split('}'):
                count += 1

        triangles_count[obj['name']] = count

    print(f"\nTriangle count statistics:")
    for name, count in triangles_count.items():
        print(f"  - {name}: {count} triangles")

    return triangles_count


# ============ Colorgroup Insertion Functions ============

def insert_colorgroups(resources: ET.Element, color_mode: str) -> Dict[str, str]:
    """
    Insert colorgroup elements into the resources element

    Args:
        resources: resources XML element
        color_mode: Color mode ("rybw" or "cmyw")

    Returns:
        Dictionary mapping color names to colorgroup IDs
    """
    # Select color mapping
    if color_mode == 'cmyw':
        colors = CMYW_COLORS
    else:  # rybw
        colors = RYBW_COLORS

    print(f"\n" + "=" * 50)
    print(f"Step 3: Insert colorgroup elements ({color_mode.upper()} mode)")
    print("=" * 50)

    # Create colorgroup element mapping
    colorgroup_map = {}

    # Find first object element, insert colorgroups before it
    # First create all colorgroup elements
    colorgroup_elements = []
    for color_name, color_value in colors.items():
        # Create colorgroup element (with namespace)
        colorgroup_id = get_colorgroup_id(color_name, color_mode)
        colorgroup = ET.Element(f'{{{MATERIAL_NAMESPACE}}}colorgroup')
        colorgroup.set('id', colorgroup_id)

        # Create color sub-element
        color = ET.SubElement(colorgroup, f'{{{MATERIAL_NAMESPACE}}}color')
        color.set('color', color_value)

        colorgroup_elements.append(colorgroup)
        colorgroup_map[color_name] = colorgroup_id

        print(f"  Created colorgroup: {colorgroup_id} -> {color_name} ({color_value})")

    # Find index of first object element
    object_index = None
    for i, child in enumerate(resources):
        if 'object' in child.tag.split('}'):
            object_index = i
            break

    # Insert colorgroup elements
    # Insert in reverse order so first colorgroup appears at the front
    for colorgroup in reversed(colorgroup_elements):
        if object_index is not None:
            resources.insert(object_index, colorgroup)
        else:
            # If no object element found, append to resources
            resources.append(colorgroup)

    print(f"  Successfully inserted {len(colorgroup_elements)} colorgroup elements")

    return colorgroup_map


# ============ Triangle Attribute Addition Functions ============

def add_triangle_colors(objects: List[Dict[str, Any]], colorgroup_map: Dict[str, str]) -> Dict[str, int]:
    """
    Add pid and p1 attributes to triangle elements for each object

    Args:
        objects: Object info list
        colorgroup_map: Dictionary mapping color names to colorgroup IDs

    Returns:
        Modification statistics {object_name: modified_triangle_count}
    """
    print(f"\n" + "=" * 50)
    print(f"Step 4: Add pid/p1 attributes to triangles")
    print("=" * 50)

    stats = {}

    for obj in objects:
        obj_name = obj['name']
        obj_element = obj['element']

        # Skip assembly objects
        if obj['is_assembly']:
            print(f"  Skip assembly object: {obj_name}")
            continue

        # Find colorgroup ID
        if obj_name not in colorgroup_map:
            print(f"  Warning: Object '{obj_name}' has no corresponding colorgroup, skipping")
            continue

        colorgroup_id = colorgroup_map[obj_name]

        # Find mesh element
        mesh = None
        for child in obj_element:
            if 'mesh' in child.tag.split('}'):
                mesh = child
                break

        if mesh is None:
            print(f"  Warning: Object '{obj_name}' has no mesh element, skipping")
            continue

        # Find triangles element
        triangles = None
        for child in mesh:
            if 'triangles' in child.tag.split('}'):
                triangles = child
                break

        if triangles is None:
            print(f"  Warning: Object '{obj_name}' has no triangles element, skipping")
            continue

        # Add pid and p1 attributes to each triangle
        count = 0
        for triangle in triangles:
            if 'triangle' in triangle.tag.split('}'):
                triangle.set('pid', colorgroup_id)
                triangle.set('p1', '0')
                count += 1

        stats[obj_name] = count
        print(f"  {obj_name}: Added pid='{colorgroup_id}' p1='0' to {count} triangles")

    print(f"\nTotal: Added color attributes to {sum(stats.values())} triangles in {len(stats)} objects")

    return stats


# ============ Build Element Modification Functions ============

def add_build_materials(root: ET.Element, colorgroup_map: Dict[str, str]) -> Dict[str, str]:
    """
    Add materialid attribute to build items

    Args:
        root: XML root element
        colorgroup_map: Dictionary mapping color names to colorgroup IDs

    Returns:
        Modification statistics {objectid: materialid}
    """
    print(f"\n" + "=" * 50)
    print(f"Step 4.5: Add materialid attribute to build items")
    print("=" * 50)

    # Find build element
    build = None
    for child in root:
        if 'build' in child.tag.split('}'):
            build = child
            break

    if build is None:
        print("  Warning: Cannot find build element")
        return {}

    stats = {}

    # Iterate through all item elements in build
    for item in build:
        if 'item' in item.tag.split('}'):
            # Get item's objectid and partnumber
            objectid = item.get('objectid')
            partnumber = item.get('partnumber', '')

            # Find corresponding colorgroup ID based on partnumber
            if partnumber in colorgroup_map:
                materialid = colorgroup_map[partnumber]
                item.set('materialid', materialid)
                stats[objectid] = materialid
                print(f"  objectid={objectid} (partnumber={partnumber}): Added materialid='{materialid}'")
            else:
                if partnumber:
                    print(f"  Warning: objectid={objectid}'s partnumber='{partnumber}' has no corresponding colorgroup")

    print(f"\nTotal: Added materialid attribute to {len(stats)} build items")

    return stats


# ============ 3MF File Repackaging Functions ============

def write_3mf(root: ET.Element, other_files: Dict[str, bytes], output_path: Path) -> None:
    """
    Repackage modified XML and other files into a 3MF file

    Args:
        root: Modified XML root element
        other_files: Dictionary of other files {file_path: file_content}
        output_path: Output file path
    """
    print(f"\n" + "=" * 50)
    print(f"Step 5: Repackage 3MF file")
    print("=" * 50)

    # Convert XML to bytes
    # Use short_empty_elements=False to preserve original format
    xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=True, short_empty_elements=False)

    # Create new ZIP file
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
        # Write 3D/3dmodel.model
        zip_out.writestr('3D/3dmodel.model', xml_bytes)
        print(f"  Wrote: 3D/3dmodel.model")

        # Write other files
        for file_path, file_content in other_files.items():
            zip_out.writestr(file_path, file_content)
            print(f"  Wrote: {file_path}")

    print(f"\nSuccessfully generated 3MF file: {output_path}")
    print(f"File size: {output_path.stat().st_size / 1024:.2f} KB")


# ============ Main Function (New - In-Memory) ============

def add_colors_to_xml_string(xml_string: str, color_mode: str = 'auto') -> str:
    """
    Add color information to 3MF XML string using colorgroup + triangle.pid/p1 method

    This function works directly with XML strings in memory, avoiding file I/O overhead.

    Args:
        xml_string: XML content string from 3D/3dmodel.model
        color_mode: Color mode ("rybw", "cmyw", or "auto" for auto-detection)

    Returns:
        Modified XML string with color information

    Raises:
        ValueError: If XML is invalid or missing required elements
    """
    register_namespaces()

    print(f"\nColor mode: {color_mode}")

    # Parse XML from string
    print("\n" + "=" * 50)
    print("Step 1: Parse XML from string")
    print("=" * 50)
    root = ET.fromstring(xml_string.encode('utf-8'))
    print(f"Successfully parsed XML, root element: {root.tag}")

    # Get object information
    print("\n" + "=" * 50)
    print("Step 2: Parse object information")
    print("=" * 50)
    objects = get_objects_info(root)

    # Auto-detect color mode
    object_names = [obj['name'] for obj in objects if not obj['is_assembly']]
    detected_mode = detect_color_mode(object_names)

    # Determine final color mode to use
    if color_mode == 'auto':
        final_color_mode = detected_mode
        print(f"\nAuto-detected color mode: {final_color_mode}")
    else:
        final_color_mode = color_mode
        print(f"\nUsing specified color mode: {final_color_mode}")

    # Insert colorgroup elements
    # Find resources element
    resources = None
    for child in root:
        if 'resources' in child.tag.split('}'):
            resources = child
            break

    if resources is None:
        raise ValueError("Cannot find resources element")

    colorgroup_map = insert_colorgroups(resources, final_color_mode)

    # Add pid/p1 attributes to triangles
    triangle_stats = add_triangle_colors(objects, colorgroup_map)

    # Add materialid attributes to build items
    build_stats = add_build_materials(root, colorgroup_map)

    # Convert back to string
    print("\n" + "=" * 50)
    print("Step 3: Convert modified XML to string")
    print("=" * 50)
    modified_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True, short_empty_elements=False)
    modified_xml_string = modified_xml.decode('utf-8')

    print(f"\n" + "=" * 50)
    print("Done!")
    print("=" * 50)
    print(f"Successfully added color information to XML")
    print(f"Color mode: {final_color_mode}")
    print("=" * 50)

    return modified_xml_string