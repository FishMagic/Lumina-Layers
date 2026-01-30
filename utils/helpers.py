"""
Lumina Studio - Helper Functions
Helper functions module
"""

import shutil
import zipfile
import re

from pathlib import Path
from typing import List, Optional

from utils.add_3mf_colors import register_namespaces, detect_color_mode, add_colors_to_xml_string


def safe_fix_3mf_names(filepath: str, slot_names: List[str], create_assembly: bool = True,
                       enable_colors: bool = True, color_mode: Optional[str] = None):
    """
    Fix object names in 3MF file and optionally create an assembly.
    Maps objects to slot_names in order they appear in file.

    Args:
        filepath: 3MF file path
        slot_names: Object name list
        create_assembly: Whether to create assembly
        enable_colors: Whether to enable 3MF colors (default: True)
        color_mode: Color mode ("rybw" or "cmyw", auto-detect if None)
    """
    try:


        # Read original 3MF
        with zipfile.ZipFile(filepath, 'r') as zf_in:
            files_data = {}
            for name in zf_in.namelist():
                files_data[name] = zf_in.read(name)

        # Find the 3D model file
        model_file = None
        for name in files_data:
            if name.endswith('.model') and '3D/' in name:
                model_file = name
                break

        if model_file and model_file in files_data:
            content = files_data[model_file].decode('utf-8')

            # Find all <object> tags with their IDs (in order of appearance)
            object_pattern = re.compile(r'<object\s+([^>]*)>', re.IGNORECASE)

            # Track which objects we've seen
            obj_info = []  # List of (start_pos, end_pos, full_tag, id)

            for match in object_pattern.finditer(content):
                attrs = match.group(1)
                id_match = re.search(r'\bid="(\d+)"', attrs)
                if id_match:
                    obj_id = id_match.group(1)
                    obj_info.append((match.start(), match.end(), match.group(0), obj_id))

            # Collect object IDs for assembly
            object_ids = [info[3] for info in obj_info]
            print(f"[DEBUG] Found {len(object_ids)} objects in 3MF: {object_ids}")

            # Process in reverse order to preserve positions (for name fixing)
            for idx, (start, end, old_tag, obj_id) in enumerate(reversed(obj_info)):
                real_idx = len(obj_info) - 1 - idx
                if real_idx >= len(slot_names):
                    continue

                color_name = slot_names[real_idx]

                # Remove existing name attribute and add new one
                new_tag = re.sub(r'\s+name="[^"]*"', '', old_tag)
                new_tag = new_tag[:-1] + f' name="{color_name}">'

                content = content[:start] + new_tag + content[end:]

            # Create assembly if requested
            if create_assembly and len(object_ids) > 1:
                # Find the maximum object ID
                max_id = max(int(oid) for oid in object_ids)
                assembly_id = max_id + 1

                # Create assembly object XML
                components_xml = '\n'.join([f'      <component objectid="{oid}" />' for oid in object_ids])
                assembly_xml = f'''
  <object id="{assembly_id}" type="model" name="Lumina_Model">
    <components>
{components_xml}
    </components>
  </object>
'''

                # Insert assembly before </resources>
                resources_end = content.find('</resources>')
                if resources_end != -1:
                    content = content[:resources_end] + assembly_xml + content[resources_end:]
                    print(f"[DEBUG] Created assembly with id={assembly_id}, containing {len(object_ids)} components")

                # Modify <build> section to only reference the assembly
                # Find and replace the build section
                build_pattern = re.compile(r'<build>.*?</build>', re.DOTALL)
                build_match = build_pattern.search(content)
                if build_match:
                    new_build = f'<build>\n    <item objectid="{assembly_id}" />\n  </build>'
                    content = content[:build_match.start()] + new_build + content[build_match.end():]
                    print(f"[DEBUG] Updated build section to reference assembly")

            # [NEW] If colors are enabled, add color information directly to XML string
            if enable_colors:
                try:
                    print(f"[COLORS] Registered 3MF namespaces")

                    # Determine color mode if not provided
                    if color_mode is None:
                        # Auto-detect from slot_names
                        detected_mode = detect_color_mode(slot_names)
                        actual_color_mode = detected_mode
                        print(f"[COLORS] Auto-detected color mode: {actual_color_mode}")
                    else:
                        actual_color_mode = color_mode
                        print(f"[COLORS] Using specified color mode: {actual_color_mode}")

                    # Add colors directly to XML string in memory (no file I/O)
                    # This is much faster than re-reading the entire 3MF file
                    modified_content = add_colors_to_xml_string(content, actual_color_mode)

                    # Update files_data with modified XML (this will be written to file next)
                    files_data[model_file] = modified_content.encode('utf-8')
                    print(f"[COLORS] Successfully added colors to XML in memory")
                except Exception as e:
                    print(f"[COLORS] Warning: Failed to add colors: {e}")
            else:
                # If no colors, encode original content
                files_data[model_file] = content.encode('utf-8')

        # Write back to file
        with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zf_out:
            for name, data in files_data.items():
                zf_out.writestr(name, data)

        print(f"[DEBUG] 3MF file updated successfully: {filepath}")

    except Exception as e:
        print(f"Warning: Could not fix 3MF names: {e}")
