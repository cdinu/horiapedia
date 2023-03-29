import os
import shutil

from pptx import Presentation


def create_directory_structure(output_folder, slide_data):
    """
    Create the directory structure based on slide_data and return the output path.
    """
    category_hierarchy = slide_data.get("reference", "").split(".")
    output_path = os.path.join(output_folder, *category_hierarchy)

    # Create the directories if they don't exist
    os.makedirs(output_path, exist_ok=True)

    return output_path
