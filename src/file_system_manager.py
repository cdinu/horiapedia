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


def extract_images(slide_data, output_folder):
    """
    Extract images from the slide and save them in the output folder.
    Return a list of image paths.
    """
    image_paths = []
    pptx = Presentation(slide_data["pptx_path"])
    slide = pptx.slides[slide_data["slide_index"]]

    for idx, shape in enumerate(slide.shapes):
        if shape.shape_type == 13:  # Picture shape type
            image = shape.image
            img_path = os.path.join(
                output_folder, f"{slide_data['title']}-{idx + 1}.png")

            with open(img_path, "wb") as img_file:
                img_file.write(image.blob)

            # Store the relative path of the image
            relative_img_path = f"./{slide_data['title']}-{idx + 1}.png"
            image_paths.append(relative_img_path)

    return image_paths
