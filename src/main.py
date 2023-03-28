import logging
import os

from file_system_manager import create_directory_structure
from markdown_writer import write_markdown
from pptx_parser import parse_pptx

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')


def process_pptx_files(input_folder, output_folder):
    # Process each PPTX file in the input folder
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.pptx'):
            file_path = os.path.join(input_folder, file_name)

            logging.debug(f'Processing file: {file_path}')

            slides_data = parse_pptx(file_path)

            [slides_data, meta, output_dir] = parse_pptx(
                file_path, base_dir=output_folder)

            write_markdown(slides_data=slides_data, meta=meta,
                           output_folder=output_dir)


if __name__ == "__main__":
    input_folder = "../input_pptx_files"
    output_folder = "../output_markdown_files"

    process_pptx_files(input_folder, output_folder)
