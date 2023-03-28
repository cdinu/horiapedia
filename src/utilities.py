import os
import string

import yaml


def format_date(date_str):
    """
    Format the date string (e.g., 2011_08_18) as required.
    """
    date_parts = date_str.split('_')

    # Check if there are exactly three elements after splitting
    if len(date_parts) != 3:
        print(f"Invalid date string: {date_str}")
        return None
    year, month, day = date_parts
    formatted_date = f"{year}-{month}-{day}"
    return formatted_date


def get_reference(reference_string):
    """
    Extract the reference components from a given reference string.

    Args:
        reference_string (str): A string containing the reference, possibly with a .ppt or .pptx extension.

    Returns:
        list: A list of strings representing the reference components obtained by splitting the input string by '.'.
    """
    if reference_string.endswith(".ppt"):
        reference_string = reference_string[:-4]
    elif reference_string.endswith(".pptx"):
        reference_string = reference_string[:-5]

    # Split the reference string using '.' as the separator
    reference_components = reference_string.split('.')

    return reference_components


def get_dir_for_file_path(file_path, base_dir="."):
    """
    Get the directory path for a given file path based on the file's reference.

    Args:
        file_path (str): The file path of the PowerPoint file.
        base_dir (str, optional): The base directory where the file's directory should be located. Defaults to ".".

    Returns:
        str: The directory path for the given file based on its reference.
    """
    file_name = os.path.basename(file_path)
    return os.path.join(base_dir, ".".join(get_reference(file_name)))


def sanitize_filename(text, replacement="-"):
    """
    Sanitize a text variable so it can be safely used as a filename.

    Args:
        text (str): The text to sanitize.
        replacement (str, optional): The character to replace invalid characters with. Defaults to "_".

    Returns:
        str: The sanitized filename.
    """
    valid_chars = ".-_%s%s" % (string.ascii_letters, string.digits)
    sanitized_filename = "".join(
        c if c in valid_chars else replacement for c in text)
    sanitized_filename = sanitized_filename.strip(replacement)

    # remove double
    sanitized_filename = sanitized_filename.replace(
        replacement + replacement, replacement)
    sanitized_filename = sanitized_filename.replace(
        replacement + replacement, replacement)

    sanitized_filename = sanitized_filename[:128]
    return sanitized_filename


def generate_frontmatter(dict):
    """
    Generate the frontmatter for a markdown file based on the given dictionary.

    Args:
        dict (dict): A dictionary containing the frontmatter data.

    Returns:
        str: The frontmatter for the markdown file.
    """
    yaml_string = yaml.dump(dict, default_flow_style=False)
    return f"---\n{yaml_string}---\n"
