# %%
import os

from pptx_parser import parse_pptx
from utilities import generate_frontmatter, sanitize_filename


def write_slide_file(slide_data, output_folder="."):
    front_matter_items = {}
    front_matter_items["title"] = slide_data["title"]

    has_update_dates = "update_dates" in slide_data and slide_data['update_dates'] != None and len(
        slide_data['update_dates']) > 0
    if "creation_date" in slide_data:
        front_matter_items["created"] = slide_data["creation_date"]

    if has_update_dates:
        front_matter_items["update_dates"] = ", ".join(
            slide_data["update_dates"])

    reference = None

    if "reference" in slide_data:
        reference = ".".join(slide_data["reference"])
        front_matter_items["reference"] = reference

    yaml_front_matter = generate_frontmatter(front_matter_items)

    # Prepare the content
    content = f"# {slide_data['title']}\n\n"

    if has_update_dates:
        content += f"Update Dates: {', '.join(slide_data['update_dates'])}\n\n"

    filename = f"{slide_data['filename']}.md"
    file_path = os.path.join(output_folder, filename)

    # Add the reference if it exists
    if reference is not None:
        link = os.path.join("..", reference, filename)
        content += f"Reference: [{reference}]({link})\n\n"

    content += slide_data["content"]

    with open(file_path, "w") as f:
        f.write(yaml_front_matter)
        f.write(content)

    return file_path


def write_markdown(slides_data, meta={}, output_folder="."):
    for slide_data in slides_data:
        # if the title is a duplicate (check case insesitive), add the number
        # of the slide to the title
        if any(slide_data['title'].lower() == s['title'].lower() for s in slides_data):
            slide_data['filename'] = sanitize_filename(
                f"{slide_data['title']} ({slides_data.index(slide_data) + 1})")
        else:
            slide_data['filename'] = sanitize_filename(slide_data['title'])

        write_slide_file(slide_data, output_folder)

    # Write the index file
    index_file_path = os.path.join(output_folder, "_index.md")

    with open(index_file_path, "w") as f:
        fm = {
            "title": "Index",
            "author": meta.get("author", None),
            "modified": meta.get("modified", None)
        }

        f.write("# Index\n\n")

        idx = 0
        for slide_data in slides_data:
            link = f"./{slide_data['filename']}.md"
            idx += 1
            f.write(f"{idx}. [{slide_data['title']}]({link})\n")

    # Write the index file
    index_file_path = os.path.join(output_folder, "_slides.md")
    with open(index_file_path, "w") as f:
        fm = {
            "title": "Slides",
            "author": meta.get("author", None),
            "modified": meta.get("modified", None)
        }

        f.write(generate_frontmatter(fm))

        f.write("# Full Index\n\n")

        idx = 0
        for slide_data in slides_data:
            link = f"./{slide_data['filename']}.md"
            idx += 1
            f.write(f"---\n![{slide_data['title']}]({link})\n")


# %%
