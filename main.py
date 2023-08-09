import os
import sys
from typing import List

import magic

from PIL import Image
from docx2pdf import convert
from pillow_heif import register_heif_opener


def get_file_mime_type(file_path):
    try:
        mime = magic.Magic()
        return magic.from_file(file_path, mime=True)
    except Exception as e:
        print(f"Error while detecting mime type: {e}")
        return None


def is_heic_file(file_path):
    mime_type = get_file_mime_type(file_path)
    return mime_type and "application/octet-stream" in mime_type


def is_png_file(file_path):
    mime_type = get_file_mime_type(file_path)
    return mime_type and "image/png" in mime_type


def is_docx_file(file_path):
    mime_type = get_file_mime_type(file_path)
    return (
        mime_type
        and "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        in mime_type
    )


if __name__ == "__main__":
    files: List[str] = sys.argv

    register_heif_opener()

    for f_path in files[1:]:  # the executable is always the first file
        try:
            with open(f_path, "r") as file:
                output = os.path.dirname(f_path) + "\\results"
                try:
                    os.makedirs(output)
                    print(f"New folder created: {output}")
                except Exception as e:
                    print(f"Error creating folder: {e}")

                output += "\\" + os.path.basename(f_path)

                if is_docx_file(f_path):
                    out_pdf_path = output + ".pdf"
                    convert(f_path, out_pdf_path)
                    print("Successful conversion from Word to PDF.")

                if is_png_file(f_path) or is_heic_file(f_path):
                    try:
                        img = Image.open(f_path)
                        out_jpg_path = output + ".jpg"
                        if img.mode != "RGB":
                            img = img.convert("RGB")

                        img.save(out_jpg_path, format="JPEG")
                        print(f"Conversion successful: {out_jpg_path}")
                    except Exception as e:
                        print(f"Error converting: {e}")

        except FileNotFoundError:
            print(f"File not found at path: {f_path}")
        except Exception as e:
            print(f"An error occurred: {e}")
    input("Press Enter to exit...")
