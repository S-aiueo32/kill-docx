from io import BytesIO
from pathlib import Path
import argparse

from docx import Document
from docx.shared import Mm
from pdf2image import convert_from_path


PAPER_SIZE = {
    # A papers
    'a0': {'short': 841, 'long': 1189},
    'a1': {'short': 594, 'long': 841},
    'a2': {'short': 420, 'long': 594},
    'a3': {'short': 297, 'long': 420},
    'a4': {'short': 210, 'long': 297},
    'a5': {'short': 148, 'long': 210},
    # B papers
    'b0': {'short': 1000, 'long': 1414},
    'b1': {'short': 707, 'long': 1000},
    'b2': {'short': 500, 'long': 707},
    'b3': {'short': 353, 'long': 500},
    'b4': {'short': 250, 'long': 353},
    'b5': {'short': 176, 'long': 250},
    # US paper size
    'letter': {'short': 216, 'long': 279},
    'legal': {'short': 216, 'long': 356}
}


def add_pages(doc, path, save_pages=False):
    pages = convert_from_path(str(path))
    for i, page in enumerate(pages):
        if args.save_pages:
            name_ = path.name.replace('.pdf', f'_{i:0>4}.png')
            output = str(path.with_name(name_))
        else:
            output = BytesIO()
        page.save(output, 'PNG')
        doc.add_picture(output, width, height)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', type=str, default='./src/chicken.pdf')
    parser.add_argument('--output', type=str, default='./output.docx')
    parser.add_argument('--size', default='a4', choices=PAPER_SIZE.keys())
    parser.add_argument('--direction', default='portrait', choices=['portrait', 'landscape'])
    parser.add_argument('--save_pages', action='store_true', default=False)
    args = parser.parse_args()

    # store args
    input_path = Path(args.input)
    output_path = Path(args.output)
    size = PAPER_SIZE[args.size]
    if args.direction == 'portrait':
        width, height = Mm(size['short']), Mm(size['long'])
    else:
        width, height = Mm(size['long']), Mm(size['short'])

        # instantiation docx
    doc = Document()

    # set parameters
    section, *_ = doc.sections  # store 1st section
    section.bottom_margin = 0
    section.left_margin = 0
    section.right_margin = 0
    section.top_margin = 0
    section.page_width = width
    section.page_height = height

    # convert and add pages to docx
    add_pages(doc, input_path, args.save_pages)

    # save
    doc.save(output_path)


if __name__ == '__main__':
    main()
