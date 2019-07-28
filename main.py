from pathlib import Path
import argparse

from docx import Document
from docx.shared import Mm
from pdf2image import convert_from_path

src_dir = Path('./src')

def save_pages(filename):
	pages = convert_from_path(str(src_dir / filename))
	out_files = []
	for i, page in enumerate(pages):
		out_file = src_dir / f'{filename.replace(".pdf", "")}_{i:0>4}.png'
		out_files.append(out_file)
		page.save(out_file)
	return out_files

if __name__ == '__main__':
	parser = argparse.ArgumentParser()
	parser.add_argument('--filename', type=str, default='chicken.pdf')
	args = parser.parse_args()

	doc = Document()

	sections = doc.sections
	for section in sections:
		section.top_margin = 0
		section.bottom_margin = 0
		section.left_margin = 0
		section.right_margin = 0
		section.page_height = Mm(297)
		section.page_width = Mm(210)

		
	p = doc.add_paragraph()	
	r = p.add_run()

	image_paths = save_pages(args.filename)
	
	for image_path in image_paths:
		r.add_picture(str(image_path), height=Mm(297), width=Mm(210))

	doc.save('./output.docx')