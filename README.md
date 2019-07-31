# :fu: kill-docx
> Say NO to MS Word, Say YES to LaTeX.

## Introduction
These codes convert PDF to PNG and embed them to Word documents, in the other words, kill Microsoft Word.
You can do any assignments with LaTeX even if `.docx` format is required.

## Requirements
Here works with `Python` with a few package below. I recommed to create virtual environment with `pipenv`.
```
Python 3.6
 - python-docx
 - pdf2image
 - pillow
```

**[IMPORTANT]** `pdf2image` package requires to install `poppler` which is out of `pip` package managements. So check [here](https://github.com/Belval/pdf2image) out!

## Usage
Let's kill them! You will see a `docx` file like `./src/chicken.pdf`.
```
python main.py
```
You can set arguments below:
```
--input : PDF file to input.
--output: Filepath to output docx file.
--size: Paper size to output. choose from A0-A5, B0-B5, US Letter and Legal.
--direction: Portrait mode or Landscape mode.
--save_pages: Whether each page is saved in the directory.
```

## :chichen: Reference
1. Zongker, Doug. "Chicken chicken chicken: Chicken chicken." Annals of Improbable Research 12 , no. 5 (2006): 16--21.
