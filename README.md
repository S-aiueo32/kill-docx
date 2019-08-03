# :fu: kill-docx
> SAY NO TO MS WORD

**[INFO]**
[This article](https://qiita.com/S-aiueo32/items/02dc7a6ea481d0f234f5) based on [this commit](https://github.com/S-aiueo32/kill-docx/tree/c20baaa8064e3633d9884c40a4b26ed015442f3e), not `master`.

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

**[IMPORTANT]** `pdf2image` package requires to install `poppler` which is out of `pip`. So check [here](https://github.com/Belval/pdf2image) out!

## Usage
Let's kill them! You will see a `docx` file like `./src/chicken.pdf`.
```
python main.py
```
You can set arguments:
```
--input : PDF file to input.
--output: Filepath to output docx file.
--size: Paper size to output. choose from A0-A5, B0-B5, US Letter and Legal.
--direction: Portrait mode or Landscape mode.
--save_pages: Whether each page is saved in the directory.
```

## :chicken: Reference
1. Zongker, Doug. "Chicken chicken chicken: Chicken chicken." Annals of Improbable Research 12 , no. 5 (2006): 16--21.
