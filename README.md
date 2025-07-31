# PDF TO EPUB Converter GUI software in Python
#### Forked from TufayelLUS

This is a software that helps you generate calibre style epub file from a given PDF file. <br>

## Requirements to use
Please be sure to have poetry package manager installed

# How to run ?
```sh
git clone REPO
cd REPO
poetry install

# Working for MacOS, not tested on Windows or Linux, should be same
poetry run pyinstaller epub_maker.py --name "Epub Maker" --windowed --icon icon.png
```

# MacOS
You can drop dist/Epub Maker.app in application folder