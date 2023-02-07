# What is office_pdf.py.
The scripts that generates PDF files from MS Office files and manipulate them.

# Installation
This document describes how to set up the environment for the application.

## Install Python 3.
Install Python 3.10 from https://www.python.org/downloads/.<br>
Add python3 to the PATH environment.

## Install packages using requirement.txt
All the below packages can be install by following commands.
Or each package can be installed one by one.
```
pip install -r requirement.txt
```

# Python + Windows COM
https://pbpython.com/windows-com.html

# PDF + Python
https://johannesfilter.com/python-and-pdf-a-review-of-existing-tools/
https://nanonets.com/blog/pypdf2-library-working-with-pdf-files-in-python/
https://realpython.com/pdf-python/

# Directory structure
```
```

# The structure of the application.
```
```

# Application design
There are several command patterns and two of them produce PDFs from Access and Word files.
Others like `gen_pdf_cmd`, `combine_pdf_cmd` and `impose_pdf_cmd` manipulate PDF and produce another PDF.
`CommandExecutor` manages all these and produces the final output the developer wants to produce.

## base_pdf_cmd
The `base_pdf_cmd` is the base class for all the `xxx_pdf_cmd` that has input parameters and output PDF files.

## access_pdf_cmd
This command produces PDF file from a Access report.

## word_pdf_cmd
This command produces PDF files from Word document files
by creating a docx file first and exporting the range as a PDF.

## gen_pdf_cmd
The gen_pdf_cmd command generates PDF from scratch using reportlab.

## combine_pdf_cmd
The `combine_pdf_cmd` combines multiple pdfs into one by choosing specified ranges of PDF pages.
It uses `PyPDF2` to do the job.

## impose_pdf_cmd
The `impose_pdf_cmd` command imposes source PDF and produced imposed PDF.
i.e. Perfect or Saddle Stitch.

## CommandExecutor
The `CommandExecutor` manages list of `xxx_pdf_cmd`s and run them based on their dependency
and produce the final output PDF.
