# extract-xml-from-docx

Extract document.xml (file containing main text elements in .docx files) from .docx

document.xml file is placed in working directory and has file name prepended with the name of the original .docx, e.g.: myfile.docx -> myfile.docx.document.xml

Accepts two options, **-r** and **-p**

**-r**: Walk directory tree, extracting document.xml from any .docx found (only searches in working directory by default)

**-p**: Pretty print document.xml in output file
