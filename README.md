# pptxTemplater

This module can be used only in combination with docxtemplater.

The pptxTemplater modifies pptx template and provided data before docxtemplate generates the final version of the document.

## Installation

## Usage

There are two new features added to the standard functionality of docxtemplater:

### Splitt PowerPoint table slides
You can now split one long table into multiple slides with the smaller tables. To do so, you have to provide maximal number of rows in your template that should be shown on one slide.

    {#array;max_rows:10}{your_data}{/array}



### Duplicate PowerPoint slides