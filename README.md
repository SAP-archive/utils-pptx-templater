# pptxTemplater

This module can be used only in combination with docxtemplater.

The pptxtemplater modifies powerpoint template and provided data before docxtemplater generates the final version of the document.

## Installation
Use npm install to get the module
```
npm install docxtemplater --save
npm install pptxtemplater --save
```

In your code use the pptxtemplater as a module:
```
let pptxTemplaterModule = require('pptxtemplater');
let doc = new this.Docxtemplater(content);
doc.attachModule(pptxTemplaterModule);
```


## Usage

There are two new features added to the standard functionality of docxtemplater:

### Splitt PowerPoint table slides
You can now split one long table into multiple slides with the smaller number of table rows. To do so, you have to provide maximal number of rows in your template that should be shown on one slide.

    {#array;max_rows:10}{your_data}{/array}


### Duplicate PowerPoint slides
If you have a slide which should be duplicated dynamicaly depending on provideded data you can use "multiplier" tag: `$`. Each slide, containing this tag will be as many times duplicated as many entries are stored in the array with the same name.

If there are no data for this tag, this slide will be deleted.

    {$data_to_bring_on_multiple_slides}

Here is an example:

Template
![Template sample](https://github.wdf.sap.corp/raw/rapid-release/pptxtemplater/master/img/sample_template.png)

Otput
![Output sample](https://github.wdf.sap.corp/raw/rapid-release/pptxtemplater/master/img/sample_output.png)

For data
```
{
    "DATA_TO_BRING_ON_MULTIPLE_SLIDES": [{
        "COUNTRY": "Germany",
        "REVENUE_PER_REGION": [{
                "Region": "Berlin",
                "ZIP": "55014",
                "Revenue": "€93896.94"
            }, {
                "Region": "Saxony-Anhalt",
                "ZIP": "30652",
                "Revenue": "€51759.11"
            }, {
                "Region": "HH",
                "ZIP": "05652",
                "Revenue": "€51355.41"
            }]
        },{
        "COUNTRY": "Austria",
        "REVENUE_PER_REGION": [{
                "Region": "Bgl",
                "ZIP": "8571",
                "Revenue": "€77899.04"
            },
            {
                "Region": "Vienna",
                "ZIP": "4403",
                "Revenue": "€77276.25"
            },
            {
                "Region": "Vienna",
                "ZIP": "9122",
                "Revenue": "€48795.89"
            }]
        }
    ]
}
```
