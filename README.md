# pptxTemplater

[![REUSE status](https://api.reuse.software/badge/github.com/SAP/utils-pptx-templater)](https://api.reuse.software/info/github.com/SAP/utils-pptx-templater)

The pptxtemplater modifies PowerPoint template and provided data before docxtemplater generates the final version of the document.

This module can be used only in combination with [docxtemplater][docxtemplater].


Installation
===
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

Requirements
===
 * nodejs - v6.0 and higher
 * [docxtemplater][docxtemplater] - v2.1

Usage
===
There are two new features added to the standard functionality of docxtemplater:

### Split PowerPoint table slides
You can now split one long table into multiple slides with the smaller number of table rows. To do so, you have to provide maximal number of rows in your template that should be shown on one slide.

    {#array;max_rows:10}{your_data}{/array}


### Duplicate PowerPoint slides
If you have a slide which should be duplicated dynamically depending on provided data you can use "multiplier" tag: `$`. Each slide, containing this tag will be as many times duplicated as many entries are stored in the array with the same name.

If there are no data for this tag, this slide will be deleted.

    {$data_to_show_on_multiple_slides}

Here is an example:

Template
![Template sample][pptx-templater-template]

Output
![Output sample][pptx-templater-output]

For data
```
{
    "DATA_TO_SHOW_ON_MULTIPLE_SLIDES": [{
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

Contributions
===
Contributions are greatly appreciated. See [CONTRIBUTING][pptx-templater-contribution] for details


How to obtain support
===
Feel free to open new issues for feature requests, bugs or general feedback on
the [GitHub issues page of this project][pptx-templater-issues].

License
===
Copyright 2016-2021 SAP SE or an SAP affiliate company and utils-pptx-templater contributors. Please see our [LICENSE][pptx-templater-license] for copyright and license information. Detailed information including third-party components and their licensing/copyright information is available [via the REUSE tool][pptx-templater-reuse-tool].

[docxtemplater]:https://github.com/open-xml-templating/docxtemplater
[pptx-templater-template]: https://github.com/sap/utils-pptx-templater/blob/master/img/sample_template.png
[pptx-templater-output]: https://github.com/sap/utils-pptx-templater/blob/master/img/sample_output.png
[pptx-templater-license]: https://github.com/SAP/utils-pptx-templater/blob/master/LICENSE.md
[pptx-templater-contribution]: https://github.com/SAP/utils-pptx-templater/blob/master/CONTRIBUTING.md
[pptx-templater-issues]: https://github.com/SAP/utils-pptx-templater/issues
[pptx-templater-reuse-tool]: https://api.reuse.software/info/github.com/SAP/utils-pptx-templater
