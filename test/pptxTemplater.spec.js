'use strict';

let assert = require('assert');
let Docxtemplater = require('docxtemplater');
let fs = require('fs');
let path = require( 'path');
let PptxTemplater = require('./../src/pptxTemplater.js');
let pptxTemplater = new PptxTemplater(null, null);
let pptxtemplaterModule = require('./../index.js');
let sinon = require('sinon');
let templatePath = 'template.pptx';
let templateZip = null;
let testData = require('./testData.json');

// function checkFileExists(filepath){
//     return new Promise((resolve) => {
//         fs.access(filepath, fs.F_OK, e => {
//             resolve(!e);
//         });
//     });
// }
function fillTemplate(data, templatePath) {
    let content = fs.readFileSync(path.join(__dirname, templatePath), 'binary');
    let doc = new Docxtemplater(content);
    doc.setOptions({ fileType: 'pptx' });
    doc.templatedFiles = doc.fileTypeConfig.getTemplatedFiles(doc.zip);
    doc.setData(data);

    return doc;
}

function getSplittedSlides(slideNr) {
    if (slideNr) {
        return pptxTemplater.splitTemplateSlides.returnValues[0][slideNr - 1];
    }
    return pptxTemplater.splitTemplateSlides.returnValues[0];
};

sinon.stub(pptxtemplaterModule, 'getPptxTemplater', (zip, tags) => {
    templateZip = zip.clone();
    pptxTemplater.zip = zip;
    pptxTemplater.tags = tags;
    return pptxTemplater;
});
sinon.spy(pptxTemplater, 'splitTemplateSlides');

describe('pptxtemplater', () => {
    let pptx;
    beforeEach(() => {
        pptx = fillTemplate(testData, templatePath);
        pptx.attachModule(pptxtemplaterModule);

    });
    describe('on rendering event', () => {
        it('should split slides with multiplier tag $', () => {
            pptx.render();

            assert(pptxTemplater.zip !== null);
            assert(getSplittedSlides().length > 4);
            assert(getSplittedSlides(4).content.indexOf('text') > -1);
            assert(getSplittedSlides(5).content.indexOf('text') > -1);
        });
        xit('should increase the number of notesSlides', () => {
            function countFiles(files, path) {
                let count = 0;
                while (files[path + (count + 1) + '.xml']) {
                    count++;
                }
                return count;
            }

            assert.equal(countFiles(templateZip.files, 'ppt/notesSlides/notesSlide'), 1);
            pptx.render();

            assert.equal(countFiles(pptx.zip.files, 'ppt/notesSlides/notesSlide'), 2);
        });

    });
});
