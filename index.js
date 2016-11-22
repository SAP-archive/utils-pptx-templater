'use strict';
let PptxTemplater = require('./src/pptxTemplater');

let PptxTemplaterModule = class PptxTemplaterModule {
    constructor(options) {
        this.options = options | {};
        this.name = 'pptxTemplater';
    }

    get() {
        return null;
    }

    getPptxTemplater(zip, tags) {
        return new PptxTemplater(zip, tags);
    }

    handleEvent(event, eventData) {
        let gen, newTemplateSlides;
        if (event === 'rendering') {
            this.renderingFileName = eventData;
            gen = this.manager.getInstance('gen');
            this.pptxTemplater = this.getPptxTemplater(gen.zip, gen.tags);
            newTemplateSlides = this.pptxTemplater.splitTemplateSlides();
            this.pptxTemplater.reAddFiles(newTemplateSlides);
            return gen.templatedFiles = gen.fileTypeConfig.getTemplatedFiles(gen.zip);
        } else if (event === 'rendered') {
            return this.finished();
        }
    }

    handle() {
        return null;
    }

    finished() {
        return null;
    }
};

module.exports = new PptxTemplaterModule();
