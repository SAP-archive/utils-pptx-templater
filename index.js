'use strict';
let PptxTemplater = require('./src/pptxTemplater');

let PptxTemplaterModule = class PptxTemplaterModule {
    constructor(options) {
        this.options = options | {};
        this.name = 'pptxTemplater';
    }
    handleEvent(event, eventData) {
        let gen, newTemplateSlides;
        if (event === 'rendering') {
            this.renderingFileName = eventData;
            gen = this.manager.getInstance('gen');
            this.pptxTemplater = new PptxTemplater(gen.zip, gen.tags);
            newTemplateSlides = this.pptxTemplater.splitTemplateSlides();
            return this.pptxTemplater.reAddFiles(newTemplateSlides);
        } else if (event === 'rendered') {
            return this.finished();
        }
    }

    get() {
        return null;
    }

    handle() {
        return null;
    }

    finished() {
        return null;
    }
};

module.exports = new PptxTemplaterModule();
