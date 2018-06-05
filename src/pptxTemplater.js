'use strict';

let PptxTemplater;

PptxTemplater =
    class PptxTemplater {
        constructor(zip, tags) {
            this.zip = zip;
            this.tags = tags;
        }

        addSlideWithIndex(content, allSlides, slideIndex, originalSlideIndex) {
            let index = slideIndex - 1;
            if (allSlides[index]) {
                let originalIndexOfSlideInArray = parseInt(allSlides[index].originalIndex);
                let originalIndexOfSlideToAdd = parseInt(originalSlideIndex);
                index = originalIndexOfSlideInArray <= originalIndexOfSlideToAdd ? index + 1 : index;
            } else if (index > allSlides.length) {
                --index;
            }
            allSlides.splice(index, 0, {
                content: content,
                originalIndex: originalSlideIndex
            });
        }

        addSlideRelationships(allSlides, slideRelationshipTemplate, content, relId) {
            let self = this;
            let rId = relId;
            let newRelsContent = content;
            allSlides.forEach((slide, index) => {
                let slideNr = index + 1;
                let relationship = self.changeSlideReference(slideNr, slideRelationshipTemplate);
                newRelsContent = self.replaceOldRel(newRelsContent, rId++, relationship);
            });
            return newRelsContent;
        }

        changeSlideReference(slideNr, content) {
            return content.replace(/slides\/slide\d+/g, 'slides/slide' + slideNr);
        }

        changeNotesSlideReference(slideNr, content) {
            return content.replace(/notesSlides\/notesSlide\d+/g, 'notesSlides/notesSlide' + slideNr);
        }

        collectSlides(content, slideIndex, allSlides, additionalSlideIndex) {
            let originalSlideIndex = slideIndex;
            slideIndex = allSlides.length + 1;
            this.addSlideWithIndex(content, allSlides, slideIndex, originalSlideIndex);
            return allSlides;
        }

        findTags(regex, content) {
            let contentWithoutXMLTags = content.replace(/<.*?>/g, '');
            return regex.exec(contentWithoutXMLTags);
        }

        getFileRelationship(slideIndex, path, fileNamePattern, fileExtention) {
            // <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/<relationshipType>" Target="<path>/<file>.<fileExtention>"/>
            let relationshipRegex = new RegExp('<Relationship[^<]*Target="[^<]*' + fileNamePattern + '(\\d+)' + fileExtention + '"\\/>', 'g');
            let slideRelsPath = 'ppt/slides/_rels/slide' + slideIndex + '.xml.rels';
            let slideRelationshipContent = this.zip.files[slideRelsPath] ? this.zip.files[slideRelsPath].asText() : '';
            let fileRelationship = relationshipRegex.exec(slideRelationshipContent);
            return fileRelationship ? {
                slideRelationship: fileRelationship,
                index: fileRelationship[1],
                path: path + fileNamePattern + fileRelationship[1] + fileExtention
            } : null;
        }

        getIndexFromPath(path) {
            return parseInt(/\d+/g.exec(path)[0]);
        }

        getNotesSlide(originalIndex, originalPath, changedNotesSlides, slideNr) {
            let notesSlideContent = this.zip.files[originalPath].asText();
            let notesSlideIndex = this.getIndexFromPath(originalPath);

            let notesSlidePath = this.getNotesSlidePath(changedNotesSlides, notesSlideIndex);
            if (notesSlidePath !== originalPath) {
                notesSlideIndex = this.getIndexFromPath(notesSlidePath);
                this.updateFile('ppt/slides/_rels/slide' + slideNr + '.xml.rels', 'notesSlide' + originalIndex + '.xml', 'notesSlide' + notesSlideIndex + '.xml');
            }

            return {
                path: notesSlidePath,
                content: notesSlideContent
            };
        }

        getNotesSlidePath(changedNotesSlides, index) {
            let notesSlidePath = 'ppt/notesSlides/notesSlide' + index + '.xml';
            if (changedNotesSlides.indexOf(notesSlidePath) > -1) {
                return this.getNotesSlidePath(changedNotesSlides, ++index);
            } else if (index - 1 > changedNotesSlides.length) {
                return this.getNotesSlidePath(changedNotesSlides, --index);
            }
            return notesSlidePath;
        }

        getTableData(foundTag) {
            let tagName = foundTag[1];
            let maxRows = foundTag[2] ? parseInt(foundTag[2]) : null;
            let data = this.tags[tagName];
            if (!data) return [];
            if (maxRows && data.length > maxRows) {
                return this.splitTableData(data, maxRows);
            } else {
                return [data];
            }
        }

        manipulateTags(content, id, dataTag, data) {
            let indexOfFirstTag = 0;
            let hasOuterTag = false;
            let maxRowsRegex = /;max_rows:\d+/i;
            let newContent = '';
            let substring = content;
            let tagRegex = /\{([^\w]*)([^-]*?)}/g;
            let tags;

            while (tags = this.findTags(tagRegex, content)) {
                let foundTagName = tags[2] ? tags[2].replace(maxRowsRegex, '') : tags[1];

                indexOfFirstTag = substring.indexOf(foundTagName);
                hasOuterTag = tags[1] === '/' ? false : hasOuterTag;
                if (foundTagName === dataTag && tags[1] === '$') {
                    substring = substring.replace(/\{\$[^}]*?}/g, '');
                } else if (!hasOuterTag) {
                    let newTagName = foundTagName;
                    if (data){
                        newTagName = dataTag + '_' + foundTagName + '_' + id;
                        this.tags[newTagName] = data[foundTagName] || data;
                    }

                    newContent += substring.slice(0, indexOfFirstTag);
                    substring = substring.slice(indexOfFirstTag);
                    newContent += newTagName;
                    let oldStringToRemove = /([^{]*)}/g.exec(substring)[1]; // e.g. array;</a:t></a:r><a:r><a:rPr lang=\"en-US\" smtClean=\"0\"/><a:t>max_rows:</a:t></a:r><a:r><a:rPr lang=\"en-US\" smtClean=\"0\"/><a:t>15</a:t></a:r><a:r><a:rPr lang=\"en-US\" dirty=\"0\" smtClean=\"0\"/><a:t>}
                    substring = substring.replace(oldStringToRemove, '');
                } else {
                    newContent += substring.slice(0, indexOfFirstTag + foundTagName.length);
                    substring = substring.slice(indexOfFirstTag + foundTagName.length);
                }
                if (foundTagName === dataTag && tags[1] === '/') {
                    break;
                }
                hasOuterTag = tags[1] === '#' ? true : hasOuterTag;
            }
            newContent += substring;
            return newContent;
        }

        prepareChangedFile(originId, path) {
            let originPath = path.replace(/\d+/g, originId);
            if (this.zip.files[originPath]) {
                return {
                    path: path,
                    content: this.zip.files[originPath].asText()
                };
            }
        }

        reAddFiles(allSlides) {
            let self = this;
            this.updateRelationships(allSlides);
            allSlides.forEach((slide, index) => {
                let slideNr = index + 1;
                let path = 'ppt/slides/slide' + slideNr + '.xml';
                self.zip.file(path, slide.content);
            });
            return this;
        }

        removeSlide(slideNr) {
            let slidePath = 'ppt/slides/slide' + slideNr + '.xml';
            let notesSlidePath = 'ppt/notesSlides/notesSlide' + slideNr + '.xml';
            let slideRelsPath = 'ppt/slides/_rels/slide' + slideNr + '.xml.rels';
            let notesSlideRelsPath = 'ppt/notesSlides/_rels/notesSlide' + slideNr + '.xml.rels';
            delete this.zip.files[slidePath];
            delete this.zip.files[slideRelsPath];
        }

        replaceOldRel(relsContent, id, rel) {
            let newRelsContent;
            let newRelString = rel.replace(/Id="rId\d+"/g, 'Id="rId' + id + '"');
            newRelsContent = relsContent.replace(/><\/Relationships/, '>' + newRelString + '</Relationships');
            return newRelsContent;
        }

        splitBy(type, slideNr, content, allSlides) {
            let self = this;
            let maxRowsTagRegex = /\{\#([^}]+);max_rows:(\d+)}/i;
            let multiplierTagRegex = /\{\$([^}]+)}/g;
            let newTemplateSlides = allSlides;
            let splitterTag = type === 'multiplier' ? this.findTags(multiplierTagRegex, content) : this.findTags(maxRowsTagRegex, content);
            if (splitterTag) {
                let tagName = splitterTag[1];
                let dataForTag = type === 'multiplier' ? this.tags[tagName] : this.getTableData(splitterTag);
                if ((!dataForTag || dataForTag.length === 0) && type === 'multiplier') {
                    newTemplateSlides.splice(slideNr - 1, 1);
                    this.removeSlide(slideNr);
                    return newTemplateSlides;
                } else if (dataForTag.length < 1) {
                    let newSlideContent = self.manipulateTags(content, null, tagName);
                    newTemplateSlides = self.collectSlides(newSlideContent, slideNr, newTemplateSlides);
                } else if (dataForTag && !Array.isArray(dataForTag)) {
                    throw new Error('Data for tag ' + tagName + ' must be an array!');
                }

                dataForTag.forEach((object, index) => {
                    let newSlideContent = self.manipulateTags(content, index, tagName, object);
                    newTemplateSlides = self.collectSlides(newSlideContent, slideNr, newTemplateSlides, index);
                });
            } else {
                newTemplateSlides = this.collectSlides(content, slideNr, newTemplateSlides);
            }
            return newTemplateSlides;
        }

        splitTableData(data, maxRows, newData) {
            newData = newData ? newData : [];
            if (data !== undefined) {
                let dataCopy = data.concat();
                newData.push(dataCopy.splice(0, maxRows));
                dataCopy.length > maxRows ? this.splitTableData(dataCopy, maxRows, newData) : newData.push(dataCopy);
            }
            return newData;
        }

        splitTemplateSlides() {
            let newTemplateSlides = [];
            let slideIndex = 1;
            let fileName = 'ppt/slides/slide' + slideIndex + '.xml';
            while (this.zip.files[fileName]) {
                let content = this.zip.files[fileName].asText();
                newTemplateSlides = this.splitBy('multiplier', slideIndex, content, newTemplateSlides);
                fileName = 'ppt/slides/slide' + ++slideIndex + '.xml';
            }
            let allTemplateSlides = [];
            newTemplateSlides.forEach((slide) => {
                allTemplateSlides = this.splitBy('maxRows', slide.originalIndex, slide.content, allTemplateSlides);
            });
            return allTemplateSlides;
        }

        updateContentTypes(slides, part) {
            // <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
            let self = this;
            let contentTypesContext = this.zip.files['[Content_Types].xml'].asText();
            let newContentTypesContext = contentTypesContext;
            let override;
            let overrideIndex = -1;
            let overrideSlideRegex = /<Override PartName="\/ppt\/slides\/slide[^<]*\.slide\+xml"\/>/g;
            let overrideNotesSlideRegex = /<Override PartName="\/ppt\/notesSlides\/notesSlide[^<]*\.notesSlide\+xml"\/>/g;
            let overrideRegex = part == 'slide' ? overrideSlideRegex : overrideNotesSlideRegex;
            let overrideTemplate = '';

            while (override = overrideRegex.exec(contentTypesContext)) {
                overrideIndex = overrideIndex < 0 ? override.index : overrideIndex;
                overrideTemplate = overrideTemplate === '' ? override[0] : overrideTemplate;
                newContentTypesContext = newContentTypesContext.replace(override[0], '');
            }
            slides.forEach((slide, index) => {
                let slideNr = index + 1;
                let newOverride = part == 'slide' ? self.changeSlideReference(slideNr, overrideTemplate) : self.changeNotesSlideReference(slideNr, overrideTemplate);
                newContentTypesContext = newContentTypesContext.slice(0, overrideIndex) + newOverride + newContentTypesContext.slice(overrideIndex);
                overrideIndex += newOverride.length;
            });

            return newContentTypesContext;
        }

        updateFile(path, regex, newContent) {
            let fileContent = newContent;
            if (this.zip.files[path] && regex) {
                fileContent = this.zip.files[path].asText();
                fileContent = fileContent.replace(regex, newContent);
            }
            this.zip.file(path, fileContent);
        }

        updateNotesSlides(allSlides) {
            let self = this;
            let newNotesSlides = [];
            let changedNotesSlides = [];

            allSlides.forEach((slide, index) => {
                let slideNr = index + 1;
                let notesSlideRelationship = self.getFileRelationship(slideNr, 'ppt/notesSlides/', 'notesSlide', '.xml');
                if (notesSlideRelationship) {
                    let notesSlide = self.getNotesSlide(notesSlideRelationship.index, notesSlideRelationship.path, changedNotesSlides, slideNr);

                    changedNotesSlides.push(notesSlide.path);
                    newNotesSlides.push(notesSlide);
                }
            });
            this.updateSlideRelFiles(newNotesSlides, 'ppt/notesSlides/_rels/notesSlide1.xml.rels');
            newNotesSlides.forEach((notesSlide) => {
                self.updateFile(notesSlide.path, null, notesSlide.content);
            });

            return newNotesSlides;
        }

        updatePresentationRelationships(allSlides, slideStartId, changedRelationshipIds) {
            let newContext;
            let presentationContext = this.zip.files['ppt/presentation.xml'].asText();

            newContext = this.updatePresentationRelTags(presentationContext, changedRelationshipIds);
            return this.updatePresentationSlideList(newContext, allSlides, slideStartId);
        }

        updatePresentationRelTags(context, changedRelationshipIds) {
            let newContext = context;
            changedRelationshipIds.forEach((change) => {
                let tagToChange = new RegExp('<[^>]*' + change.oldRid).exec(context);
                if (tagToChange) {
                    change.tagToChange = tagToChange[0];
                }
            });
            changedRelationshipIds.forEach((change) => {
                if (change.tagToChange) {
                    let tagWithNewRid = change.tagToChange.replace(/rId\d+/g, change.newRid);
                    newContext = newContext.replace(change.tagToChange, tagWithNewRid);
                }
            });
            return newContext;
        }

        updatePresentationSlideList(context, allSlides, startId) {
            // <p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/><p:sldId id="259" r:id="rId5"/></p:sldIdLst>
            let slideList = /<p:sldIdLst>.*<\/p:sldIdLst>/g.exec(context)[0];
            let firstId = parseInt(/id="(\d+)"/g.exec(slideList)[1]);
            let newSlideList = '<p:sldIdLst>';
            for (let i = 0; i < allSlides.length; i++) {
                newSlideList += '<p:sldId id="' + (firstId + i) + '" r:id="rId' + (startId + i) + '"/>';
            }
            newSlideList += '</p:sldIdLst>';

            return context.replace(slideList, newSlideList);
        }

        updateRelationships(allSlides) {
            // <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
            let self = this;
            let relsContent = this.zip.files['ppt/_rels/presentation.xml.rels'].asText();
            let newRelsContent;
            let relationshipRegex = /<Relationship Id=[^<]*\/>/g;
            let allRelationshipsRegex = /<Relationship Id=.*\/>/g;
            let rel;
            let rId = 0;
            let slideRelationshipTemplate = '';
            let changedRelationships = [];

            newRelsContent = relsContent.replace(allRelationshipsRegex, '');

            while (rel = relationshipRegex.exec(relsContent)) {
                let relationship = rel[0];
                if (relationship.indexOf('Target="slides/slide') == -1) {
                    newRelsContent = self.replaceOldRel(newRelsContent, ++rId, relationship);
                    changedRelationships.push({
                        oldRid: /rId\d+/g.exec(relationship)[0],
                        newRid: 'rId' + rId
                    });
                } else {
                    slideRelationshipTemplate = relationship;
                }
            }

            this.zip.file('ppt/_rels/presentation.xml.rels', this.addSlideRelationships(allSlides, slideRelationshipTemplate, newRelsContent, ++rId));
            this.zip.file('[Content_Types].xml', this.updateContentTypes(allSlides, 'slide'));
            this.zip.file('ppt/presentation.xml', this.updatePresentationRelationships(allSlides, rId, changedRelationships));
            this.updateSlideRelFiles(allSlides, 'ppt/slides/_rels/slide1.xml.rels');
            let allNotesSlides = this.updateNotesSlides(allSlides);
            this.zip.file('[Content_Types].xml', this.updateContentTypes(allNotesSlides, 'notesSlide'));
        }

        updateSlideRelFiles(allSlides, path) {
            let self = this;
            let newRelsFiles = [];
            allSlides.forEach((slide, index) => {
                let slideNr = index + 1;
                path = path.replace(/\d+/g, slideNr);
                let newFile = self.prepareChangedFile(slide.originalIndex, path);
                if (newFile) {
                    newRelsFiles.push(newFile);
                }
            });
            newRelsFiles.forEach((file) => {
                this.zip.file(file.path, file.content);
            });
        }
    };

module.exports = PptxTemplater;
