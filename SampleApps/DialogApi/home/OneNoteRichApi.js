var OneNoteRichApi;
(function (OneNoteRichApi) {
    var EntityProperties = (function () {
        function EntityProperties() {
            var notebookPropertiesArray = EntityProperties.notebookProperties;
            this.notebookProperties = EntityProperties.joinWithSeparator(notebookPropertiesArray);
            var sectionPropertiesArray = EntityProperties.sectionProperties.concat(EntityProperties.addPrefix(notebookPropertiesArray, "notebook"));
            this.sectionProperties = EntityProperties.joinWithSeparator(sectionPropertiesArray);
            var sectionGroupPropertiesArray = EntityProperties.sectionGroupProperties.concat(EntityProperties.addPrefix(notebookPropertiesArray, "notebook"));
            this.sectionGroupProperties = EntityProperties.joinWithSeparator(sectionGroupPropertiesArray);
            var pagePropertiesArray = EntityProperties.pageProperties.concat(EntityProperties.addPrefix(sectionPropertiesArray, "parentSection"));
            this.pageProperties = EntityProperties.joinWithSeparator(pagePropertiesArray);
            var pageContentPropertiesArray = EntityProperties.pageContentProperties.concat(EntityProperties.addPrefix(pagePropertiesArray, "parentPage"));
            this.pageContentProperties = EntityProperties.joinWithSeparator(pageContentPropertiesArray);
            var activeOutlinePropertiesArray = EntityProperties.outlineProperties
                .concat(EntityProperties.addPrefix(pageContentPropertiesArray, "pageContent"))
                .concat(EntityProperties.addPrefix(EntityProperties.paragraphProperties, "paragraph"));
            this.activeOutlineProperties = EntityProperties.joinWithSeparator(activeOutlinePropertiesArray);
            var paragraphPropertiesArray = EntityProperties.paragraphProperties.concat(EntityProperties.addPrefix(activeOutlinePropertiesArray, "outline"));
            this.paragraphProperties = EntityProperties.joinWithSeparator(paragraphPropertiesArray);
            var richTextPropertiesArray = EntityProperties.richTextProperties.concat(EntityProperties.addPrefix(paragraphPropertiesArray, "paragraph"));
            this.richTextProperties = EntityProperties.joinWithSeparator(richTextPropertiesArray);
            var tablePropertiesArray = EntityProperties.tableProperties.concat(EntityProperties.addPrefix(paragraphPropertiesArray, "paragraph"));
            this.tableProperties = EntityProperties.joinWithSeparator(tablePropertiesArray);
            var tableRowPropertiesArray = EntityProperties.tableRowProperties.concat(EntityProperties.addPrefix(tablePropertiesArray, "parentTable"));
            this.tableRowProperties = EntityProperties.joinWithSeparator(tableRowPropertiesArray);
            var tableCellPropertiesArray = EntityProperties.tableCellProperties.concat(EntityProperties.addPrefix(tableRowPropertiesArray, "parentRow"));
            this.tableCellProperties = EntityProperties.joinWithSeparator(tableCellPropertiesArray);
            var imagePropertiesArray = EntityProperties.imageProperties; // TODO: How do we approach images, since they can have multiple parents? (either in paragraph or outline)
            this.imageProperties = EntityProperties.joinWithSeparator(imagePropertiesArray);
            this.pageContentImageProperties = EntityProperties.joinWithSeparator([
                EntityProperties.joinWithSeparator(EntityProperties.addPrefix(EntityProperties.imageProperties, "image"))
            ]);
            this.pageContentOutlineProperties = EntityProperties.joinWithSeparator([
                EntityProperties.joinWithSeparator(EntityProperties.addPrefix(EntityProperties.outlineProperties, "outline"))
            ]);
        }
        EntityProperties.joinWithSeparator = function (arr) {
            return arr.join(",");
        };
        EntityProperties.addPrefix = function (arr, prefix) {
            var res = [];
            for (var i = 0; i < arr.length; i++) {
                res.push(prefix + "/" + arr[i]);
            }
            return res;
        };
        EntityProperties.getNotebookProperties = function () {
            return EntityProperties.singleton.notebookProperties;
        };
        EntityProperties.getSectionProperties = function () {
            return EntityProperties.singleton.sectionProperties;
        };
        EntityProperties.getPageProperties = function () {
            return EntityProperties.singleton.pageProperties;
        };
        EntityProperties.getPageContentProperties = function () {
            return EntityProperties.singleton.pageContentProperties;
        };
        EntityProperties.getActiveOutlineProperties = function () {
            return EntityProperties.singleton.activeOutlineProperties;
        };
        EntityProperties.getParagraphProperties = function () {
            return EntityProperties.singleton.paragraphProperties;
        };
        EntityProperties.getTableProperties = function () {
            return EntityProperties.singleton.tableProperties;
        };
        EntityProperties.getImageProperties = function () {
            return EntityProperties.singleton.imageProperties;
        };
        EntityProperties.getRichTextProperties = function () {
            return EntityProperties.singleton.richTextProperties;
        };
        EntityProperties.getPageContentPropertiesWithImage = function () {
            return EntityProperties.singleton.pageContentImageProperties;
        };
        EntityProperties.getPageContentPropertiesWithOutline = function () {
            return EntityProperties.singleton.pageContentOutlineProperties;
        };
        EntityProperties.getTableRowProperties = function () {
            return EntityProperties.singleton.tableRowProperties;
        };
        EntityProperties.getTableCellProperties = function () {
            return EntityProperties.singleton.tableCellProperties;
        };
        EntityProperties.getNotebookStructureProperties = function (includePages) {
            var pagePropertiesArrayForNotebookStructure = includePages ? EntityProperties.pageProperties : [];
            var sectionPropertiesArrayForNotebookStructure = EntityProperties.sectionProperties.concat(EntityProperties.addPrefix(pagePropertiesArrayForNotebookStructure, "pages"));
            var sectionGroupPropertiesArrayForNotebookStructure = EntityProperties.sectionGroupProperties.concat(EntityProperties.addPrefix(sectionPropertiesArrayForNotebookStructure, "sections"));
            var sectionGroupPropertiesArrayForNotebookStructure2 = EntityProperties.sectionGroupProperties.concat(EntityProperties.addPrefix(sectionPropertiesArrayForNotebookStructure, "sections"))
                .concat(EntityProperties.addPrefix(sectionGroupPropertiesArrayForNotebookStructure, "sectionGroups"));
            var sectionGroupPropertiesArrayForNotebookStructure3 = EntityProperties.sectionGroupProperties.concat(EntityProperties.addPrefix(sectionPropertiesArrayForNotebookStructure, "sections"))
                .concat(EntityProperties.addPrefix(sectionGroupPropertiesArrayForNotebookStructure2, "sectionGroups"));
            var sectionGroupPropertiesArrayForNotebookStructure4 = EntityProperties.sectionGroupProperties.concat(EntityProperties.addPrefix(sectionPropertiesArrayForNotebookStructure, "sections"))
                .concat(EntityProperties.addPrefix(sectionGroupPropertiesArrayForNotebookStructure3, "sectionGroups"));
            var notebookStructurePropertiesArray = EntityProperties.notebookProperties.concat(EntityProperties.addPrefix(sectionPropertiesArrayForNotebookStructure, "sections"))
                .concat(EntityProperties.addPrefix(sectionGroupPropertiesArrayForNotebookStructure4, "sectionGroups"));
            var notebookStructureProperties = EntityProperties.joinWithSeparator(notebookStructurePropertiesArray);
            return notebookStructureProperties;
        };
        EntityProperties.getImagePropertiesOnly = function () {
            return EntityProperties.joinWithSeparator(EntityProperties.imageProperties);
        };
        EntityProperties.getRichTextPropertiesOnly = function () {
            return EntityProperties.joinWithSeparator(EntityProperties.richTextProperties);
        };
        EntityProperties.getParagraphPropertiesOnly = function () {
            return EntityProperties.joinWithSeparator(EntityProperties.paragraphProperties);
        };
        EntityProperties.getOutlinePropertiesOnly = function () {
            return EntityProperties.joinWithSeparator(EntityProperties.outlineProperties
                .concat(EntityProperties.addPrefix(EntityProperties.paragraphProperties, "paragraphs")));
        };
        EntityProperties.getTableRowPropertiesOnly = function () {
            return EntityProperties.joinWithSeparator(EntityProperties.tableRowProperties
                .concat(EntityProperties.addPrefix(EntityProperties.tableCellProperties, "cells"))
                .concat(EntityProperties.addPrefix(EntityProperties.paragraphProperties, "cells/paragraphs")));
        };
        EntityProperties.getPageContentPropertiesOnly = function () {
            return EntityProperties.joinWithSeparator(EntityProperties.pageContentProperties);
        };
        EntityProperties.getTablePropertiesOnly = function () {
            return EntityProperties.joinWithSeparator(EntityProperties.tableProperties
                .concat(EntityProperties.addPrefix(EntityProperties.tableRowProperties, "rows"))
                .concat(EntityProperties.addPrefix(EntityProperties.tableCellProperties, "rows/cells"))
                .concat(EntityProperties.addPrefix(EntityProperties.paragraphProperties, "rows/cells/paragraphs")));
        };
        EntityProperties.notebookProperties = ["id", "name", "clientUrl"];
        EntityProperties.sectionGroupProperties = ["id", "name", "clientUrl"];
        EntityProperties.sectionProperties = ["id", "name", "clientUrl"];
        EntityProperties.pageProperties = ["id", "title", "pageLevel", "clientUrl"];
        EntityProperties.pageContentProperties = ["id", "left", "top", "type"];
        EntityProperties.outlineProperties = ["id"];
        EntityProperties.richTextProperties = ["id", "text"];
        EntityProperties.paragraphProperties = ["id", "type"];
        EntityProperties.imageProperties = ["id", "description", "height", "hyperlink", "width"];
        EntityProperties.tableProperties = ["id", "columnCount", "rowCount"];
        EntityProperties.tableRowProperties = ["id", "cellCount", "rowIndex"];
        EntityProperties.tableCellProperties = ["id", "cellIndex", "rowIndex"];
        EntityProperties.singleton = new EntityProperties();
        return EntityProperties;
    }());
    OneNoteRichApi.EntityProperties = EntityProperties;
})(OneNoteRichApi || (OneNoteRichApi = {}));
/// <reference path="EntityProperties.ts"/>
/// <reference path="Entities.ts"/>
var OneNoteRichApi;
(function (OneNoteRichApi) {
    var EntityBuilder = (function () {
        function EntityBuilder() {
        }
        // Build a notebook object from a notebook proxy
        EntityBuilder.buildNotebookFromNotebookProxy = function (notebookProxy) {
            return new OneNoteRichApi.Notebook(notebookProxy.id, notebookProxy.name, notebookProxy.clientUrl);
        };
        // Build a section object from a section proxy
        EntityBuilder.buildSectionFromSectionProxyWithNotebook = function (sectionProxy, notebook) {
            return new OneNoteRichApi.Section(sectionProxy.id, sectionProxy.name, sectionProxy.clientUrl, notebook);
        };
        EntityBuilder.buildSectionFromSectionProxy = function (sectionProxy) {
            var notebook = EntityBuilder.buildNotebookFromNotebookProxy(sectionProxy.notebook);
            return EntityBuilder.buildSectionFromSectionProxyWithNotebook(sectionProxy, notebook);
        };
        // Build a section object from a section proxy
        EntityBuilder.buildSectionGroupFromSectionGroupProxyWithNotebook = function (sectionGroupProxy, notebook) {
            return new OneNoteRichApi.SectionGroup(sectionGroupProxy.id, sectionGroupProxy.name, sectionGroupProxy.clientUrl, notebook);
        };
        EntityBuilder.buildSectionGroupFromSectionGroupProxy = function (sectionGroupProxy) {
            var notebook = EntityBuilder.buildNotebookFromNotebookProxy(sectionGroupProxy.notebook);
            return EntityBuilder.buildSectionGroupFromSectionGroupProxyWithNotebook(sectionGroupProxy, notebook);
        };
        // Build a page object from a page proxy
        EntityBuilder.buildPageFromPageProxyWithParentSection = function (pageProxy, parentSection) {
            return new OneNoteRichApi.Page(pageProxy.id, pageProxy.title, pageProxy.pageLevel, pageProxy.clientUrl, parentSection);
        };
        EntityBuilder.buildPageFromPageProxy = function (pageProxy) {
            var section = EntityBuilder.buildSectionFromSectionProxy(pageProxy.parentSection);
            return this.buildPageFromPageProxyWithParentSection(pageProxy, section);
        };
        // Build a page content proxy from a page content object
        EntityBuilder.buildPageContentFromPageContentProxyWithPage = function (pageContentProxy, parentPage) {
            return new OneNoteRichApi.PageContent(pageContentProxy.id, pageContentProxy.left, pageContentProxy.top, pageContentProxy.type, parentPage);
        };
        EntityBuilder.buildPageContentFromPageContentProxy = function (pageContentProxy) {
            var page = EntityBuilder.buildPageFromPageProxy(pageContentProxy.parentPage);
            return EntityBuilder.buildPageContentFromPageContentProxyWithPage(pageContentProxy, page);
        };
        // Build an image object from an image proxy
        // An image might have a parent paragraph or a parent pageContent
        EntityBuilder.buildImageFromImageProxyWithParent = function (imageProxy, parentPageContent, parentParagraph) {
            var image = new OneNoteRichApi.Image(imageProxy.id, imageProxy.height, imageProxy.width, imageProxy.description, imageProxy.hyperlink);
            image.parentPagecontent = parentPageContent;
            image.parentParagraph = parentParagraph;
            return image;
        };
        // Build an outline object from an outline proxy
        EntityBuilder.buildOutlineFromOutlineProxyWithParentPageContent = function (outlineProxy, pageContent) {
            return new OneNoteRichApi.Outline(outlineProxy.id, pageContent);
        };
        EntityBuilder.buildOutlineFromOutlineProxy = function (outlineProxy) {
            var pageContentObject = EntityBuilder.buildPageContentFromPageContentProxy(outlineProxy.pageContent);
            return EntityBuilder.buildOutlineFromOutlineProxyWithParentPageContent(outlineProxy, pageContentObject);
        };
        // Build a paragraph object from a paragraph proxy
        EntityBuilder.buildParagraphFromParagraphProxyWithOutline = function (paragraphProxy, parentOutline, parentTableCell) {
            var paragraph = new OneNoteRichApi.Paragraph(paragraphProxy.id, paragraphProxy.type);
            if (parentOutline) {
                paragraph.parentOutline = parentOutline;
            }
            else if (parentTableCell) {
                paragraph.parentTableCell = parentTableCell;
            }
            else {
                throw new Error("Not implemented");
            }
            return paragraph;
        };
        EntityBuilder.buildParagraphFromParagraphProxy = function (paragraphProxy) {
            var outline = EntityBuilder.buildOutlineFromOutlineProxy(paragraphProxy.outline);
            return EntityBuilder.buildParagraphFromParagraphProxyWithOutline(paragraphProxy, outline, null);
        };
        // Build a rich text object from a rich text proxy
        EntityBuilder.buildRichTextFromRichTextProxyWithParagraph = function (richTextProxy, parentParagraph) {
            return new OneNoteRichApi.RichText(richTextProxy.id, richTextProxy.text, parentParagraph);
        };
        EntityBuilder.buildRichTextFromRichTextProxy = function (richTextProxy) {
            var paragraph = EntityBuilder.buildParagraphFromParagraphProxy(richTextProxy.paragraph);
            return EntityBuilder.buildRichTextFromRichTextProxyWithParagraph(richTextProxy, paragraph);
        };
        // Build a table object from a table proxy
        EntityBuilder.buildTableFromTableProxyWithParagraph = function (tableProxy, parentParagraph) {
            return new OneNoteRichApi.Table(tableProxy.id, tableProxy.columnCount, tableProxy.rowCount, parentParagraph);
        };
        EntityBuilder.buildTableFromTableProxy = function (tableProxy) {
            var paragraph = EntityBuilder.buildParagraphFromParagraphProxy(tableProxy.paragraph);
            return EntityBuilder.buildTableFromTableProxyWithParagraph(tableProxy, paragraph);
        };
        // Build a table row object from a table row proxy
        EntityBuilder.buildTableRowFromTableRowProxyWithTable = function (tableRowProxy, parentTable) {
            return new OneNoteRichApi.TableRow(tableRowProxy.id, tableRowProxy.cellCount, tableRowProxy.rowIndex, parentTable);
        };
        EntityBuilder.buildTableRowFromTableRowProxy = function (tableRowProxy) {
            var table = EntityBuilder.buildTableFromTableProxy(tableRowProxy.table);
            return EntityBuilder.buildTableRowFromTableRowProxyWithTable(tableRowProxy, table);
        };
        // Build a table row object from a table row proxy
        EntityBuilder.buildTableCellFromTableCellProxyWithTableRow = function (tableCellProxy, parentTableRow) {
            return new OneNoteRichApi.TableCell(tableCellProxy.id, tableCellProxy.rowIndex, tableCellProxy.cellIndex, parentTableRow);
        };
        EntityBuilder.buildTableCellFromTableCellProxy = function (tableCellProxy) {
            var tableRow = EntityBuilder.buildTableRowFromTableRowProxy(tableCellProxy.parentRow);
            return EntityBuilder.buildTableCellFromTableCellProxyWithTableRow(tableCellProxy, tableRow);
        };
        return EntityBuilder;
    }());
    OneNoteRichApi.EntityBuilder = EntityBuilder;
})(OneNoteRichApi || (OneNoteRichApi = {}));
var OneNoteRichApi;
(function (OneNoteRichApi) {
    // Manage host object tracking - this is crucial to avoid proxy object lookup
    var TrackedHostObjects = (function () {
        function TrackedHostObjects() {
            this.trackedObjects = {};
        }
        // Adds a proxy object to the list of tracked objects
        TrackedHostObjects.prototype.add = function (id, entityToTrack) {
            this.trackedObjects[id] = entityToTrack;
        };
        // Disconnect and dispose all tracked proxy objects
        TrackedHostObjects.prototype.close = function () {
            for (var i = 0; i < this.trackedObjects.length; i++) {
                OneNoteRichApi.StaticMethods.trackedHostObjects[i].context.trackedObjects.remove(OneNoteRichApi.StaticMethods.trackedHostObjects[i]);
                OneNoteRichApi.StaticMethods.trackedHostObjects[i].context.sync();
            }
            this.trackedObjects = {};
        };
        // IMPORTANT NOTE:
        // If we had a context.application.entity.getById or getByPath method, we'd be able to make this call in a single ctx.sync
        // Instead, we have to ennumerate through some of the objects to find the proxy to act on
        // Manoj is investigating the feasibility of having these methods. I'd appreciate input here.
        TrackedHostObjects.prototype.getNotebookProxyFromNotebookAsync = function (notebook) {
            // Serve the proxy from tracked objects, which should always be the case 
            // This saves the cost of fetching notebook entities
            if (notebook.id in this.trackedObjects) {
                return Promise.resolve(this.trackedObjects[notebook.id]);
            }
            else {
                throw new Error("Not implemented");
            }
        };
        TrackedHostObjects.prototype.getSectionProxyFromSectionAsync = function (section) {
            // Serve the proxy from tracked objects, which should always be the case
            // This saves the cost of fetching section entities
            if (section.id in this.trackedObjects) {
                return Promise.resolve(this.trackedObjects[section.id]);
            }
            else {
                throw new Error("Not implemented");
            }
        };
        TrackedHostObjects.prototype.getPageProxyFromPageAsync = function (page) {
            // Serve the proxy from tracked objects, which should always be the case
            // This saves the cost of fetching page entities
            if (page.id in this.trackedObjects) {
                return Promise.resolve(this.trackedObjects[page.id]);
            }
            else {
                throw new Error("Not implemented");
            }
        };
        TrackedHostObjects.prototype.getPageContentProxyFromPageContentAsync = function (pageContent) {
            // Serve the proxy from tracked objects, which should always be the case
            // This saves the cost of fetching page entities
            if (pageContent.id in this.trackedObjects) {
                return Promise.resolve(this.trackedObjects[pageContent.id]);
            }
            else {
                throw new Error("Not implemented");
            }
        };
        TrackedHostObjects.prototype.getOutlineProxyFromOutlineAsync = function (outline) {
            // Serve the proxy from tracked objects, which should always be the case
            // This saves the cost of loading and syncing from the host
            if (outline.id in this.trackedObjects) {
                return Promise.resolve(this.trackedObjects[outline.id]);
            }
            else {
                throw new Error("Not implemented");
            }
        };
        TrackedHostObjects.prototype.getImageProxyFromImageAsync = function (image) {
            // Serve the proxy from tracked objects, which should always be the case
            // This saves the cost of loading and syncing from the host
            if (image.id in this.trackedObjects) {
                return Promise.resolve(this.trackedObjects[image.id]);
            }
            else {
                throw new Error("Not implemented");
            }
        };
        TrackedHostObjects.prototype.getParagraphProxyFromParagraph = function (paragraph) {
            // Serve the proxy from tracked objects, which should always be the case
            // This saves the cost of loading and syncing from the host
            if (paragraph.id in this.trackedObjects) {
                return Promise.resolve(this.trackedObjects[paragraph.id]);
            }
            else {
                throw new Error("Not implemented");
            }
        };
        TrackedHostObjects.prototype.getTableProxyFromTable = function (table) {
            // Serve the proxy from tracked objects, which should always be the case
            // This saves the cost of loading and syncing from the host
            if (table.id in this.trackedObjects) {
                return Promise.resolve(this.trackedObjects[table.id]);
            }
            else {
                throw new Error("Not implemented");
            }
        };
        TrackedHostObjects.prototype.getTableRowProxyFromTableRow = function (tableRow) {
            // Serve the proxy from tracked objects, which should always be the case
            // This saves the cost of loading and syncing from the host
            if (tableRow.id in this.trackedObjects) {
                return Promise.resolve(this.trackedObjects[tableRow.id]);
            }
            else {
                throw new Error("Not implemented");
            }
        };
        TrackedHostObjects.prototype.getTableCellProxyFromTableCell = function (tableCell) {
            // Serve the proxy from tracked objects, which should always be the case
            // This saves the cost of loading and syncing from the host
            if (tableCell.id in this.trackedObjects) {
                return Promise.resolve(this.trackedObjects[tableCell.id]);
            }
            else {
                throw new Error("Not implemented");
            }
        };
        TrackedHostObjects.prototype.getoutlineProxiesFromOutlinesAsync = function (outlines) {
            var outlineProxies = [];
            for (var i = 0; i < outlines.length; i++) {
                var outline = outlines[i];
                if (outline.id in this.trackedObjects) {
                    var outlineProxy = this.trackedObjects[outline.id];
                    outlineProxy.outlineObject = outline;
                    outlineProxies.push(outlineProxy);
                }
                else {
                    throw new Error("Not implemented");
                }
            }
            return Promise.resolve(outlineProxies);
        };
        TrackedHostObjects.prototype.getParagraphProxiesFromParagraphs = function (paragraphs) {
            var paragraphProxies = [];
            for (var i = 0; i < paragraphs.length; i++) {
                var paragraph = paragraphs[i];
                if (paragraph.id in this.trackedObjects) {
                    var paragraphProxy = this.trackedObjects[paragraph.id];
                    paragraphProxy.outlineObject = paragraph.parentOutline;
                    paragraphProxy.tableCellObject = paragraph.parentTableCell;
                    paragraphProxies.push(paragraphProxy);
                }
                else {
                    throw new Error("Not implemented");
                }
            }
            return Promise.resolve(paragraphProxies);
        };
        return TrackedHostObjects;
    }());
    OneNoteRichApi.TrackedHostObjects = TrackedHostObjects;
})(OneNoteRichApi || (OneNoteRichApi = {}));
/// <reference path="Constants.ts"/>
/// <reference path="EntityProperties.ts"/>
/// <reference path="EntityBuilder.ts"/>
/// <reference path="OneNote.d.ts"/>
/// <reference path="TrackedHostObjects.ts"/>
var OneNoteRichApi;
(function (OneNoteRichApi) {
    // These are the top level objects - the developer should only need to call these in case they are
    // serializing / deserializing entities
    var StaticMethods = (function () {
        function StaticMethods() {
        }
        // Clear the cache
        StaticMethods.close = function () {
            this.trackedHostObjects.close();
        };
        // REGION: Application methods
        StaticMethods.getActiveNotebookAsync = function () {
            return OneNote.run(function (context) {
                var notebookProxy = context.application.getActiveNotebookOrNull();
                notebookProxy.load(OneNoteRichApi.EntityProperties.getNotebookProperties());
                context.trackedObjects.add(notebookProxy);
                return context.sync()
                    .then(function () {
                    // Proxy objects can be null - abstract this complexity
                    if (notebookProxy.isNull) {
                        return null;
                    }
                    StaticMethods.trackedHostObjects.add(notebookProxy.id, notebookProxy);
                    return OneNoteRichApi.EntityBuilder.buildNotebookFromNotebookProxy(notebookProxy);
                });
            });
        };
        StaticMethods.getActiveSectionAsync = function () {
            return OneNote.run(function (context) {
                var sectionProxy = context.application.getActiveSectionOrNull();
                sectionProxy.load(OneNoteRichApi.EntityProperties.getSectionProperties());
                context.trackedObjects.add(sectionProxy);
                return context.sync()
                    .then(function () {
                    // Proxy objects can be null - abstract this complexity
                    if (sectionProxy.isNull) {
                        return null;
                    }
                    StaticMethods.trackedHostObjects.add(sectionProxy.id, sectionProxy);
                    return OneNoteRichApi.EntityBuilder.buildSectionFromSectionProxy(sectionProxy);
                });
            });
        };
        StaticMethods.getActivePageAsync = function () {
            return OneNote.run(function (context) {
                var pageProxy = context.application.getActivePageOrNull();
                pageProxy.load(OneNoteRichApi.EntityProperties.getPageProperties());
                context.trackedObjects.add(pageProxy);
                return context.sync()
                    .then(function () {
                    // Proxy objects can be null - abstract this complexity
                    if (pageProxy.isNull) {
                        return null;
                    }
                    StaticMethods.trackedHostObjects.add(pageProxy.id, pageProxy);
                    return OneNoteRichApi.EntityBuilder.buildPageFromPageProxy(pageProxy);
                });
            });
        };
        StaticMethods.getActiveOutlineAsync = function () {
            return OneNote.run(function (context) {
                var outlineProxy = context.application.getActiveOutlineOrNull();
                outlineProxy.load(OneNoteRichApi.EntityProperties.getActiveOutlineProperties());
                context.trackedObjects.add(outlineProxy);
                return context.sync()
                    .then(function () {
                    // Proxy objects can be null - abstract this complexity
                    if (outlineProxy.isNull) {
                        return null;
                    }
                    StaticMethods.trackedHostObjects.add(outlineProxy.id, outlineProxy);
                    return OneNoteRichApi.EntityBuilder.buildOutlineFromOutlineProxy(outlineProxy);
                });
            });
        };
        StaticMethods.navigateToPageAsync = function (page) {
            return OneNote.run(function (context) {
                var pageProxy = context.application.navigateToPageWithClientUrl(page.clientUrl);
                pageProxy.load(OneNoteRichApi.EntityProperties.getPageProperties());
                context.trackedObjects.add(pageProxy);
                return context.sync().then(function () {
                    StaticMethods.trackedHostObjects.add(pageProxy.id, pageProxy);
                    return OneNoteRichApi.EntityBuilder.buildPageFromPageProxy(pageProxy);
                });
            });
        };
        StaticMethods.navigateToPageWithClientUrlAsync = function (clientUrl) {
            return OneNote.run(function (context) {
                var pageProxy = context.application.navigateToPageWithClientUrl(clientUrl);
                pageProxy.load(OneNoteRichApi.EntityProperties.getPageProperties());
                context.trackedObjects.add(pageProxy);
                return context.sync().then(function () {
                    StaticMethods.trackedHostObjects.add(pageProxy.id, pageProxy);
                    return OneNoteRichApi.EntityBuilder.buildPageFromPageProxy(pageProxy);
                });
            });
        };
        StaticMethods.getNotebooksAsync = function () {
            return OneNote.run(function (context) {
                var notebookProxies = context.application.notebooks;
                notebookProxies.load(OneNoteRichApi.EntityProperties.getNotebookProperties());
                context.trackedObjects.add(notebookProxies);
                return context.sync()
                    .then(function () {
                    // Proxy objects can be null - abstract this complexity
                    if (notebookProxies.isNull) {
                        return null;
                    }
                    var notebooks = new Array();
                    for (var i = 0; i < notebookProxies.items.length; i++) {
                        var notebookProxy = notebookProxies.items[i];
                        notebooks.push(OneNoteRichApi.EntityBuilder.buildNotebookFromNotebookProxy(notebookProxy));
                        StaticMethods.trackedHostObjects.add(notebookProxy.id, notebookProxy);
                    }
                    return notebooks;
                });
            });
        };
        // END REGION: Application methods
        // REGION: Notebook methods
        StaticMethods.createSectionInNotebookAsync = function (notebook, sectionName) {
            return this.trackedHostObjects.getNotebookProxyFromNotebookAsync(notebook)
                .then(function (notebookProxyToUse) {
                var sectionProxy = notebookProxyToUse.addSection(sectionName);
                notebookProxyToUse.context.trackedObjects.add(sectionProxy);
                sectionProxy.load(OneNoteRichApi.EntityProperties.getSectionProperties());
                return notebookProxyToUse.context.sync()
                    .then(function () {
                    StaticMethods.trackedHostObjects.add(sectionProxy.id, sectionProxy);
                    return OneNoteRichApi.EntityBuilder.buildSectionFromSectionProxy(sectionProxy);
                });
            });
        };
        StaticMethods.getSectionsInNotebookAsync = function (notebook, includePages) {
            return this.trackedHostObjects.getNotebookProxyFromNotebookAsync(notebook).then(function (notebookProxy) {
                var sectionProxies = notebookProxy.sections;
                sectionProxies.load(OneNoteRichApi.EntityProperties.getSectionProperties());
                notebookProxy.context.trackedObjects.add(sectionProxies);
                return notebookProxy.context.sync()
                    .then(function () {
                    var sections = StaticMethods.getSectionsFromSectionProxies(notebookProxy.sections, notebook, null, includePages);
                    return sections;
                });
            });
        };
        StaticMethods.getSectionsFromSectionProxies = function (sectionProxies, notebook, sectionGroup, includePages) {
            var sections = [];
            for (var i = 0; i < sectionProxies.items.length; i++) {
                var sectionProxy = sectionProxies.items[i];
                StaticMethods.trackedHostObjects.add(sectionProxy.id, sectionProxy);
                var section = OneNoteRichApi.EntityBuilder.buildSectionFromSectionProxyWithNotebook(sectionProxy, notebook);
                // Only assign if present
                if (sectionGroup != null) {
                    section.parentSectionGroup = sectionGroup;
                }
                if (includePages) {
                    var pages = StaticMethods.getPagesFromPagesProxies(sectionProxy.pages, section);
                    section.pages = pages;
                }
                sections.push(section);
            }
            return sections;
        };
        ;
        StaticMethods.getSectionGroupsFromSectionGroupProxies = function (sectionGroupProxies, notebook, parentSectionGroup, includePages) {
            var sectionGroups = [];
            for (var i = 0; i < sectionGroupProxies.items.length; i++) {
                var sectionGroupProxy = sectionGroupProxies.items[i];
                StaticMethods.trackedHostObjects.add(sectionGroupProxy.id, sectionGroupProxy);
                var sectionGroup = OneNoteRichApi.EntityBuilder.buildSectionGroupFromSectionGroupProxyWithNotebook(sectionGroupProxy, notebook);
                sectionGroup.parentSectionGroup = parentSectionGroup; // might be null
                sectionGroup.sections = StaticMethods.getSectionsFromSectionProxies(sectionGroupProxy.sections, notebook, sectionGroup, includePages);
                sectionGroup.sectionGroups = StaticMethods.getSectionGroupsFromSectionGroupProxies(sectionGroupProxy.sectionGroups, notebook, sectionGroup, includePages);
                sectionGroups.push(sectionGroup);
            }
            return sectionGroups;
        };
        ;
        StaticMethods.getNotebookStructureAsync = function (notebook, includePages) {
            return this.trackedHostObjects.getNotebookProxyFromNotebookAsync(notebook).then(function (notebookProxy) {
                notebookProxy.load(OneNoteRichApi.EntityProperties.getNotebookStructureProperties(includePages));
                // TODO: How will tracked objects work?
                return notebookProxy.context.sync()
                    .then(function () {
                    notebook.sections = StaticMethods.getSectionsFromSectionProxies(notebookProxy.sections, notebook, null, includePages);
                    notebook.sectionGroups = StaticMethods.getSectionGroupsFromSectionGroupProxies(notebookProxy.sectionGroups, notebook, null, includePages);
                    return notebook;
                });
            });
        };
        // END REGION: Notebook methods
        // REGION: Section methods
        StaticMethods.createPageInSectionAsync = function (section, pageTitle) {
            return this.trackedHostObjects.getSectionProxyFromSectionAsync(section).then(function (sectionProxy) {
                var pageProxy = sectionProxy.addPage(pageTitle);
                pageProxy.load(OneNoteRichApi.EntityProperties.getPageProperties());
                sectionProxy.context.trackedObjects.add(pageProxy);
                return sectionProxy.context.sync()
                    .then(function () {
                    StaticMethods.trackedHostObjects.add(pageProxy.id, pageProxy);
                    return OneNoteRichApi.EntityBuilder.buildPageFromPageProxy(pageProxy);
                });
            });
        };
        StaticMethods.getPagesFromPagesProxies = function (pageProxies, parentSection) {
            var pages = [];
            for (var i = 0; i < pageProxies.items.length; i++) {
                var pageProxy = pageProxies.items[i];
                StaticMethods.trackedHostObjects.add(pageProxy.id, pageProxy);
                var page = OneNoteRichApi.EntityBuilder.buildPageFromPageProxyWithParentSection(pageProxy, parentSection);
                pages.push(page);
            }
            return pages;
        };
        StaticMethods.getPagesInSectionAsync = function (section) {
            var _this = this;
            return this.trackedHostObjects.getSectionProxyFromSectionAsync(section).then(function (sectionProxy) {
                var pagesProxy = sectionProxy.pages;
                pagesProxy.load(OneNoteRichApi.EntityProperties.getPageProperties());
                sectionProxy.context.trackedObjects.add(pagesProxy);
                return sectionProxy.context.sync().then(function () {
                    var pages = _this.getPagesFromPagesProxies(pagesProxy, section);
                    return pages;
                });
            });
        };
        StaticMethods.insertSectionAsSiblingAsync = function (section, name) {
            return this.trackedHostObjects.getSectionProxyFromSectionAsync(section).then(function (sectionProxy) {
                var siblingSectionProxy = sectionProxy.insertSectionAsSibling(location, name);
                siblingSectionProxy.load(OneNoteRichApi.EntityProperties.getSectionProperties());
                sectionProxy.context.trackedObjects.add(siblingSectionProxy);
                return sectionProxy.context.sync()
                    .then(function () {
                    StaticMethods.trackedHostObjects.add(siblingSectionProxy.id, siblingSectionProxy);
                    return OneNoteRichApi.EntityBuilder.buildSectionFromSectionProxy(siblingSectionProxy);
                });
            });
        };
        // END REGION: Section methods
        // REGION: Page methods
        StaticMethods.getPageContentsAsync = function (page) {
            return this.trackedHostObjects.getPageProxyFromPageAsync(page).then(function (pageProxy) {
                var pageContentsProxy = pageProxy.contents;
                pageContentsProxy.load(OneNoteRichApi.EntityProperties.getPageContentPropertiesOnly());
                pageProxy.context.trackedObjects.add(pageContentsProxy);
                return pageProxy.context.sync().then(function () {
                    var pageContentProxiesWithImages = [];
                    var pageContentProxiesWithOutlines = [];
                    var otherPageContentProxies = [];
                    var i;
                    for (i = 0; i < pageContentsProxy.items.length; i++) {
                        var pageContentProxy = pageContentsProxy.items[i];
                        StaticMethods.trackedHostObjects.add(pageContentProxy.id, pageContentProxy);
                        if (pageContentProxy.type === OneNoteRichApi.Constants.pageContentImageType) {
                            pageContentProxiesWithImages.push(pageContentProxy);
                            var imageProxy = pageContentProxy.image;
                            pageProxy.context.trackedObjects.add(imageProxy);
                            pageContentProxy.image.load(OneNoteRichApi.EntityProperties.getImagePropertiesOnly());
                        }
                        else if (pageContentProxy.type === OneNoteRichApi.Constants.pageContentOutlineType) {
                            pageContentProxiesWithOutlines.push(pageContentProxy);
                            var outlineProxy = pageContentProxy.outline;
                            pageProxy.context.trackedObjects.add(outlineProxy);
                            pageContentProxy.outline.load(OneNoteRichApi.EntityProperties.getOutlinePropertiesOnly());
                        }
                        else {
                            otherPageContentProxies.push(pageContentProxy);
                        }
                    }
                    return pageProxy.context.sync().then(function () {
                        var pageContents = [];
                        for (i = 0; i < pageContentsProxy.items.length; i++) {
                            var pageContentProxy = pageContentsProxy.items[i];
                            var pageContent = OneNoteRichApi.EntityBuilder.buildPageContentFromPageContentProxyWithPage(pageContentProxy, page);
                            pageContents.push(pageContent);
                            if (pageContentProxy.type === OneNoteRichApi.Constants.pageContentImageType) {
                                StaticMethods.trackedHostObjects.add(pageContentProxy.image.id, pageContentProxy.image);
                                var image = OneNoteRichApi.EntityBuilder.buildImageFromImageProxyWithParent(pageContentProxy.image, pageContent, null);
                                pageContent.image = image;
                            }
                            else if (pageContentProxy.type === OneNoteRichApi.Constants.pageContentOutlineType) {
                                StaticMethods.trackedHostObjects.add(pageContentProxy.outline.id, pageContentProxy.outline);
                                var outline = OneNoteRichApi.EntityBuilder.buildOutlineFromOutlineProxyWithParentPageContent(pageContentProxy.outline, pageContent);
                                pageContent.outline = outline;
                                pageContent.outline.paragraphs = [];
                                // Paragraphs
                                for (var p = 0; p < pageContentProxy.outline.paragraphs.items.length; p++) {
                                    var paragraphProxy = pageContentProxy.outline.paragraphs.items[p];
                                    StaticMethods.trackedHostObjects.add(paragraphProxy.id, paragraphProxy);
                                    var paragraph = OneNoteRichApi.EntityBuilder.buildParagraphFromParagraphProxyWithOutline(paragraphProxy, outline, null);
                                    pageContent.outline.paragraphs.push(paragraph);
                                }
                            }
                        }
                        page.contents = pageContents;
                        return pageContents;
                    });
                });
            });
        };
        StaticMethods.createOutlineInPageAsync = function (page, left, top, html) {
            return this.trackedHostObjects.getPageProxyFromPageAsync(page).then(function (pageProxy) {
                var outlineProxy = pageProxy.addOutline(left, top, html);
                outlineProxy.load(OneNoteRichApi.EntityProperties.getActiveOutlineProperties());
                outlineProxy.context.trackedObjects.add(outlineProxy);
                return pageProxy.context.sync()
                    .then(function () {
                    StaticMethods.trackedHostObjects.add(outlineProxy.id, outlineProxy);
                    var outline = OneNoteRichApi.EntityBuilder.buildOutlineFromOutlineProxy(outlineProxy);
                    return outline;
                });
            });
        };
        StaticMethods.insertPageAsSiblingAsync = function (page, location, title) {
            return this.trackedHostObjects.getPageProxyFromPageAsync(page).then(function (pageProxy) {
                var siblingPageProxy = pageProxy.insertPageAsSibling(location, title);
                siblingPageProxy.load(OneNoteRichApi.EntityProperties.getPageProperties());
                pageProxy.context.trackedObjects.add(siblingPageProxy);
                return pageProxy.context.sync()
                    .then(function () {
                    StaticMethods.trackedHostObjects.add(siblingPageProxy.id, siblingPageProxy);
                    return OneNoteRichApi.EntityBuilder.buildPageFromPageProxy(siblingPageProxy);
                });
            });
        };
        StaticMethods.updatePagePropertiesAsync = function (page) {
            return this.trackedHostObjects.getPageProxyFromPageAsync(page).then(function (pageProxy) {
                // List the properties that can be updates from a page
                pageProxy.Title = page.title;
                pageProxy.pageLevel = page.pageLevel;
                pageProxy.load(OneNoteRichApi.EntityProperties.getPageProperties());
                pageProxy.context.trackedObjects.add(pageProxy);
                return pageProxy.context.sync()
                    .then(function () {
                    StaticMethods.trackedHostObjects.add(pageProxy.id, pageProxy);
                    return OneNoteRichApi.EntityBuilder.buildPageFromPageProxy(pageProxy);
                });
            });
        };
        StaticMethods.getPageStructureAsync = function (page) {
            return page.getContentsAsync().then(function (contents) {
                page.contents = contents;
                // Need to populate these contents appropiately - they've been expanded up to outline paragraphs
                var paragraphs = [];
                for (var c = 0; c < page.contents.length; c++) {
                    var content = page.contents[c];
                    if (content.outline) {
                        for (var p = 0; p < content.outline.paragraphs.length; p++) {
                            var paragraph = content.outline.paragraphs[p];
                            paragraphs.push(paragraph);
                        }
                        content.outline.paragraphs = []; // clean it first
                    }
                }
                return StaticMethods.trackedHostObjects.getParagraphProxiesFromParagraphs(paragraphs).then(function (paragraphProxies) {
                    return StaticMethods.getParagraphsFromParagraphs(paragraphProxies, true).then(function (tableParagraphs) {
                        return page;
                    });
                });
            });
        };
        // END REGION: Page methods
        // REGION: Page content methods
        StaticMethods.deletePageContentAsync = function (pageContent) {
            return this.trackedHostObjects.getPageContentProxyFromPageContentAsync(pageContent).then(function (pageContentProxy) {
                pageContentProxy.delete();
                return pageContentProxy.context.sync().then(function () {
                    // Nothing to do here - just return a successful promise
                    return;
                });
            });
        };
        StaticMethods.selectPageContentAsync = function (pageContent) {
            return this.trackedHostObjects.getPageContentProxyFromPageContentAsync(pageContent).then(function (pageContentProxy) {
                pageContentProxy.select();
                return pageContentProxy.context.sync().then(function () {
                    // Nothing to do here - just return a successful promise
                    return;
                });
            });
        };
        // END REGION: Page content methods
        // REGION: Outline methods
        StaticMethods.appendImageToOutlineAsync = function (outline, base64EncodedImage, width, height) {
            return this.trackedHostObjects.getOutlineProxyFromOutlineAsync(outline).then(function (outlineProxy) {
                var imageProxy = outlineProxy.appendImage(base64EncodedImage, width, height);
                imageProxy.pageContent.load(OneNoteRichApi.EntityProperties.getPageContentPropertiesWithImage());
                outlineProxy.context.trackedObjects.add(imageProxy);
                return outlineProxy.context.sync().then(function () {
                    StaticMethods.trackedHostObjects.add(imageProxy.id, imageProxy);
                    return OneNoteRichApi.EntityBuilder.buildImageFromImageProxyWithParent(imageProxy, null, null); // TODO: no parent is being assigned here!
                });
            });
        };
        StaticMethods.appendHtmlToOutlineAsync = function (outline, html) {
            return this.trackedHostObjects.getOutlineProxyFromOutlineAsync(outline).then(function (outlineProxy) {
                outlineProxy.appendHtml(html);
                return outlineProxy.context.sync().then(function () {
                    // Nothing to do here - just return a successful promise
                    return;
                });
            });
        };
        StaticMethods.appendRichTextToOutlineAsync = function (outline, paragraphText) {
            return this.trackedHostObjects.getOutlineProxyFromOutlineAsync(outline).then(function (outlineProxy) {
                var richTextProxy = outlineProxy.appendRichText(paragraphText);
                richTextProxy.load(OneNoteRichApi.EntityProperties.getRichTextProperties());
                outlineProxy.context.trackedObjects.add(richTextProxy);
                return outlineProxy.context.sync().then(function () {
                    StaticMethods.trackedHostObjects.add(richTextProxy.id, richTextProxy);
                    return OneNoteRichApi.EntityBuilder.buildRichTextFromRichTextProxy(richTextProxy);
                });
            });
        };
        StaticMethods.appendTableToOutlineAsync = function (outline, rowCount, columnCount, values) {
            return this.trackedHostObjects.getOutlineProxyFromOutlineAsync(outline).then(function (outlineProxy) {
                var tableProxy = outlineProxy.appendTable(rowCount, columnCount, values);
                tableProxy.load(OneNoteRichApi.EntityProperties.getTableProperties());
                outlineProxy.context.trackedObjects.add(tableProxy);
                return outlineProxy.context.sync().then(function () {
                    StaticMethods.trackedHostObjects.add(tableProxy.id, tableProxy);
                    return OneNoteRichApi.EntityBuilder.buildTableFromTableProxy(tableProxy);
                });
            });
        };
        StaticMethods.selectOutlineAsync = function (outline) {
            return this.trackedHostObjects.getOutlineProxyFromOutlineAsync(outline).then(function (outlineProxy) {
                outlineProxy.select();
                return outlineProxy.context.sync().then(function () {
                    // Nothing to do here - just return a successful promise
                    return;
                });
            });
        };
        StaticMethods.getParagraphsFromOutline = function (outlines, recursive) {
            return this.trackedHostObjects.getoutlineProxiesFromOutlinesAsync(outlines).then(function (outlineProxies) {
                if (outlineProxies.length === 0) {
                    return Promise.resolve([]);
                }
                var o;
                var i;
                for (o = 0; o < outlineProxies.length; o++) {
                    var outlineProxy = outlineProxies[o];
                    var paragraphsProxy = outlineProxy.paragraphs;
                    paragraphsProxy.load(OneNoteRichApi.EntityProperties.getParagraphPropertiesOnly());
                    outlineProxy.context.trackedObjects.add(paragraphsProxy);
                }
                var context = outlineProxies[0].context;
                return context.sync().then(function () {
                    var allParagraphProxies = [];
                    for (o = 0; o < outlineProxies.length; o++) {
                        var outlineProxy = outlineProxies[o];
                        var paragraphsProxy = outlineProxy.paragraphs;
                        for (i = 0; i < paragraphsProxy.items.length; i++) {
                            paragraphsProxy.items[i].outlineObject = outlineProxy.outlineObject;
                            paragraphsProxy.items[i].outlineObject.paragraphs = [];
                            allParagraphProxies.push(paragraphsProxy.items[i]);
                        }
                    }
                    return StaticMethods.getParagraphsFromParagraphs(allParagraphProxies, recursive);
                });
            });
        };
        // END REGION: Outline methods
        // REGION: Image methods
        StaticMethods.getBase64ImageFromImage = function (image) {
            return this.trackedHostObjects.getImageProxyFromImageAsync(image).then(function (imageProxy) {
                var base64Image = imageProxy.getBase64Image();
                return imageProxy.context.sync().then(function () {
                    return base64Image.value;
                });
            });
        };
        // END REGION: Image methods
        // REGION: Paragraph methods
        StaticMethods.deleteParagraphAsync = function (paragraph) {
            return this.trackedHostObjects.getParagraphProxyFromParagraph(paragraph).then(function (paragraphProxy) {
                paragraphProxy.delete();
                return paragraphProxy.context.sync().then(function () {
                    // Nothing to do here - just return a successful promise
                    return;
                });
            });
        };
        StaticMethods.selectParagraphAsync = function (paragraph) {
            return this.trackedHostObjects.getParagraphProxyFromParagraph(paragraph).then(function (paragraphProxy) {
                paragraphProxy.select();
                return paragraphProxy.context.sync().then(function () {
                    // Nothing to do here - just return a successful promise
                    return;
                });
            });
        };
        StaticMethods.insertHtmlInParagraphAsSiblingAsync = function (paragraph, insertLocation, html) {
            return this.trackedHostObjects.getParagraphProxyFromParagraph(paragraph).then(function (paragraphProxy) {
                paragraphProxy.insertHtml(location, html);
                return paragraphProxy.context.sync().then(function () {
                    // Nothing to do here - just return a successful promise
                    return;
                });
            });
        };
        StaticMethods.insertImageInParagraphAsSiblingAsync = function (paragraph, insertLocation, base64EncodedImage, width, height) {
            return this.trackedHostObjects.getParagraphProxyFromParagraph(paragraph).then(function (paragraphProxy) {
                var imageProxy = paragraphProxy.appendImage(insertLocation, base64EncodedImage);
                imageProxy.load(OneNoteRichApi.EntityProperties.getImageProperties());
                paragraphProxy.context.trackedObjects.add(imageProxy);
                return paragraphProxy.context.sync().then(function () {
                    StaticMethods.trackedHostObjects.add(imageProxy.id, imageProxy);
                    return OneNoteRichApi.EntityBuilder.buildImageFromImageProxyWithParent(imageProxy, null, null); // TODO: no image parent?
                });
            });
        };
        StaticMethods.insertRichTextInParagraphAsSiblingAsync = function (paragraph, insertLocation, paragraphText) {
            return this.trackedHostObjects.getParagraphProxyFromParagraph(paragraph).then(function (paragraphProxy) {
                var richTextProxy = paragraphProxy.insertRichText(insertLocation, paragraphText);
                richTextProxy.load(OneNoteRichApi.EntityProperties.getRichTextProperties());
                paragraphProxy.context.trackedObjects.add(richTextProxy);
                return paragraphProxy.context.sync().then(function () {
                    StaticMethods.trackedHostObjects.add(richTextProxy.id, richTextProxy);
                    return OneNoteRichApi.EntityBuilder.buildRichTextFromRichTextProxy(richTextProxy);
                });
            });
        };
        StaticMethods.insertTableInParagraphAsSiblingAsync = function (paragraph, insertLocation, rowCount, columnCount, values) {
            return this.trackedHostObjects.getParagraphProxyFromParagraph(paragraph).then(function (paragraphProxy) {
                var tableProxy = paragraphProxy.insertTable(insertLocation, rowCount, columnCount, values);
                tableProxy.load(OneNoteRichApi.EntityProperties.getTableProperties());
                paragraphProxy.context.trackedObjects.add(tableProxy);
                return paragraphProxy.context.sync().then(function () {
                    StaticMethods.trackedHostObjects.add(tableProxy.id, tableProxy);
                    return OneNoteRichApi.EntityBuilder.buildTableFromTableProxy(tableProxy);
                });
            });
        };
        // END REGION: Paragraph methods
        // REGION: Table methods
        StaticMethods.appendColumnToTableAsync = function (table, values) {
            return this.trackedHostObjects.getTableProxyFromTable(table).then(function (tableProxy) {
                tableProxy.appendColumn(values);
                return tableProxy.context.sync().then(function () {
                    // Nothing to do here - just return a successful promise
                    return;
                });
            });
        };
        StaticMethods.appendRowToTableAsync = function (table, values) {
            return this.trackedHostObjects.getTableProxyFromTable(table).then(function (tableProxy) {
                var tableRowProxy = tableProxy.appendRow(values);
                tableRowProxy.load(OneNoteRichApi.EntityProperties.getTableRowPropertiesOnly());
                tableProxy.context.trackedObjects.add(tableRowProxy);
                return tableProxy.context.sync().then(function () {
                    StaticMethods.trackedHostObjects.add(tableRowProxy.id, tableRowProxy);
                    return OneNoteRichApi.EntityBuilder.buildTableRowFromTableRowProxy(tableRowProxy);
                });
            });
        };
        StaticMethods.deleteColumnsInTableAsync = function (table, columnIndex, columnCount) {
            return this.trackedHostObjects.getTableProxyFromTable(table).then(function (tableProxy) {
                tableProxy.deleteColumns(columnIndex, columnCount);
                return tableProxy.context.sync().then(function () {
                    // Nothing to do here - just return a successful promise
                    return;
                });
            });
        };
        // END REGION: Table methods
        // REGION: TableRow methods
        StaticMethods.deleteTableRowAsync = function (tableRow) {
            return this.trackedHostObjects.getTableRowProxyFromTableRow(tableRow).then(function (tableRowProxy) {
                tableRowProxy.delete();
                return tableRowProxy.context.sync().then(function () {
                    // Nothing to do here - just return a successful promise
                    return;
                });
            });
        };
        StaticMethods.insertRowAsTableRowSiblingAsync = function (tableRow, insertLocation, values) {
            return this.trackedHostObjects.getTableRowProxyFromTableRow(tableRow).then(function (tableRowProxy) {
                var siblingProxy = tableRowProxy.insertRowAsSinling(insertLocation, values);
                siblingProxy.load(OneNoteRichApi.EntityProperties.getTableRowProperties());
                tableRowProxy.context.trackedObjects.add(siblingProxy);
                return tableRowProxy.context.sync().then(function () {
                    StaticMethods.trackedHostObjects.add(siblingProxy.id, siblingProxy);
                    return OneNoteRichApi.EntityBuilder.buildTableRowFromTableRowProxy(siblingProxy);
                });
            });
        };
        // END REGION: TableRow methods
        // REGION: TableCell methods
        StaticMethods.appendHtmlToTableCell = function (tableCell, html) {
            return this.trackedHostObjects.getTableCellProxyFromTableCell(tableCell).then(function (tableCellProxy) {
                tableCellProxy.appendHtml(html);
                return tableCellProxy.context.sync().then(function () {
                    // Nothing to do here - just return a successful promise
                    return;
                });
            });
        };
        // END REGION: TableRow methods
        StaticMethods.buildParagraphsFromParagraphsProxiesWithChildren = function (allParagraphProxies) {
            var paragraphs = [];
            var i;
            for (i = 0; i < allParagraphProxies.length; i++) {
                var paragraphProxy = allParagraphProxies[i];
                var paragraph = OneNoteRichApi.EntityBuilder.buildParagraphFromParagraphProxyWithOutline(paragraphProxy, paragraphProxy.outlineObject, paragraphProxy.tableCellObject);
                paragraphs.push(paragraph);
                if (paragraphProxy.outlineObject) {
                    var parentOutline = paragraphProxy.outlineObject;
                    parentOutline.paragraphs.push(paragraph);
                }
                else if (paragraphProxy.tableCellObject) {
                    var parentTableCell = paragraphProxy.tableCellObject;
                    parentTableCell.paragraphs.push(paragraph);
                }
                else {
                    throw new Error("Not implemented");
                }
                if (paragraphProxy.type === OneNoteRichApi.Constants.paragraphImageType) {
                    StaticMethods.trackedHostObjects.add(paragraphProxy.image.id, paragraphProxy.image);
                    var image = OneNoteRichApi.EntityBuilder.buildImageFromImageProxyWithParent(paragraphProxy.image, null, paragraph);
                    paragraph.image = image;
                }
                else if (paragraphProxy.type === OneNoteRichApi.Constants.paragraphRichTextType) {
                    StaticMethods.trackedHostObjects.add(paragraphProxy.richText.id, paragraphProxy.richText);
                    var richText = OneNoteRichApi.EntityBuilder.buildRichTextFromRichTextProxyWithParagraph(paragraphProxy.richText, paragraph);
                    paragraph.richText = richText;
                }
                else if (paragraphProxy.type === OneNoteRichApi.Constants.paragraphTableType) {
                    StaticMethods.trackedHostObjects.add(paragraphProxy.table.id, paragraphProxy.table);
                    var table = OneNoteRichApi.EntityBuilder.buildTableFromTableProxyWithParagraph(paragraphProxy.table, paragraph);
                    table.tableRows = [];
                    // TableRows
                    for (var k = 0; k < paragraphProxy.table.rows.items.length; k++) {
                        var tableRowProxy = paragraphProxy.table.rows.items[k];
                        StaticMethods.trackedHostObjects.add(tableRowProxy.id, tableRowProxy);
                        var tableRow = OneNoteRichApi.EntityBuilder.buildTableRowFromTableRowProxyWithTable(tableRowProxy, table);
                        table.tableRows.push(tableRow);
                        tableRow.cells = [];
                        // Table cells
                        for (var f = 0; f < tableRowProxy.cells.items.length; f++) {
                            var tableCellProxy = tableRowProxy.cells.items[f];
                            StaticMethods.trackedHostObjects.add(tableCellProxy.id, tableCellProxy);
                            var tableCell = OneNoteRichApi.EntityBuilder.buildTableCellFromTableCellProxyWithTableRow(tableCellProxy, tableRow);
                            tableRow.cells.push(tableCell);
                            tableCell.paragraphs = [];
                            // Paragraphs
                            for (var p = 0; p < tableCellProxy.paragraphs.items.length; p++) {
                                var childParagraphProxy = tableCellProxy.paragraphs.items[p];
                                StaticMethods.trackedHostObjects.add(childParagraphProxy.id, childParagraphProxy);
                                var childParagraph = OneNoteRichApi.EntityBuilder.buildParagraphFromParagraphProxyWithOutline(childParagraphProxy, null, tableCell);
                                tableCell.paragraphs.push(childParagraph);
                            }
                        }
                    }
                    paragraph.table = table;
                }
            }
            return paragraphs;
        };
        StaticMethods.getParagraphsFromParagraphs = function (allParagraphProxies, recursive) {
            // Given a list of paragraphs, it expands them up to their table cells (if any)
            if (allParagraphProxies.length === 0) {
                return Promise.resolve([]);
            }
            var paragraphsWithImage = [];
            var paragraphsWithRichText = [];
            var paragraphsWithTable = [];
            var otherParagraphProxies = [];
            for (var i = 0; i < allParagraphProxies.length; i++) {
                var paragraphProxy = allParagraphProxies[i];
                StaticMethods.trackedHostObjects.add(paragraphProxy.id, paragraphProxy);
                if (paragraphProxy.type === OneNoteRichApi.Constants.paragraphImageType) {
                    paragraphsWithImage.push(paragraphProxy);
                    var imageProxy = paragraphProxy.image;
                    paragraphProxy.context.trackedObjects.add(imageProxy);
                    imageProxy.load(OneNoteRichApi.EntityProperties.getImagePropertiesOnly());
                }
                else if (paragraphProxy.type === OneNoteRichApi.Constants.paragraphRichTextType) {
                    paragraphsWithRichText.push(paragraphProxy);
                    var richTextProxy = paragraphProxy.richText;
                    paragraphProxy.context.trackedObjects.add(richTextProxy);
                    richTextProxy.load(OneNoteRichApi.EntityProperties.getRichTextPropertiesOnly());
                }
                else if (paragraphProxy.type === OneNoteRichApi.Constants.paragraphTableType) {
                    paragraphsWithTable.push(paragraphProxy);
                    var tableProxy = paragraphProxy.table;
                    paragraphProxy.context.trackedObjects.add(tableProxy);
                    tableProxy.load(OneNoteRichApi.EntityProperties.getTablePropertiesOnly());
                }
                else {
                    otherParagraphProxies.push(paragraphProxy);
                }
            }
            return allParagraphProxies[0].context.sync().then(function () {
                // At this point, we're expanded until paragraph proxies below tables we can build those paragraphs
                var paragraphs = StaticMethods.buildParagraphsFromParagraphsProxiesWithChildren(allParagraphProxies);
                if (!recursive) {
                    // Short circuit if we're not recursive
                    return Promise.resolve(paragraphs);
                }
                var allParagraphsToExpandAgain = [];
                for (var p = 0; p < paragraphs.length; p++) {
                    var paragraph = paragraphs[p];
                    if (paragraph.table) {
                        for (var tr = 0; tr < paragraph.table.tableRows.length; tr++) {
                            var tableRow = paragraph.table.tableRows[tr];
                            for (var tc = 0; tc < tableRow.cells.length; tc++) {
                                var tableCell = tableRow.cells[tc];
                                for (var cp = 0; cp < tableCell.paragraphs.length; cp++) {
                                    var childParagraph = tableCell.paragraphs[cp];
                                    allParagraphsToExpandAgain.push(childParagraph);
                                }
                                tableCell.paragraphs = [];
                            }
                        }
                    }
                }
                return StaticMethods.trackedHostObjects.getParagraphProxiesFromParagraphs(allParagraphsToExpandAgain).then(function (paragraphProxies) {
                    return StaticMethods.getParagraphsFromParagraphs(paragraphProxies, recursive).then(function (tableParagraphs) {
                        return paragraphs;
                    });
                });
            });
        };
        // Keeps tracked of host objects in their contexts
        StaticMethods.trackedHostObjects = new OneNoteRichApi.TrackedHostObjects();
        return StaticMethods;
    }());
    OneNoteRichApi.StaticMethods = StaticMethods;
})(OneNoteRichApi || (OneNoteRichApi = {}));
// The purpose of these objects is to Build "normal" JS objects from the proxy objects
// Theese won't be invalidated after a OneNote.Run - therefore can be:
// 1. Used as a global variable
// 2. Serialized / Deserialized and stored
// 3. Referenced multiple times
// Additionally, they have pointers to the OneNoteRichApi methods for convenience
/// <reference path="StaticMethods.ts"/>
/// <reference path="../typings/globals/es6-promise/index.d.ts"/>
var OneNoteRichApi;
(function (OneNoteRichApi) {
    // Provides methods at the top level singleton application object
    var Application = (function () {
        function Application() {
        }
        // Get the active notebook
        Application.getActiveNotebookAsync = function () {
            return OneNoteRichApi.StaticMethods.getActiveNotebookAsync();
        };
        // Get the active section
        Application.getActiveSectionAsync = function () {
            return OneNoteRichApi.StaticMethods.getActiveSectionAsync();
        };
        // Get the active page
        Application.getActivePageAsync = function () {
            return OneNoteRichApi.StaticMethods.getActivePageAsync();
        };
        // Get the active outline
        Application.getActiveOutlineAsync = function () {
            return OneNoteRichApi.StaticMethods.getActiveOutlineAsync();
        };
        // Get the list of notebooks
        Application.getNotebooksAsync = function () {
            return OneNoteRichApi.StaticMethods.getNotebooksAsync();
        };
		Application.navigateToPageAsync = function (page) {
            return OneNoteRichApi.StaticMethods.navigateToPageAsync(page);
        };
		Application.navigateToPageWithClientUrlAsync = function (clientUrl) {
            return OneNoteRichApi.StaticMethods.navigateToPageWithClientUrlAsync(clientUrl);
        };
        return Application;
    }());
    OneNoteRichApi.Application = Application;
    var Notebook = (function () {
        function Notebook(id, name, clientUrl) {
            this.id = id;
            this.name = name;
            this.clientUrl = clientUrl;
        }
        Notebook.prototype.createSectionAsync = function (sectionName) {
            return OneNoteRichApi.StaticMethods.createSectionInNotebookAsync(this, sectionName);
        };
        ;
        Notebook.prototype.getSectionsAsync = function () {
            return OneNoteRichApi.StaticMethods.getSectionsInNotebookAsync(this, false);
        };
        ;
        Notebook.prototype.getStructureAsync = function (includePages) {
            if (includePages === void 0) { includePages = false; }
            return OneNoteRichApi.StaticMethods.getNotebookStructureAsync(this, includePages);
        };
        return Notebook;
    }());
    OneNoteRichApi.Notebook = Notebook;
    var SectionGroup = (function () {
        function SectionGroup(id, name, clientUrl, parentNotebook) {
            this.id = id;
            this.name = name;
            this.parentNotebook = parentNotebook;
        }
        return SectionGroup;
    }());
    OneNoteRichApi.SectionGroup = SectionGroup;
    var Section = (function () {
        function Section(id, name, clientUrl, parentNotebook) {
            this.id = id;
            this.name = name;
            this.clientUrl = clientUrl;
            this.parentNotebook = parentNotebook;
        }
        Section.prototype.createPageAsync = function (pageTitle) {
            return OneNoteRichApi.StaticMethods.createPageInSectionAsync(this, pageTitle);
        };
        Section.prototype.getPagesAsync = function () {
            return OneNoteRichApi.StaticMethods.getPagesInSectionAsync(this);
        };
        return Section;
    }());
    OneNoteRichApi.Section = Section;
    var Page = (function () {
        function Page(id, title, pageLevel, clientUrl, parentSection) {
            this.id = id;
            this.title = title;
            this.pageLevel = pageLevel;
            this.clientUrl = clientUrl;
            this.parentSection = parentSection;
        }
        Page.prototype.createOutlineAsync = function (left, top, html) {
            return OneNoteRichApi.StaticMethods.createOutlineInPageAsync(this, left, top, html);
        };
        Page.prototype.insertPageAsSiblingAsync = function (location, title) {
            return OneNoteRichApi.StaticMethods.insertPageAsSiblingAsync(this, location, title);
        };
        Page.prototype.getContentsAsync = function () {
            return OneNoteRichApi.StaticMethods.getPageContentsAsync(this);
        };
        Page.prototype.updatePropertiesAsync = function () {
            return OneNoteRichApi.StaticMethods.updatePagePropertiesAsync(this);
        };
        Page.prototype.getStructureAsync = function () {
            return OneNoteRichApi.StaticMethods.getPageStructureAsync(this);
        };
        Page.prototype.navigateAsync = function () {
            return OneNoteRichApi.StaticMethods.navigateToPageAsync(this);
        };
        return Page;
    }());
    OneNoteRichApi.Page = Page;
    var PageContent = (function () {
        function PageContent(id, left, top, type, parentPage) {
            this.id = id;
            this.left = left;
            this.top = top;
            this.type = type;
            this.parentPage = parentPage;
        }
        PageContent.prototype.deleteAsync = function () {
            return OneNoteRichApi.StaticMethods.deletePageContentAsync(this);
        };
        PageContent.prototype.selectAsync = function () {
            return OneNoteRichApi.StaticMethods.selectPageContentAsync(this);
        };
        return PageContent;
    }());
    OneNoteRichApi.PageContent = PageContent;
    var Outline = (function () {
        function Outline(id, pageContent) {
            this.id = id;
            this.parentPageContent = pageContent;
        }
        Outline.prototype.appendHtmlAsync = function (html) {
            return OneNoteRichApi.StaticMethods.appendHtmlToOutlineAsync(this, html);
        };
        Outline.prototype.appendImageAsync = function (base64EncodedImage, width, height) {
            return OneNoteRichApi.StaticMethods.appendImageToOutlineAsync(this, base64EncodedImage, width, height);
        };
        Outline.prototype.appendRichTextAsync = function (paragraphText) {
            return OneNoteRichApi.StaticMethods.appendRichTextToOutlineAsync(this, paragraphText);
        };
        Outline.prototype.appendTableAsync = function (rowCount, columnCount, values) {
            return OneNoteRichApi.StaticMethods.appendTableToOutlineAsync(this, rowCount, columnCount, values);
        };
        Outline.prototype.selectAsync = function () {
            return OneNoteRichApi.StaticMethods.selectOutlineAsync(this);
        };
        Outline.prototype.getParagraphsAsync = function () {
            return OneNoteRichApi.StaticMethods.getParagraphsFromOutline([this], false);
        };
        return Outline;
    }());
    OneNoteRichApi.Outline = Outline;
    var Image = (function () {
        function Image(id, height, width, description, hyperlink) {
            this.id = id;
            this.height = height;
            this.width = width;
            this.description = description;
            this.hyperlink = hyperlink;
        }
        Image.prototype.getBase64ImageAsync = function () {
            return OneNoteRichApi.StaticMethods.getBase64ImageFromImage(this);
        };
        return Image;
    }());
    OneNoteRichApi.Image = Image;
    var Paragraph = (function () {
        function Paragraph(id, type) {
            this.id = id;
            this.type = type;
        }
        Paragraph.prototype.deleteAsync = function () {
            return OneNoteRichApi.StaticMethods.deleteParagraphAsync(this);
        };
        Paragraph.prototype.selectAsync = function () {
            return OneNoteRichApi.StaticMethods.selectParagraphAsync(this);
        };
        Paragraph.prototype.insertHtmlAsSiblingAsync = function (insertLocation, html) {
            return OneNoteRichApi.StaticMethods.insertHtmlInParagraphAsSiblingAsync(this, insertLocation, html);
        };
        Paragraph.prototype.insertImageAsSiblingAsync = function (insertLocation, base64EncodedImage, width, height) {
            return OneNoteRichApi.StaticMethods.insertImageInParagraphAsSiblingAsync(this, insertLocation, base64EncodedImage, width, height);
        };
        Paragraph.prototype.insertRichTextAsSibling = function (insertLocation, paragraphText) {
            return OneNoteRichApi.StaticMethods.insertRichTextInParagraphAsSiblingAsync(this, insertLocation, paragraphText);
        };
        Paragraph.prototype.insertTableAsSiblingAsync = function (insertLocation, rowCount, columnCount, values) {
            return OneNoteRichApi.StaticMethods.insertTableInParagraphAsSiblingAsync(this, insertLocation, rowCount, columnCount, values);
        };
        return Paragraph;
    }());
    OneNoteRichApi.Paragraph = Paragraph;
    var RichText = (function () {
        function RichText(id, text, parentParagraph) {
            this.id = id;
            this.text = text;
            this.parentParagraph = parentParagraph;
        }
        return RichText;
    }());
    OneNoteRichApi.RichText = RichText;
    var Table = (function () {
        function Table(id, columnCount, rowCount, parentParagraph) {
            this.id = id;
            this.columnCount = columnCount;
            this.rowCount = rowCount;
            this.parentParagraph = parentParagraph;
        }
        Table.prototype.appendColumnAsync = function (values) {
            return OneNoteRichApi.StaticMethods.appendColumnToTableAsync(this, values);
        };
        Table.prototype.appendRowAsync = function (values) {
            return OneNoteRichApi.StaticMethods.appendRowToTableAsync(this, values);
        };
        Table.prototype.deleteColumnsAsync = function (columnIndex, columnCount) {
            return OneNoteRichApi.StaticMethods.deleteColumnsInTableAsync(this, columnIndex, columnCount);
        };
        return Table;
    }());
    OneNoteRichApi.Table = Table;
    var TableRow = (function () {
        function TableRow(id, cellCount, rowIndex, parentTable) {
            this.id = id;
            this.cellCount = cellCount;
            this.rowIndex = rowIndex;
            this.parentTable = parentTable;
        }
        TableRow.prototype.deleteAsync = function () {
            return OneNoteRichApi.StaticMethods.deleteTableRowAsync(this);
        };
        TableRow.prototype.insertRowAsSiblingAsync = function (insertLocation, values) {
            return OneNoteRichApi.StaticMethods.insertRowAsTableRowSiblingAsync(this, insertLocation, values);
        };
        return TableRow;
    }());
    OneNoteRichApi.TableRow = TableRow;
    var TableCell = (function () {
        function TableCell(id, rowIndex, cellIndex, parentTableRow) {
            this.id = id;
            this.rowIndex = rowIndex;
            this.cellIndex = cellIndex;
            this.parentTableRow = parentTableRow;
        }
        TableCell.prototype.appendHtml = function (html) {
            return OneNoteRichApi.StaticMethods.appendHtmlToTableCell(this, html);
        };
        return TableCell;
    }());
    OneNoteRichApi.TableCell = TableCell;
})(OneNoteRichApi || (OneNoteRichApi = {}));
/// <reference path="EntityProperties.ts"/>
/// <reference path="Entities.ts"/>
var OneNoteRichApi;
(function (OneNoteRichApi) {
    var Constants = (function () {
        function Constants() {
        }
        Constants.pageContentImageType = "Image";
        Constants.pageContentOutlineType = "Outline";
        Constants.paragraphImageType = "Image";
        Constants.paragraphTableType = "Table";
        Constants.paragraphRichTextType = "RichText";
        return Constants;
    }());
    OneNoteRichApi.Constants = Constants;
})(OneNoteRichApi || (OneNoteRichApi = {}));