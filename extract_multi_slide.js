const puppeteer = require('puppeteer');
const fs = require('fs').promises;

async function extractSlideData(htmlFilePath, outputPath) {
    const browser = await puppeteer.launch({
        headless: true,
        devtools: false,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-web-security']
    });
    const page = await browser.newPage();

    try {
        const htmlContent = await fs.readFile(htmlFilePath, 'utf-8');
        await page.setViewport({ width: 1920, height: 1080 });
        await page.setContent(htmlContent, { waitUntil: 'networkidle0' });

        // Wait for all resources (images, fonts, etc.) to load
        await page.waitForFunction(() => {
            const images = Array.from(document.querySelectorAll('img'));
            const fonts = document.fonts ? document.fonts.ready : Promise.resolve();
            return Promise.all([fonts, images.every(img => img.complete)]);
        }, { timeout: 15000 }).catch(() => console.log('Some resources may not have loaded'));

        const documentInfo = await page.evaluate(() => {
            const body = document.body;
            const html = document.documentElement;

            const actualWidth = Math.max(
                body.scrollWidth, body.offsetWidth,
                html.clientWidth, html.scrollWidth, html.offsetWidth
            );
            const actualHeight = Math.max(
                body.scrollHeight, body.offsetHeight,
                html.clientHeight, html.scrollHeight, html.offsetHeight
            );

            const viewportWidth = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
            const viewportHeight = Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0);

            const slideElements = Array.from(document.querySelectorAll('.slide'));
            const slidesInfo = slideElements.map((slide, index) => {
                const rect = slide.getBoundingClientRect();
                return {
                    index: index + 1,
                    rect: {
                        x: rect.left,
                        y: rect.top,
                        width: rect.width,
                        height: rect.height
                    }
                };
            });

            return {
                actualWidth,
                actualHeight,
                viewportWidth,
                viewportHeight,
                slidesCount: slideElements.length,
                slidesInfo
            };
        });

        console.log('Document dimensions:', documentInfo);
        console.log(`Found ${documentInfo.slidesCount} slides`);

        const targetWidth = Math.max(documentInfo.actualWidth, 1920);
        const targetHeight = Math.max(documentInfo.actualHeight, 1080);
        await page.setViewport({ width: targetWidth, height: targetHeight });

        const allSlidesData = await page.evaluate(async (docInfo) => {
            const slides = [];
            const slideElements = Array.from(document.querySelectorAll('.slide')) || [document.body];

            const IMPORTANT_ELEMENTS = [
                'div', 'span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
                'strong', 'b', 'em', 'i', 'u', 'strike', 'del', 'ins', 'mark', 'small', 'sub', 'sup',
                'ul', 'ol', 'li', 'dl', 'dt', 'dd',
                'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'caption', 'colgroup', 'col',
                'img', 'svg', 'video', 'audio', 'iframe',
                'a', 'button', 'input', 'textarea', 'select', 'option', 'label', 'fieldset', 'legend',
                'blockquote', 'pre', 'code', 'kbd', 'samp', 'var',
                'article', 'section', 'aside', 'nav', 'header', 'footer', 'main',
                'figure', 'figcaption', 'details', 'summary', 'dialog',
                'hr', 'br', 'wbr',
                'abbr', 'address', 'bdi', 'bdo', 'cite', 'dfn', 'q', 'ruby', 'rt', 'rp', 's', 'time'
            ];

            function getElementId(element) {
                const rect = element.getBoundingClientRect();
                const text = element.textContent ? element.textContent.trim().substring(0, 50) : '';
                const src = element.src || element.href || '';
                const tagName = element.tagName.toLowerCase();
                const className = element.className || '';
                const id = element.id || '';
                return `${tagName}-${id}-${className}-${rect.left}-${rect.top}-${rect.width}-${rect.height}-${text}-${src}`;
            }

            function getAllDescendants(element) {
                let descendants = [];
                function traverse(el) {
                    for (let child of el.children) {
                        descendants.push(child);
                        traverse(child);
                    }
                }
                traverse(element);
                return descendants;
            }

            function shouldProcessElement(element, slideContainer) {
                if (element === slideContainer) return false;
                if (!slideContainer.contains(element)) return false;

                const tagName = element.tagName.toLowerCase();
                if (!IMPORTANT_ELEMENTS.includes(tagName)) return false;

                const styles = window.getComputedStyle(element);
                if (styles.display === 'none' || styles.visibility === 'hidden' || styles.opacity === '0') return false;

                // Always process styled inline elements and list/table elements
                if (['strong', 'b', 'i', 'em', 'u', 'mark', 'span', 'li', 'td', 'th'].includes(tagName)) {
                    return true;
                }

                // Skip empty elements except for those that might have visual significance
                const text = element.textContent ? element.textContent.trim() : '';
                if (!text && !['img', 'svg', 'video', 'audio', 'iframe'].includes(tagName)) {
                    const hasStyling = styles.borderWidth !== '0px' || 
                                    styles.backgroundColor !== 'rgba(0, 0, 0, 0)' || 
                                    styles.boxShadow !== 'none';
                    if (!hasStyling) return false;
                }

                return true;
            }

            function getInheritedStyles(element) {
                const inheritedProps = ['color', 'fontFamily', 'fontSize', 'fontWeight', 'lineHeight', 'textAlign'];
                const inherited = {};
                
                let currentElement = element.parentElement;
                while (currentElement && currentElement.tagName !== 'BODY') {
                    const styles = window.getComputedStyle(currentElement);
                    inheritedProps.forEach(prop => {
                        if (!inherited[prop] && styles[prop] !== 'initial') {
                            inherited[prop] = styles[prop];
                        }
                    });
                    currentElement = currentElement.parentElement;
                }
                
                return inherited;
            }

            function getResolvedBackgroundColor(element) {
                let current = element;
                while (current) {
                    const style = window.getComputedStyle(current);
                    const bg = style.backgroundColor;
                    if (bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent') {
                        return bg;
                    }
                    current = current.parentElement;
                }
                return 'rgba(0, 0, 0, 0)';
            }

            function extractComprehensiveStyles(element) {
                const styles = window.getComputedStyle(element);
                const customProperties = {};
                for (const prop of styles) {
                    if (prop.startsWith('--')) customProperties[prop] = styles.getPropertyValue(prop);
                }

                return {
                    fontSize: styles.fontSize,
                    fontFamily: styles.fontFamily,
                    fontWeight: styles.fontWeight,
                    fontStyle: styles.fontStyle,
                    lineHeight: styles.lineHeight,
                    textAlign: styles.textAlign,
                    textDecoration: styles.textDecoration,
                    color: styles.color,
                    backgroundColor: getResolvedBackgroundColor(element),
                    width: styles.width,
                    height: styles.height,
                    padding: styles.padding,
                    paddingTop: styles.paddingTop,
                    paddingRight: styles.paddingRight,
                    paddingBottom: styles.paddingBottom,
                    paddingLeft: styles.paddingLeft,
                    margin: styles.margin,
                    marginTop: styles.marginTop,
                    marginRight: styles.marginRight,
                    marginBottom: styles.marginBottom,
                    marginLeft: styles.marginLeft,
                    border: styles.border,
                    borderWidth: styles.borderWidth,
                    borderStyle: styles.borderStyle,
                    borderColor: styles.borderColor,
                    borderTopWidth: styles.borderTopWidth,
                    borderRightWidth: styles.borderRightWidth,
                    borderBottomWidth: styles.borderBottomWidth,
                    borderLeftWidth: styles.borderLeftWidth,
                    borderTopStyle: styles.borderTopStyle,
                    borderRightStyle: styles.borderRightStyle,
                    borderBottomStyle: styles.borderBottomStyle,
                    borderLeftStyle: styles.borderLeftStyle,
                    borderTopColor: styles.borderTopColor || styles.borderColor,
                    borderRightColor: styles.borderRightColor || styles.borderColor,
                    borderBottomColor: styles.borderBottomColor || styles.borderColor,
                    borderLeftColor: styles.borderLeftColor || styles.borderColor,
                    borderRadius: styles.borderRadius,
                    position: styles.position,
                    display: styles.display,
                    visibility: styles.visibility,
                    zIndex: styles.zIndex,
                    boxShadow: styles.boxShadow,
                    listStyleType: styles.listStyleType,
                    listStylePosition: styles.listStylePosition,
                    listStyleImage: styles.listStyleImage
                };
            }

            function getTextContent(element, processedElements) {
                const tagName = element.tagName.toLowerCase();
                const styles = window.getComputedStyle(element);
                
                // For inline elements, get all text content
                if (styles.display === 'inline' || styles.display === 'inline-block') {
                    return element.textContent.trim();
                }
                
                // For block elements, collect only direct text nodes
                const directText = Array.from(element.childNodes)
                    .filter(node => node.nodeType === Node.TEXT_NODE)
                    .map(node => node.textContent.trim())
                    .join(' ')
                    .trim();

                return directText;
            }

            function getInlineGroup(container, slideContainer) {
                const tagName = container.tagName.toLowerCase();
                
                if (!['p', 'div', 'li', 'td', 'th'].includes(tagName)) return null;

                const hasInlineFormatting = Array.from(container.children).some(child => {
                    const childTag = child.tagName.toLowerCase();
                    return ['strong', 'b', 'em', 'i', 'u', 'mark', 'span'].includes(childTag);
                });

                if (!hasInlineFormatting) return null;

                const rect = container.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                
                const inlineElements = [];
                let fullText = '';
                
                const walkNodes = (node) => {
                    if (node.nodeType === Node.TEXT_NODE) {
                        let text = node.textContent.replace(/\s+/g, ' ');
                        if (text.trim() === '') return; // Skip pure whitespace
                        inlineElements.push({
                            type: 'text',
                            text: text,
                            styles: extractComprehensiveStyles(container)
                        });
                        fullText += text;
                    } else if (node.nodeType === Node.ELEMENT_NODE) {
                        const childTag = node.tagName.toLowerCase();
                        if (['strong', 'b', 'em', 'i', 'u', 'mark', 'span'].includes(childTag)) {
                            const text = node.textContent;
                            inlineElements.push({
                                type: childTag,
                                text: text,
                                styles: extractComprehensiveStyles(node)
                            });
                            fullText += text;
                        } else if (childTag === 'br') {
                            inlineElements.push({
                                type: 'br',
                                text: '\n',
                                styles: extractComprehensiveStyles(container)
                            });
                            fullText += '\n';
                        } else {
                            for (const child of node.childNodes) {
                                walkNodes(child);
                            }
                        }
                    }
                };

                for (const node of container.childNodes) {
                    walkNodes(node);
                }

                const filteredInline = inlineElements.filter(el => el.text.trim() !== '' || el.type === 'br');

                if (filteredInline.length > 0) {
                    // Trim leading spaces from the first text element
                    if (filteredInline[0].type === 'text') {
                        filteredInline[0].text = filteredInline[0].text.replace(/^\s+/, '');
                    }
                    // Trim trailing spaces from the last text element
                    const lastIndex = filteredInline.length - 1;
                    if (filteredInline[lastIndex].type === 'text') {
                        filteredInline[lastIndex].text = filteredInline[lastIndex].text.replace(/\s+$/, '');
                    }
                    // Rebuild fullText after trimming
                    fullText = filteredInline.map(el => el.text).join('');

                    return {
                        text: fullText,
                        inlineElements: filteredInline,
                        groupRect: {
                            x: Math.round(rect.left - slideRect.left),
                            y: Math.round(rect.top - slideRect.top),
                            width: Math.round(rect.width),
                            height: Math.round(rect.height)
                        },
                        styles: extractComprehensiveStyles(container)
                    };
                }

                return null;
            }

            function getListInfo(element, slideContainer) {
                const tagName = element.tagName.toLowerCase();
                const listInfo = {};

                if (['ul', 'ol'].includes(tagName)) {
                    const items = Array.from(element.querySelectorAll(':scope > li'));
                    const styles = window.getComputedStyle(element);
                    const rect = element.getBoundingClientRect();
                    const slideRect = slideContainer.getBoundingClientRect();
                    
                    listInfo.type = tagName;
                    listInfo.itemCount = items.length;
                    listInfo.rect = {
                        x: Math.round(rect.left - slideRect.left),
                        y: Math.round(rect.top - slideRect.top),
                        width: Math.round(rect.width),
                        height: Math.round(rect.height)
                    };
                    listInfo.listStyles = {
                        listStyleType: styles.listStyleType,
                        listStylePosition: styles.listStylePosition,
                        paddingLeft: styles.paddingLeft,
                        marginTop: styles.marginTop,
                        marginBottom: styles.marginBottom
                    };
                    
                    listInfo.items = items.map((item, index) => {
                        const itemRect = item.getBoundingClientRect();
                        const itemStyles = extractComprehensiveStyles(item);
                        const text = item.textContent.trim();
                        
                        // Check for inline elements within list items
                        const inlineGroup = getInlineGroup(item, slideContainer);
                        
                        // Check for nested list
                        const nestedListElement = item.querySelector(':scope > ul, :scope > ol');
                        const nestedList = nestedListElement ? getListInfo(nestedListElement, slideContainer) : null;
                        
                        return {
                            index,
                            text: text,
                            styles: itemStyles,
                            rect: {
                                x: Math.round(itemRect.left - slideRect.left),
                                y: Math.round(itemRect.top - slideRect.top),
                                width: Math.round(itemRect.width),
                                height: Math.round(itemRect.height)
                            },
                            inlineGroup: inlineGroup,
                            nestedList: nestedList,
                            hasNestedList: !!nestedListElement
                        };
                    });
                    
                    if (tagName === 'ol') {
                        listInfo.start = element.start || 1;
                        listInfo.reversed = element.reversed || false;
                    }
                }

                return listInfo;
            }

            function getTableInfo(element, slideContainer) {
                const tagName = element.tagName.toLowerCase();
                const tableInfo = {};

                if (tagName === 'table') {
                    const rows = Array.from(element.querySelectorAll('tr'));
                    const rect = element.getBoundingClientRect();
                    const slideRect = slideContainer.getBoundingClientRect();

                    tableInfo.type = 'table';
                    tableInfo.rect = {
                        x: Math.round(rect.left - slideRect.left),
                        y: Math.round(rect.top - slideRect.top),
                        width: Math.round(rect.width),
                        height: Math.round(rect.height)
                    };
                    tableInfo.rowCount = rows.length;
                    tableInfo.styles = extractComprehensiveStyles(element);
                    
                    // Calculate column count based on the widest row
                    let maxCols = 0;
                    rows.forEach(row => {
                        const cells = Array.from(row.querySelectorAll('td, th'));
                        let colCount = 0;
                        cells.forEach(cell => colCount += cell.colSpan || 1);
                        maxCols = Math.max(maxCols, colCount);
                    });
                    tableInfo.columnCount = maxCols;
                    
                    tableInfo.rows = rows.map((row, rowIndex) => {
                        const cells = Array.from(row.querySelectorAll('td, th'));
                        const rowRect = row.getBoundingClientRect();
                        const rowStyles = extractComprehensiveStyles(row);
                        return {
                            index: rowIndex,
                            rect: {
                                x: Math.round(rowRect.left - slideRect.left),
                                y: Math.round(rowRect.top - slideRect.top),
                                width: Math.round(rowRect.width),
                                height: Math.round(rowRect.height)
                            },
                            styles: rowStyles,
                            cells: cells.map((cell, cellIndex) => {
                                const cellRect = cell.getBoundingClientRect();
                                const inlineGroup = getInlineGroup(cell, slideContainer);
                                const text = inlineGroup ? '' : getTextContent(cell, new Set());
                                return {
                                    type: cell.tagName.toLowerCase(),
                                    text: text,
                                    rect: {
                                        x: Math.round(cellRect.left - slideRect.left),
                                        y: Math.round(cellRect.top - slideRect.top),
                                        width: Math.round(cellRect.width),
                                        height: Math.round(cellRect.height)
                                    },
                                    styles: extractComprehensiveStyles(cell),
                                    colSpan: cell.colSpan || 1,
                                    rowSpan: cell.rowSpan || 1,
                                    cellIndex,
                                    inlineGroup
                                };
                            })
                        };
                    });
                }

                return tableInfo;
            }

            function getAccuratePosition(element, slideContainer) {
                const rect = element.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                const styles = window.getComputedStyle(element);
                
                // For inline elements, use Range API for more accurate positioning
                if (styles.display === 'inline' || styles.display === 'inline-block') {
                    const range = document.createRange();
                    range.selectNodeContents(element);
                    const rangeRect = range.getBoundingClientRect();
                    
                    if (rangeRect.width > 0 && rangeRect.height > 0) {
                        return {
                            x: Math.max(0, Math.round(rangeRect.left - slideRect.left)),
                            y: Math.max(0, Math.round(rangeRect.top - slideRect.top)),
                            width: Math.round(rangeRect.width),
                            height: Math.round(rangeRect.height)
                        };
                    }
                }
                
                return {
                    x: Math.max(0, Math.round(rect.left - slideRect.left)),
                    y: Math.max(0, Math.round(rect.top - slideRect.top)),
                    width: Math.round(rect.width),
                    height: Math.round(rect.height)
                };
            }

            for (let slideIndex = 0; slideIndex < slideElements.length; slideIndex++) {
                const slideElement = slideElements[slideIndex];
                const slideRect = slideElement.getBoundingClientRect();

                const slide = {
                    slideId: slideIndex + 1,
                    elements: [],
                    slideWidth: Math.round(slideRect.width),
                    slideHeight: Math.round(slideRect.height),
                    slidePosition: {
                        x: Math.round(slideRect.left),
                        y: Math.round(slideRect.top)
                    },
                    slideStyles: extractComprehensiveStyles(slideElement)
                };

                const processedElements = new Set();
                const allElements = Array.from(slideElement.querySelectorAll('*'));
                const elementsToProcess = allElements.filter(element => shouldProcessElement(element, slideElement));

                // Sort elements by position (top to bottom, left to right)
                elementsToProcess.sort((a, b) => {
                    const rectA = a.getBoundingClientRect();
                    const rectB = b.getBoundingClientRect();
                    const topDiff = rectA.top - rectB.top;
                    return Math.abs(topDiff) < 5 ? rectA.left - rectB.left : topDiff;
                });

                // Create a comprehensive set to track all elements that should not be processed individually
                const skipElements = new Set();
                
                // First pass: identify all elements that will be part of inline groups or should be skipped
                elementsToProcess.forEach(element => {
                    const inlineGroup = getInlineGroup(element, slideElement);
                    if (inlineGroup) {
                        // Mark all its inline children as handled
                        getAllDescendants(element).forEach(child => {
                            const childTag = child.tagName.toLowerCase();
                            if (['strong', 'b', 'em', 'i', 'u', 'mark', 'span', 'br'].includes(childTag)) {
                                skipElements.add(child);
                            }
                        });
                    }
                });

                elementsToProcess.forEach(element => {
                    const elementId = getElementId(element);
                    if (processedElements.has(elementId)) return;
                    
                    const tagName = element.tagName.toLowerCase();
                    const position = getAccuratePosition(element, slideElement);
                    const styles = extractComprehensiveStyles(element);
                    const inlineGroup = getInlineGroup(element, slideElement);

                    // Skip elements that are handled by inline groups
                    if (skipElements.has(element)) {
                        processedElements.add(elementId);
                        return;
                    }

                    processedElements.add(elementId);

                    const elementData = {
                        type: tagName,
                        x: position.x,
                        y: position.y,
                        width: position.width,
                        height: position.height,
                        styles,
                        className: element.className || '',
                        id: element.id || '',
                        zIndex: parseInt(styles.zIndex) || 0,
                        inlineGroup: inlineGroup
                    };
                    
                    // Only add text if not part of an inline group and not already handled
                    if (!inlineGroup && ['div', 'span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(tagName)) {
                        const text = getTextContent(element, processedElements);
                        if (text) {
                            elementData.text = text;
                        }
                    }

                    // Handle lists - but skip if we're processing a list item that has an inline group parent
                    if (['ul', 'ol'].includes(tagName)) {
                        elementData.listInfo = getListInfo(element, slideElement);
                    }

                    // Handle tables
                    if (['table'].includes(tagName)) {
                        elementData.tableInfo = getTableInfo(element, slideElement);
                    }

                    // Handle images
                    if (tagName === 'img') {
                        elementData.mediaInfo = {
                            src: element.src || '',
                            alt: element.alt || '',
                            naturalWidth: element.naturalWidth || 0,
                            naturalHeight: element.naturalHeight || 0
                        };
                    }

                    slide.elements.push(elementData);

                    // Skip processing descendants for container elements like lists and tables
                    if (['ul', 'ol', 'table'].includes(tagName)) {
                        getAllDescendants(element).forEach(desc => {
                            processedElements.add(getElementId(desc));
                        });
                    }
                });

                // Sort elements by z-index for proper layering
                slide.elements.sort((a, b) => a.zIndex - b.zIndex || 0);
                slides.push(slide);
            }

            return slides;
        }, documentInfo);

        await fs.writeFile(outputPath, JSON.stringify(allSlidesData, null, 2), 'utf-8');
        console.log(`Successfully extracted ${allSlidesData.length} slides to ${outputPath}`);

    } catch (err) {
        console.error('Error processing slides:', err);
    } finally {
        await browser.close();
    }
}

const htmlFilePath = 'presentation.html';
const outputPath = 'slides_data.json';

extractSlideData(htmlFilePath, outputPath).catch(err => {
    console.error('Error:', err);
});