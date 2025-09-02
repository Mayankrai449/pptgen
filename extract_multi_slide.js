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
            let slideElements = Array.from(document.querySelectorAll('.slide'));
            if (slideElements.length === 0) {
                slideElements = Array.from(document.body.children).filter(el => el.tagName === 'DIV' || el.tagName === 'SECTION');
            }
            if (slideElements.length === 0) {
                slideElements = [document.body];
            }
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
            let slideElements = Array.from(document.querySelectorAll('.slide'));
            if (slideElements.length === 0) {
                slideElements = Array.from(document.body.children).filter(el => el.tagName === 'DIV' || el.tagName === 'SECTION');
            }
            if (slideElements.length === 0) {
                slideElements = [document.body];
            }
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
                'abbr', 'address', 'bdi', 'bdo', 'cite', 'dfn', 'q', 'ruby', 'rt', 'rp', 's', 'time',
                'canvas'
            ];

            function safeFloat(value) {
                try {
                    return parseFloat(value.replace(/[^0-9.-]/g, '')) || 0;
                } catch {
                    return 0;
                }
            }

            function getElementId(element) {
                if (!element || !element.tagName) return 'unknown';
                
                const rect = element.getBoundingClientRect();
                const text = element.textContent ? element.textContent.trim().substring(0, 50) : '';
                const src = element.src || element.href || '';
                const tagName = element.tagName.toLowerCase();
                const className = element.className || '';
                const id = element.id || '';
                
                return `${tagName}-${id}-${className}-${rect.left.toFixed(1)}-${rect.top.toFixed(1)}-${rect.width.toFixed(1)}-${rect.height.toFixed(1)}-${text}-${src}`;
            }

            function getResolvedBackgroundColor(elementOrStyle, isPseudo = false) {
                let bg;
                if (isPseudo) {
                    bg = elementOrStyle.backgroundColor;
                } else {
                    const element = elementOrStyle;
                    const style = window.getComputedStyle(element);
                    bg = style.backgroundColor;
                    const hasExplicitBackground = element.style.backgroundColor ||
                        element.getAttribute('style')?.includes('background') ||
                        Array.from(element.classList).some(cls => {
                            try {
                                const rules = Array.from(document.styleSheets).flatMap(sheet =>
                                    Array.from(sheet.cssRules || []));
                                return rules.some(rule =>
                                    rule.selectorText?.includes(`.${cls}`) &&
                                    rule.style?.backgroundColor
                                );
                            } catch {
                                return false;
                            }
                        });
                    if (bg === 'rgb(255, 255, 255)' && !hasExplicitBackground) {
                        let current = element.parentElement;
                        let foundExplicitParentBg = false;
                        while (current && current !== document.body) {
                            const parentStyle = window.getComputedStyle(current);
                            const parentBg = parentStyle.backgroundColor;
                            if (parentBg !== 'rgba(0, 0, 0, 0)' && parentBg !== 'transparent') {
                                foundExplicitParentBg = true;
                                break;
                            }
                            current = current.parentElement;
                        }
                        if (!foundExplicitParentBg) {
                            return 'rgba(255, 255, 255, 0)';
                        }
                    }
                }
                return bg || 'rgba(0, 0, 0, 0)';
            }

            function getTextContent(element, processedElements) {
                const tagName = element.tagName.toLowerCase();
                const styles = window.getComputedStyle(element);
                if (styles.display === 'inline' || styles.display === 'inline-block') {
                    return element.textContent.trim();
                }
                const directText = Array.from(element.childNodes)
                    .filter(node => node.nodeType === Node.TEXT_NODE)
                    .map(node => node.textContent.trim())
                    .join(' ')
                    .trim();
                return directText;
            }

            function hasTextContentChildren(element) {
                const textContentTags = ['div', 'span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'td', 'th', 'ul', 'ol', 'table'];
                return Array.from(element.children).some(child => {
                    const childTag = child.tagName.toLowerCase();
                    return textContentTags.includes(childTag) && child.textContent.trim().length > 0;
                });
            }

            function shouldExtractTextContent(element, processedTextElements) {
                const elementId = getElementId(element);
                if (processedTextElements.has(elementId)) {
                    return false;
                }
                
                if (hasTextContentChildren(element)) {
                    return false;
                }
                
                const tagName = element.tagName.toLowerCase();
                if (['ul', 'ol', 'table', 'tbody', 'thead', 'tfoot'].includes(tagName)) {
                    return false;
                }
                
                return true;
            }

            function shouldExtractInlineGroup(element, processedTextElements) {
                const elementId = getElementId(element);
                if (processedTextElements.has(elementId)) {
                    return false;
                }

                const tagName = element.tagName.toLowerCase();
                if (!['p', 'div', 'li', 'td', 'th'].includes(tagName)) return false;
                
                if (element.children.length > 10) {
                    return false;
                }
                
                const highlightWrapper = element.querySelector('[class*="highlight-wrapper"], [class*="wrapper"]');
                if (highlightWrapper) {
                    const wrapperStyles = window.getComputedStyle(highlightWrapper);
                    const hasDashedBorder = wrapperStyles.borderStyle.includes('dashed') || 
                                          wrapperStyles.borderTopStyle === 'dashed' ||
                                          wrapperStyles.borderBottomStyle === 'dashed' ||
                                          wrapperStyles.borderLeftStyle === 'dashed' ||
                                          wrapperStyles.borderRightStyle === 'dashed';
                    
                    if (hasDashedBorder || wrapperStyles.borderColor.includes('255, 0, 0') || wrapperStyles.borderColor.includes('red')) {
                        const structuredChildCount = Array.from(element.children).filter(child => {
                            const childTag = child.tagName.toLowerCase();
                            return ['div', 'ul', 'ol', 'table', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(childTag);
                        }).length;
                        return structuredChildCount === 0;
                    }
                }
                
                const hasInlineFormatting = Array.from(element.children).some(child => {
                    const childTag = child.tagName.toLowerCase();
                    return ['strong', 'b', 'em', 'i', 'u', 'mark', 'span'].includes(childTag);
                });
                
                if (!hasInlineFormatting) return false;
                
                const structuredChildren = Array.from(element.children).filter(child => {
                    const childTag = child.tagName.toLowerCase();
                    return ['div', 'ul', 'ol', 'table', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(childTag);
                });
               
                if (structuredChildren.length > 0) {
                    return false;
                }
                
                const nonInlineChildren = Array.from(element.children).filter(child => {
                    const childTag = child.tagName.toLowerCase();
                    return !['strong', 'b', 'em', 'i', 'u', 'mark', 'span', 'br'].includes(childTag);
                });
                
                if (nonInlineChildren.length > 0) {
                    return false;
                }
                
                return true;
            }

            function getInlineGroup(container, slideContainer, processedTextElements) {
                if (!shouldExtractInlineGroup(container, processedTextElements)) {
                    return null;
                }
                
                const highlightWrapper = container.querySelector('[class*="highlight-wrapper"], [class*="wrapper"]');
                let hasRedHighlight = false;
                let redHighlightStyles = {};
                
                if (highlightWrapper) {
                    const wrapperStyles = window.getComputedStyle(highlightWrapper);
                    const hasDashedBorder = wrapperStyles.borderStyle.includes('dashed') || 
                                          wrapperStyles.borderTopStyle === 'dashed' ||
                                          wrapperStyles.borderBottomStyle === 'dashed' ||
                                          wrapperStyles.borderLeftStyle === 'dashed' ||
                                          wrapperStyles.borderRightStyle === 'dashed';
                    
                    if (hasDashedBorder || wrapperStyles.borderColor.includes('255, 0, 0') || wrapperStyles.borderColor.includes('red')) {
                        hasRedHighlight = true;
                        const wrapperRect = highlightWrapper.getBoundingClientRect();
                        const containerRect = container.getBoundingClientRect();
                        const slideRect = slideContainer.getBoundingClientRect();
                        
                        redHighlightStyles = {
                            borderColor: wrapperStyles.borderColor,
                            borderWidth: wrapperStyles.borderWidth,
                            borderStyle: wrapperStyles.borderStyle,
                            backgroundColor: wrapperStyles.backgroundColor,
                            position: wrapperStyles.position,
                            top: wrapperStyles.top,
                            left: wrapperStyles.left,
                            right: wrapperStyles.right,
                            bottom: wrapperStyles.bottom,
                            zIndex: wrapperStyles.zIndex,
                            extendedRect: {
                                x: Math.round(wrapperRect.left - slideRect.left),
                                y: Math.round(wrapperRect.top - slideRect.top),
                                width: Math.round(wrapperRect.width),
                                height: Math.round(wrapperRect.height)
                            }
                        };
                    }
                }
                
                const hasInlineFormatting = Array.from(container.children).some(child => {
                    const childTag = child.tagName.toLowerCase();
                    return ['strong', 'b', 'em', 'i', 'u', 'mark', 'span'].includes(childTag);
                });
                if (!hasInlineFormatting && !hasRedHighlight) return null;
                
                const rect = container.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                const inlineElements = [];
                let fullText = '';
                const walkNodes = (node) => {
                    if (node.nodeType === Node.TEXT_NODE) {
                        let text = node.textContent.replace(/\s+/g, ' ');
                        if (text.trim() === '') return;
                        inlineElements.push({
                            type: 'text',
                            text: text,
                            styles: extractComprehensiveStyles(container)
                        });
                        fullText += text;
                    } else if (node.nodeType === Node.ELEMENT_NODE) {
                        const childTag = node.tagName.toLowerCase();
                        if (['strong', 'b', 'em', 'i', 'u', 'mark', 'span'].includes(childTag)) {
                            let text = node.textContent;
                            const lines = text.split('\n');
                            const normalizedLines = [lines[0], ...lines.slice(1).map(l => l.replace(/^[ \t]+/, ''))];
                            text = normalizedLines.join('\n');
                            if (text.trim() === '') return;
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
                if (filteredInline.length > 0 || hasRedHighlight) {
                    if (filteredInline[0] && filteredInline[0].type === 'text') {
                        filteredInline[0].text = filteredInline[0].text.replace(/^\s+/, '');
                    }
                    const lastIndex = filteredInline.length - 1;
                    if (filteredInline[lastIndex] && filteredInline[lastIndex].type === 'text') {
                        filteredInline[lastIndex].text = filteredInline[lastIndex].text.replace(/\s+$/, '');
                    }
                    fullText = filteredInline.map(el => el.text).join('');
                    const result = {
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
                    
                    if (hasRedHighlight) {
                        result.hasRedHighlight = true;
                        result.redHighlightStyles = redHighlightStyles;
                    }
                    
                    return result;
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
                        const inlineGroup = getInlineGroup(item, slideContainer, new Set());
                        const nestedListElement = item.querySelector(':scope > ul, :scope > ol');
                        const nestedList = nestedListElement ? getListInfo(nestedListElement, slideContainer) : null;
                        const beforePseudo = extractPseudo(item, slideContainer, '::before');
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
                            hasNestedList: !!nestedListElement,
                            bulletInfo: beforePseudo ? {
                                content: beforePseudo.text,
                                position: beforePseudo,
                                styles: beforePseudo.styles
                            } : null
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
                        x: Math.round((rect.left - slideRect.left) * 10) / 10,
                        y: Math.round((rect.top - slideRect.top) * 10) / 10,
                        width: Math.round(rect.width * 10) / 10,
                        height: Math.round(rect.height * 10) / 10
                    };
                    tableInfo.rowCount = rows.length;
                    tableInfo.styles = extractComprehensiveStyles(element);
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
                        let actualRowHeight = Math.round(rowRect.height * 10) / 10;
                        const inlineStyle = row.getAttribute('style');
                        if (inlineStyle && inlineStyle.includes('height:')) {
                            const heightMatch = inlineStyle.match(/height:\s*(\d+(?:\.\d+)?)px/);
                            if (heightMatch) {
                                const specifiedHeight = parseFloat(heightMatch[1]);
                                actualRowHeight = Math.min(specifiedHeight, actualRowHeight);
                            }
                        }
                        actualRowHeight = Math.max(12, actualRowHeight);
                        return {
                            index: rowIndex,
                            rect: {
                                x: Math.round((rowRect.left - slideRect.left) * 10) / 10,
                                y: Math.round((rowRect.top - slideRect.top) * 10) / 10,
                                width: Math.round(rowRect.width * 10) / 10,
                                height: actualRowHeight
                            },
                            styles: rowStyles,
                            cells: cells.map((cell, cellIndex) => {
                                const cellRect = cell.getBoundingClientRect();
                                const inlineGroup = getInlineGroup(cell, slideContainer, new Set());
                                const text = inlineGroup ? '' : getTextContent(cell, new Set());
                                return {
                                    type: cell.tagName.toLowerCase(),
                                    text: text,
                                    rect: {
                                        x: Math.round((cellRect.left - slideRect.left) * 10) / 10,
                                        y: Math.round((cellRect.top - slideRect.top) * 10) / 10,
                                        width: Math.round(cellRect.width * 10) / 10,
                                        height: Math.round(cellRect.height * 10) / 10
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

            function getAccurateImageDimensions(img, container, slideContainer) {
                const imgRect = img.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                return {
                    x: Math.round(imgRect.left - slideRect.left),
                    y: Math.round(imgRect.top - slideRect.top),
                    width: Math.round(imgRect.width),
                    height: Math.round(imgRect.height),
                    naturalWidth: img.naturalWidth || 0,
                    naturalHeight: img.naturalHeight || 0
                };
            }

            function getAccuratePosition(element, slideContainer) {
                const rect = element.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                const styles = window.getComputedStyle(element);
                let x = rect.left - slideRect.left;
                let y = rect.top - slideRect.top;
                let width = rect.width;
                let height = rect.height;
                if (element.classList.contains('main-content-area') || styles.display === 'flex') {
                    const computedRect = element.getBoundingClientRect();
                    x = computedRect.left - slideRect.left;
                    y = computedRect.top - slideRect.top;
                    width = computedRect.width;
                    height = computedRect.height;
                }
                if (element.classList.contains('table') || element.tagName.toLowerCase() === 'table') {
                    const computedRect = element.getBoundingClientRect();
                    x = computedRect.left - slideRect.left;
                    y = computedRect.top - slideRect.top;
                    width = computedRect.width;
                    height = computedRect.height;
                }
                if (element.classList.contains('timeline')) {
                    const computedRect = element.getBoundingClientRect();
                    x = computedRect.left - slideRect.left;
                    y = computedRect.top - slideRect.top;
                    width = computedRect.width;
                    height = computedRect.height;
                }
                const devicePixelRatio = window.devicePixelRatio || 1;
                x /= devicePixelRatio;
                y /= devicePixelRatio;
                width /= devicePixelRatio;
                height /= devicePixelRatio;
                if (element.classList.contains('company') || element.classList.contains('companies')) {
                    const computedRect = element.getBoundingClientRect();
                    x = (computedRect.left - slideRect.left) / devicePixelRatio;
                    y = (computedRect.top - slideRect.top) / devicePixelRatio;
                    width = computedRect.width / devicePixelRatio;
                    height = computedRect.height / devicePixelRatio;
                }
                if (styles.display === 'inline' || styles.display === 'inline-block') {
                    try {
                        const range = document.createRange();
                        range.selectNodeContents(element);
                        const rangeRect = range.getBoundingClientRect();
                        if (rangeRect.width > 0 && rangeRect.height > 0) {
                            x = (rangeRect.left - slideRect.left) / devicePixelRatio;
                            y = (rangeRect.top - slideRect.top) / devicePixelRatio;
                            width = rangeRect.width / devicePixelRatio;
                            height = rangeRect.height / devicePixelRatio;
                        }
                    } catch (e) {}
                }
                const parent = element.parentElement;
                const parentStyles = parent ? window.getComputedStyle(parent) : null;
                if (parentStyles && (parentStyles.display === 'flex' || parentStyles.display === 'inline-flex')) {
                    const computedRect = element.getBoundingClientRect();
                    x = (computedRect.left - slideRect.left) / devicePixelRatio;
                    y = (computedRect.top - slideRect.top) / devicePixelRatio;
                    width = computedRect.width / devicePixelRatio;
                    height = computedRect.height / devicePixelRatio;
                }
                if (element.closest('.footer')) {
                    const computedRect = element.getBoundingClientRect();
                    x = (computedRect.left - slideRect.left) / devicePixelRatio;
                    y = (computedRect.top - slideRect.top) / devicePixelRatio;
                    width = computedRect.width / devicePixelRatio;
                    height = computedRect.height / devicePixelRatio;
                }
                const transform = styles.transform;
                if (transform && transform !== 'none') {
                    try {
                        const matrix = new DOMMatrix(transform);
                        if (matrix.a !== 1 || matrix.d !== 1) {
                            width *= Math.abs(matrix.a);
                            height *= Math.abs(matrix.d);
                        }
                        x += matrix.e / devicePixelRatio;
                        y += matrix.f / devicePixelRatio;
                    } catch (e) {}
                }
                return {
                    x: Math.round(x * 10) / 10,
                    y: Math.round(y * 10) / 10,
                    width: Math.max(1, Math.round(width * 10) / 10),
                    height: Math.max(1, Math.round(height * 10) / 10)
                };
            }

            function extractComprehensiveStyles(elementOrStyle) {
                let styles;
                let isPseudo = false;
                if (elementOrStyle instanceof Element) {
                    styles = window.getComputedStyle(elementOrStyle);
                } else {
                    styles = elementOrStyle;
                    isPseudo = true;
                }
                const customProperties = {};
                for (const prop of styles) {
                    if (prop.startsWith('--')) customProperties[prop] = styles.getPropertyValue(prop);
                }
                const borderProps = {};
                ['Top', 'Right', 'Bottom', 'Left'].forEach(side => {
                    borderProps[`border${side}Width`] = styles[`border${side}Width`];
                    borderProps[`border${side}Style`] = styles[`border${side}Style`];
                    borderProps[`border${side}Color`] = styles[`border${side}Color`];
                });
                return {
                    fontSize: styles.fontSize,
                    fontFamily: styles.fontFamily,
                    fontWeight: styles.fontWeight,
                    fontStyle: styles.fontStyle,
                    lineHeight: styles.lineHeight,
                    textAlign: styles.textAlign,
                    textDecoration: styles.textDecoration,
                    color: styles.color,
                    background: styles.background,
                    backgroundColor: getResolvedBackgroundColor(elementOrStyle, isPseudo),
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
                    ...borderProps,
                    borderRadius: styles.borderRadius,
                    position: styles.position,
                    display: styles.display,
                    visibility: styles.visibility,
                    zIndex: styles.zIndex,
                    boxShadow: styles.boxShadow,
                    listStyleType: styles.listStyleType,
                    listStylePosition: styles.listStylePosition,
                    listStyleImage: styles.listStyleImage,
                    overflow: styles.overflow,
                    overflowX: styles.overflowX,
                    overflowY: styles.overflowY,
                    flex: styles.flex,
                    flexDirection: styles.flexDirection,
                    justifyContent: styles.justifyContent,
                    alignItems: styles.alignItems,
                    gap: styles.gap,
                    left: styles.left,
                    top: styles.top,
                    right: styles.right,
                    bottom: styles.bottom,
                    content: styles.content,
                    objectFit: styles.objectFit,
                    objectPosition: styles.objectPosition,
                    maxWidth: styles.maxWidth,
                    maxHeight: styles.maxHeight,
                    minWidth: styles.minWidth,
                    minHeight: styles.minHeight,
                    transform: styles.transform,
                    transformOrigin: styles.transformOrigin,
                    clipPath: styles.clipPath,
                    flexShrink: styles.flexShrink,
                    pointerEvents: styles.pointerEvents,
                    customProperties
                };
            }

            function getPseudoPosition(element, slideContainer, pseudoStyles, pseudo) {
                const parentRect = element.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                let width = safeFloat(pseudoStyles.width);
                let height = safeFloat(pseudoStyles.height);
                const leftRaw = pseudoStyles.left;
                const rightRaw = pseudoStyles.right;
                const topRaw = pseudoStyles.top;
                const bottomRaw = pseudoStyles.bottom;
                const left = leftRaw === 'auto' ? null : safeFloat(leftRaw);
                const right = rightRaw === 'auto' ? null : safeFloat(rightRaw);
                const top = topRaw === 'auto' ? null : safeFloat(topRaw);
                const bottom = bottomRaw === 'auto' ? null : safeFloat(bottomRaw);
                if (width === 0 && left !== null && right !== null) {
                    width = parentRect.width - left - right;
                } else if (width === 0 && right !== null) {
                    const elementStyles = window.getComputedStyle(element);
                    const parentWidth = parentRect.width;
                    width = parentWidth - right;
                } else if (width === 0) {
                    if (pseudo === '::before' && element.tagName.toLowerCase() === 'li') {
                        width = safeFloat(pseudoStyles.fontSize) || 14;
                    } else {
                        width = parentRect.width;
                    }
                }
                if (height === 0 && top !== null && bottom !== null) {
                    height = parentRect.height - top - bottom;
                } else if (height === 0) {
                    if (pseudo === '::after') {
                        height = 2;
                    } else if (pseudo === '::before' && element.tagName.toLowerCase() === 'li') {
                        height = safeFloat(pseudoStyles.fontSize) || 14;
                    } else {
                        height = 2;
                    }
                }
                let x;
                if (left !== null) {
                    x = left;
                } else if (right !== null) {
                    x = parentRect.width - width - right;
                } else {
                    x = 0;
                }
                let y;
                if (top !== null) {
                    y = top;
                } else if (bottom !== null) {
                    y = parentRect.height - height - bottom;
                } else {
                    if (pseudo === '::after') {
                        y = parentRect.height - height;
                    } else if (pseudo === '::before') {
                        y = parentRect.height - height;
                    } else {
                        y = 0;
                    }
                }
                x += parentRect.left - slideRect.left;
                y += parentRect.top - slideRect.top;

                return {
                    x: Math.round(x * 10) / 10,
                    y: Math.round(y * 10) / 10,
                    width: Math.round(width * 10) / 10,
                    height: Math.round(height * 10) / 10
                };
            }

            function extractPseudo(element, slideContainer, pseudo) {
                const pseudoStyles = window.getComputedStyle(element, pseudo);
                if (pseudoStyles.display === 'none' || pseudoStyles.visibility === 'hidden' || pseudoStyles.opacity === '0') return null;
                const content = pseudoStyles.content.replace(/['"]/g, '').trim();
                const hasContent = content !== '' && content !== 'none';
                const hasBackground = pseudoStyles.backgroundColor !== 'rgba(0, 0, 0, 0)' && pseudoStyles.backgroundColor !== 'transparent';
                const hasGradient = pseudoStyles.background && pseudoStyles.background.includes('gradient');
                const hasBorder = ['borderTopWidth', 'borderRightWidth', 'borderBottomWidth', 'borderLeftWidth']
                    .some(prop => safeFloat(pseudoStyles[prop]) > 0);
                const hasBoxShadow = pseudoStyles.boxShadow !== 'none';
                if (!hasContent && !hasBackground && !hasGradient && !hasBorder && !hasBoxShadow) return null;
                const position = getPseudoPosition(element, slideContainer, pseudoStyles, pseudo);
                if (!position || (position.width <= 0 && position.height <= 0)) return null;
                const styles = extractComprehensiveStyles(pseudoStyles);
                
                const parentZ = parseInt(window.getComputedStyle(element).zIndex) || 0;
                let assignedZ = parentZ - 2;
                if (pseudo === '::after') assignedZ = parentZ - 1;
                const elementData = {
                    type: 'pseudo',
                    pseudoType: pseudo,
                    x: position.x,
                    y: position.y,
                    width: position.width,
                    height: position.height,
                    styles,
                    zIndex: assignedZ,
                    parentClassName: element.className || '',
                    parentTagName: element.tagName.toLowerCase()
                };
                if (hasContent) {
                    elementData.text = content;
                }
                return elementData;
            }

            function extractShapeInfo(element, slideContainer) {
                const styles = window.getComputedStyle(element);
                const isShapeByClipPath = styles.clipPath && styles.clipPath.startsWith('polygon');
                const isShapeByTransform = styles.transform !== 'none' && styles.transform.includes('matrix');

                if (!isShapeByClipPath && !isShapeByTransform) return null;

                const rect = element.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                
                let rotation = 0;
                const transform = styles.transform;
                if (transform && transform !== 'none') {
                    try {
                        const matrix = new DOMMatrix(transform);
                        rotation = Math.round(Math.atan2(matrix.b, matrix.a) * (180 / Math.PI));
                    } catch (e) {
                        const match = transform.match(/rotate\(([-\d.]+)deg\)/);
                        if (match) {
                            rotation = parseFloat(match[1]);
                        }
                    }
                }

                // Use offsetWidth/Height for pre-transform dimensions to avoid oversized shapes
                const width = element.offsetWidth;
                const height = element.offsetHeight;

                // Adjust x/y for rotated elements, since getBoundingClientRect gives the top-left of the bounding box
                const centerX = rect.left - slideRect.left + rect.width / 2;
                const centerY = rect.top - slideRect.top + rect.height / 2;
                const x = centerX - width / 2;
                const y = centerY - height / 2;

                return {
                    type: 'shape',
                    rect: {
                        x: Math.round(x),
                        y: Math.round(y),
                        width: Math.round(width),
                        height: Math.round(height)
                    },
                    styles: extractComprehensiveStyles(element),
                    className: element.className || '',
                    rotation: rotation,
                    clipPath: styles.clipPath,
                    text: element.textContent.trim()
                };
            }

            function extractOverlayInfo(element, slideContainer) {
                const styles = window.getComputedStyle(element);
                if (styles.borderStyle === 'none' && styles.backgroundColor === 'rgba(0, 0, 0, 0)' && !styles.boxShadow) return null;
                const rect = element.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                return {
                    type: 'overlay',
                    rect: {
                        x: Math.round((rect.left - slideRect.left) * 10) / 10,
                        y: Math.round((rect.top - slideRect.top) * 10) / 10,
                        width: Math.round(rect.width * 10) / 10,
                        height: Math.round(rect.height * 10) / 10
                    },
                    styles: extractComprehensiveStyles(element),
                    className: element.className || '',
                    zIndex: parseInt(styles.zIndex) || 0,
                    pointerEvents: styles.pointerEvents
                };
            }

            function extractCompanyElements(companyDiv, slideContainer) {
                const elements = [];
                const rect = companyDiv.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                const img = companyDiv.querySelector('img');
                const span = companyDiv.querySelector('span');
                if (img) {
                    const imgRect = img.getBoundingClientRect();
                    elements.push({
                        type: 'img',
                        x: Math.round(imgRect.left - slideRect.left),
                        y: Math.round(imgRect.top - slideRect.top),
                        width: Math.round(imgRect.width),
                        height: Math.round(imgRect.height),
                        styles: extractComprehensiveStyles(img),
                        className: img.className || '',
                        id: img.id || '',
                        zIndex: 0,
                        mediaInfo: {
                            src: img.src || '',
                            alt: img.alt || '',
                            naturalWidth: img.naturalWidth || 0,
                            naturalHeight: img.naturalHeight || 0
                        }
                    });
                }
                if (span) {
                    const spanRect = span.getBoundingClientRect();
                    elements.push({
                        type: 'span',
                        x: Math.round(spanRect.left - slideRect.left),
                        y: Math.round(spanRect.top - slideRect.top),
                        width: Math.round(spanRect.width),
                        height: Math.round(spanRect.height),
                        styles: extractComprehensiveStyles(span),
                        className: span.className || '',
                        id: span.id || '',
                        zIndex: 0,
                        text: span.textContent.trim()
                    });
                }
                return elements;
            }

            function extractFooterElements(footerDiv, slideContainer) {
                const elements = [];
                const slideRect = slideContainer.getBoundingClientRect();
                for (const child of footerDiv.children) {
                    const childTag = child.tagName.toLowerCase();
                    if (childTag === 'img') {
                        const imgRect = child.getBoundingClientRect();
                        elements.push({
                            type: 'img',
                            x: Math.round(imgRect.left - slideRect.left),
                            y: Math.round(imgRect.top - slideRect.top),
                            width: Math.round(imgRect.width),
                            height: Math.round(imgRect.height),
                            styles: extractComprehensiveStyles(child),
                            className: child.className || '',
                            id: child.id || '',
                            zIndex: 0,
                            mediaInfo: {
                                src: child.src || '',
                                alt: child.alt || '',
                                naturalWidth: child.naturalWidth || 0,
                                naturalHeight: child.naturalHeight || 0
                            }
                        });
                    } else if (childTag === 'span' || childTag === 'div') {
                        const spanRect = child.getBoundingClientRect();
                        elements.push({
                            type: childTag,
                            x: Math.round(spanRect.left - slideRect.left),
                            y: Math.round(spanRect.top - slideRect.top),
                            width: Math.round(spanRect.width),
                            height: Math.round(spanRect.height),
                            styles: extractComprehensiveStyles(child),
                            className: child.className || '',
                            id: child.id || '',
                            zIndex: 0,
                            text: child.textContent.trim()
                        });
                    }
                }
                return elements;
            }

            function isRedundantInlineElement(element) {
                const tag = element.tagName.toLowerCase();
                if (!['span', 'strong', 'b', 'em', 'i', 'u', 'mark'].includes(tag)) return false;
                const parent = element.parentElement;
                if (!parent) return false;
                const parentTag = parent.tagName.toLowerCase();
                
                if (
                    ['div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'td', 'th'].includes(parentTag) &&
                    !parent.closest('footer')
                ) {
                    const elementStyles = window.getComputedStyle(element);
                    const parentStyles = window.getComputedStyle(parent);
                    
                    const hasDifferentStyling = (
                        elementStyles.fontWeight !== parentStyles.fontWeight ||
                        elementStyles.fontStyle !== parentStyles.fontStyle ||
                        elementStyles.color !== parentStyles.color ||
                        elementStyles.backgroundColor !== parentStyles.backgroundColor ||
                        elementStyles.textDecoration !== parentStyles.textDecoration
                    );
                    
                    return !hasDifferentStyling;
                }
                return false;
            }

            function shouldProcessElement(element, slideContainer) {
                if (element === slideContainer) return false;
                if (!slideContainer.contains(element)) return false;
                const tagName = element.tagName.toLowerCase();
                if (!IMPORTANT_ELEMENTS.includes(tagName)) return false;
                const styles = window.getComputedStyle(element);
                if (styles.display === 'none' || styles.visibility === 'hidden' || styles.opacity === '0') return false;
                
                const hasBorder = ['borderTopWidth', 'borderRightWidth', 'borderBottomWidth', 'borderLeftWidth']
                    .some(prop => parseFloat(styles[prop]) > 0 && styles[prop.replace('Width', 'Style')] !== 'none');
                    
                const hasBackground = styles.backgroundColor !== 'rgba(0, 0, 0, 0)' && 
                                    styles.backgroundColor !== 'transparent' &&
                                    styles.backgroundColor !== 'initial';
                                    
                const hasShadow = styles.boxShadow !== 'none';
                const hasGradient = styles.background && styles.background.includes('gradient');
                
                if (hasBorder || hasBackground || hasShadow || hasGradient) {
                    return true;
                }
                
                if (['img', 'canvas', 'svg'].includes(tagName)) {
                    return true;
                }
                
                if (styles.clipPath.includes('polygon')) {
                    return true;
                }
                
                if (['strong', 'b', 'i', 'em', 'u', 'mark', 'li', 'td', 'th', 'span'].includes(tagName)) {
                    return true;
                }
                
                if (tagName === 'div') {
                    const parent = element.parentElement;
                    const parentStyles = parent ? window.getComputedStyle(parent) : null;
                    
                    if (parentStyles && (parentStyles.display === 'flex' || parentStyles.display === 'inline-flex')) {
                        return true;
                    }
                    
                    if (styles.display === 'flex' || styles.display === 'inline-flex') {
                        return true;
                    }
                    
                    if (Array.from(element.classList).some(cls => 
                        cls.includes('chart') || cls.includes('chartjs-chart')) || styles.clipPath.includes('polygon')) {
                        return true;
                    }
                    
                    if (styles.borderStyle !== 'none' && parseFloat(styles.borderWidth) > 0) {
                        return true;
                    }
                }
                
                const text = element.textContent ? element.textContent.trim() : '';
                if (!text && !['img', 'svg', 'video', 'audio', 'iframe', 'hr', 'br', 'canvas'].includes(tagName)) {
                    const rect = element.getBoundingClientRect();
                    if (rect.width === 0 && rect.height === 0) {
                        return false;
                    }
                    
                    if (tagName === 'div' && !hasBorder && !hasBackground && !hasShadow) {
                        return false;
                    }
                }
                
                return true;
            }

            function isOverlayElement(element) {
                const styles = window.getComputedStyle(element);
                const position = styles.position;
                const hasBorder = (
                    (parseFloat(styles.borderTopWidth) > 0 && styles.borderTopStyle !== 'none') ||
                    (parseFloat(styles.borderRightWidth) > 0 && styles.borderRightStyle !== 'none') ||
                    (parseFloat(styles.borderBottomWidth) > 0 && styles.borderBottomStyle !== 'none') ||
                    (parseFloat(styles.borderLeftWidth) > 0 && styles.borderLeftStyle !== 'none')
                );
                const hasZ = parseInt(styles.zIndex) > 0;
                const tag = element.tagName.toLowerCase();
                const isContentTag = [
                    'li', 'p', 'span', 'strong', 'b', 'em', 'i', 'u', 'mark'
                ].includes(tag);
                return (
                    (position === 'absolute' || position === 'fixed' || position === 'relative') &&
                    (hasBorder || styles.pointerEvents === 'none') &&
                    !isContentTag
                );
            }

            function getAllDescendants(element) {
                let descendants = [];
                function traverse(el) {
                    if (!el || !el.children) return;
                    for (let child of el.children) {
                        if (child && child.tagName) {
                            descendants.push(child);
                            traverse(child);
                        }
                    }
                }
                traverse(element);
                return descendants;
            }

            for (let slideIndex = 0; slideIndex < slideElements.length; slideIndex++) {
                const slideElement = slideElements[slideIndex];
                const slideRect = slideElement.getBoundingClientRect();
                const slide = {
                    slideId: slideIndex + 1,
                    elements: [],
                    slideWidth: parseFloat(slideRect.width.toFixed(2)),
                    slideHeight: parseFloat(slideRect.height.toFixed(2)),
                    slidePosition: {
                        x: parseFloat(slideRect.left.toFixed(2)),
                        y: parseFloat(slideRect.top.toFixed(2))
                    },
                    slideStyles: extractComprehensiveStyles(slideElement)
                };
                const processedElements = new Set();
                const processedTextElements = new Set();
                const allElements = Array.from(slideElement.querySelectorAll('*'));
                const elementsToProcess = allElements.filter(element => {
                    if (isRedundantInlineElement(element)) return false;
                    return shouldProcessElement(element, slideElement);
                });
                const overlayElements = [];
                
                elementsToProcess.forEach(element => {
                    if (isOverlayElement(element)) {
                        const position = getAccuratePosition(element, slideElement);
                        const styles = extractComprehensiveStyles(element);
                        let parentContent = element.parentElement;
                        while (parentContent && isOverlayElement(parentContent)) {
                            parentContent = parentContent.parentElement;
                        }
                        const overlayInfo = extractOverlayInfo(element, slideElement);
                        if (overlayInfo) {
                            overlayElements.push({
                                ...overlayInfo,
                                parentContentSelector: parentContent ? parentContent.tagName.toLowerCase() + (parentContent.className ? '.' + parentContent.className.split(' ').join('.') : '') : null
                            });
                        }
                        processedElements.add(getElementId(element));
                        for (const child of element.children) processedElements.add(getElementId(child));
                    }
                });
                elementsToProcess.sort((a, b) => {
                    const rectA = a.getBoundingClientRect();
                    const rectB = b.getBoundingClientRect();
                    const topDiff = rectA.top - rectB.top;
                    return Math.abs(topDiff) < 5 ? rectA.left - rectB.left : topDiff;
                });
                elementsToProcess.forEach(element => {
                    if (processedElements.has(getElementId(element))) return;
                    const elementId = getElementId(element);
                    const tagName = element.tagName.toLowerCase();
                    const position = getAccuratePosition(element, slideElement);
                    const styles = extractComprehensiveStyles(element);
                    
                    const shapeInfo = extractShapeInfo(element, slideElement);

                    const shouldExtractInline = shouldExtractInlineGroup(element, processedTextElements);
                    const inlineGroup = shouldExtractInline ? getInlineGroup(element, slideElement, processedTextElements) : null;
                    
                    if (processedElements.has(elementId)) {
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
                        inlineGroup: inlineGroup,
                        shapeInfo: shapeInfo
                    };
                    
                    if (inlineGroup) {
                        processedTextElements.add(elementId);
                        getAllDescendants(element).forEach(desc => {
                            processedTextElements.add(getElementId(desc));
                        });
                    }

                    if (shapeInfo) {
                        processedElements.add(elementId);
                        getAllDescendants(element).forEach(desc => {
                            processedElements.add(getElementId(desc));
                        });
                    }

                    if (tagName === 'img') {
                        const imgDimensions = getAccurateImageDimensions(element, element.parentElement, slideElement);
                        elementData.mediaInfo = {
                            src: element.src || '',
                            alt: element.alt || '',
                            naturalWidth: element.naturalWidth || 0,
                            naturalHeight: element.naturalHeight || 0,
                            currentWidth: imgDimensions.width,
                            currentHeight: imgDimensions.height
                        };
                        processedTextElements.add(elementId);
                    }

                    if (tagName === 'canvas' && (element.classList.contains('chartjs-chart') || element.closest('.chart'))) {
                        const canvasRect = element.getBoundingClientRect();
                        const slideRect = slideElement.getBoundingClientRect();
                        let chartConfig = null;
                        try {
                            const chartData = element.getAttribute('data-chart');
                            if (chartData) {
                                chartConfig = JSON.parse(chartData);
                            }
                        } catch(e) {
                            console.log('Could not parse chart data for element', element.id);
                            chartConfig = element.getAttribute('data-chart');
                        }
                        
                        elementData.chartInfo = {
                            chartId: element.id || '',
                            chartClass: element.className || '',
                            chartData: chartConfig,
                            width: Math.round(canvasRect.width),
                            height: Math.round(canvasRect.height),
                            attributeWidth: element.width || element.getAttribute('width') || 0,
                            attributeHeight: element.height || element.getAttribute('height') || 0,
                            canvasRect: {
                                x: Math.round(canvasRect.left - slideRect.left),
                                y: Math.round(canvasRect.top - slideRect.top),
                                width: Math.round(canvasRect.width),
                                height: Math.round(canvasRect.height)
                            }
                        };
                        processedTextElements.add(elementId);
                    }

                    if (tagName === 'svg') {
                        elementData.svgInfo = {
                            svgContent: element.outerHTML,
                            viewBox: element.getAttribute('viewBox') || '',
                            width: element.getAttribute('width') || position.width,
                            height: element.getAttribute('height') || position.height
                        };
                        processedTextElements.add(elementId);
                    }
                    
                    if (!inlineGroup && ['div', 'span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(tagName)) {
                        if (shouldExtractTextContent(element, processedTextElements)) {
                            const text = getTextContent(element, processedElements);
                            if (text) {
                                elementData.text = text;
                                processedTextElements.add(elementId);
                                getAllDescendants(element).forEach(desc => {
                                    processedTextElements.add(getElementId(desc));
                                });
                            }
                        }
                    }

                    if (['ul', 'ol'].includes(tagName)) {
                        elementData.listInfo = getListInfo(element, slideElement);
                        processedTextElements.add(elementId);
                        getAllDescendants(element).forEach(desc => {
                            processedTextElements.add(getElementId(desc));
                        });
                    }
                    if (['table'].includes(tagName)) {
                        elementData.tableInfo = getTableInfo(element, slideElement);
                        processedTextElements.add(elementId);
                        getAllDescendants(element).forEach(desc => {
                            processedTextElements.add(getElementId(desc));
                        });
                    }

                    if (element.classList.contains('company')) {
                        // This class name is too generic, avoid special handling
                    }

                    if (element.tagName.toLowerCase() === 'footer' || (element.closest('footer') && element.children.length > 0)) {
                        const footerElements = extractFooterElements(element, slideElement);
                        if (footerElements.length > 0) {
                            elementData.footerElements = footerElements;
                            getAllDescendants(element).forEach(desc => {
                                processedElements.add(getElementId(desc));
                                processedTextElements.add(getElementId(desc));
                            });
                        }
                    }

                    slide.elements.push(elementData);

                    const before = extractPseudo(element, slideElement, '::before');
                    if (before) slide.elements.push(before);
                    const after = extractPseudo(element, slideElement, '::after');
                    if (after) slide.elements.push(after);

                    if (['ul', 'ol', 'table'].includes(tagName)) {
                        getAllDescendants(element).forEach(desc => {
                            processedElements.add(getElementId(desc));
                        });
                    }
                });

                const missedElements = Array.from(slideElement.querySelectorAll('img, canvas, svg')).filter(element => 
                    !processedElements.has(getElementId(element))
                );
                
                missedElements.forEach(element => {
                    const elementId = getElementId(element);
                    const tagName = element.tagName.toLowerCase();
                    const position = getAccuratePosition(element, slideElement);
                    const styles = extractComprehensiveStyles(element);
                    
                    const elementData = {
                        type: tagName,
                        x: position.x,
                        y: position.y,
                        width: position.width,
                        height: position.height,
                        styles,
                        className: element.className || '',
                        id: element.id || '',
                        zIndex: parseInt(styles.zIndex) || 0
                    };

                    if (tagName === 'img') {
                        const imgDimensions = getAccurateImageDimensions(element, element.parentElement, slideContainer);
                        elementData.mediaInfo = {
                            src: element.src || '',
                            alt: element.alt || '',
                            naturalWidth: element.naturalWidth || 0,
                            naturalHeight: element.naturalHeight || 0,
                            currentWidth: imgDimensions.width,
                            currentHeight: imgDimensions.height
                        };
                    }

                    if (tagName === 'canvas' && (element.classList.contains('chartjs-chart') || element.closest('.chart'))) {
                        const canvasRect = element.getBoundingClientRect();
                        const slideRect = slideElement.getBoundingClientRect();
                        let chartConfig = null;
                        try {
                            const chartData = element.getAttribute('data-chart');
                            if (chartData) {
                                chartConfig = JSON.parse(chartData);
                            }
                        } catch(e) {
                            console.log('Could not parse chart data for element', element.id);
                            chartConfig = element.getAttribute('data-chart');
                        }
                        
                        elementData.chartInfo = {
                            chartId: element.id || '',
                            chartClass: element.className || '',
                            chartData: chartConfig,
                            width: Math.round(canvasRect.width),
                            height: Math.round(canvasRect.height),
                            attributeWidth: element.width || element.getAttribute('width') || 0,
                            attributeHeight: element.height || element.getAttribute('height') || 0,
                            canvasRect: {
                                x: Math.round(canvasRect.left - slideRect.left),
                                y: Math.round(canvasRect.top - slideRect.top),
                                width: Math.round(canvasRect.width),
                                height: Math.round(canvasRect.height)
                            }
                        };
                    }

                    if (tagName === 'svg') {
                        elementData.svgInfo = {
                            svgContent: element.outerHTML,
                            viewBox: element.getAttribute('viewBox') || '',
                            width: element.getAttribute('width') || position.width,
                            height: element.getAttribute('height') || position.height
                        };
                    }

                    slide.elements.push(elementData);
                    processedElements.add(elementId);
                });

                slide.elements.push(...overlayElements);
                slide.elements.sort((a, b) => a.zIndex - b.zIndex || a.y - b.y || a.x - b.x);
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

const htmlFilePath = 'dolbix.html';
const outputPath = 'slides_data.json';
extractSlideData(htmlFilePath, outputPath).catch(err => {
    console.error('Error:', err);
});