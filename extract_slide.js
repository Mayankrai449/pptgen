const puppeteer = require('puppeteer');
const fs = require('fs').promises;

async function extractSlideData(htmlFilePath, outputPath) {
    const browser = await puppeteer.launch({ headless: false, devtools: false });
    const page = await browser.newPage();
    
    await page.setViewport({ width: 1920, height: 1080 });
    
    const htmlContent = await fs.readFile(htmlFilePath, 'utf-8');
    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });

    const slidesData = await page.evaluate(async () => {
        const slides = [];
        const slideDivs = document.querySelectorAll('.slides');
        const processedElements = new Set();

        function getElementId(element) {
            const rect = element.getBoundingClientRect();
            const text = element.textContent ? element.textContent.trim().substring(0, 50) : '';
            const src = element.src || '';
            return `${element.tagName}-${rect.left}-${rect.top}-${rect.width}-${rect.height}-${text}-${src}`;
        }

        function shouldProcessElement(element) {
            const elementId = getElementId(element);
            if (processedElements.has(elementId)) return false;
            
            const rect = element.getBoundingClientRect();
            if (rect.width === 0 || rect.height === 0) return false;

            const styles = window.getComputedStyle(element);
            const hasVisibleText = element.textContent && element.textContent.trim().length > 0;
            const isImage = element.tagName.toLowerCase() === 'img';
            const hasBackground = styles.backgroundColor !== 'rgba(0, 0, 0, 0)' && styles.backgroundColor !== 'transparent';
            const hasBorder = styles.border !== 'none' && styles.borderWidth !== '0px';
            const hasBorderRadius = styles.borderRadius !== '0px';
            const isVisible = styles.display !== 'none' && styles.visibility !== 'hidden' && styles.opacity !== '0';

            if (!isVisible) return false;

            if (element.tagName.toLowerCase() === 'div') {
                const hasDirectText = Array.from(element.childNodes).some(node => 
                    node.nodeType === Node.TEXT_NODE && node.textContent.trim().length > 0
                );
                return hasDirectText || hasBackground || hasBorder || hasBorderRadius;
            }

            if (element.tagName.toLowerCase() === 'span') {
                return hasVisibleText;
            }

            if (isImage) {
                return true;
            }

            return false;
        }

        function extractStyles(element) {
            const styles = window.getComputedStyle(element);
            
            return {
                fontSize: styles.fontSize,
                fontFamily: styles.fontFamily.replace(/"/g, '').split(',')[0].trim(),
                fontWeight: styles.fontWeight,
                fontStyle: styles.fontStyle,
                textAlign: styles.textAlign,
                textDecoration: styles.textDecoration,
                lineHeight: styles.lineHeight,
                letterSpacing: styles.letterSpacing,
                color: styles.color,
                backgroundColor: styles.backgroundColor,
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
                borderTop: styles.borderTop,
                borderRight: styles.borderRight,
                borderBottom: styles.borderBottom,
                borderLeft: styles.borderLeft,
                borderRadius: styles.borderRadius,
                borderColor: styles.borderColor,
                borderWidth: styles.borderWidth,
                borderStyle: styles.borderStyle,
                position: styles.position,
                display: styles.display,
                visibility: styles.visibility,
                opacity: styles.opacity,
                zIndex: styles.zIndex,
                boxShadow: styles.boxShadow,
                transform: styles.transform,
                overflow: styles.overflow
            };
        }

        for (let slideIndex = 0; slideIndex < slideDivs.length; slideIndex++) {
            document.querySelectorAll('.slides').forEach((div, idx) => {
                div.style.display = idx === slideIndex ? 'block' : 'none';
            });

            const slideDiv = slideDivs[slideIndex];
            const slideContent = { slideId: slideIndex + 1, elements: [] };

            window.scrollTo(0, 0);
            const slideRect = slideDiv.getBoundingClientRect();
            const slideWidth = slideRect.width;
            const slideHeight = slideRect.height;

            const allElements = Array.from(slideDiv.querySelectorAll('*'));
            const elementsToProcess = allElements.filter(element => {
                const rect = element.getBoundingClientRect();
                const relativeX = rect.left - slideRect.left;
                const relativeY = rect.top - slideRect.top;
                if (relativeX < -10 || relativeY < -10 || relativeX > slideWidth + 10 || relativeY > slideHeight + 10) {
                    return false;
                }
                return shouldProcessElement(element);
            });

            elementsToProcess.sort((a, b) => {
                const rectA = a.getBoundingClientRect();
                const rectB = b.getBoundingClientRect();
                const topDiff = rectA.top - rectB.top;
                return Math.abs(topDiff) < 5 ? rectA.left - rectB.left : topDiff;
            });

            elementsToProcess.forEach(element => {
                const elementId = getElementId(element);
                processedElements.add(elementId);

                const rect = element.getBoundingClientRect();
                const styles = extractStyles(element);
                
                const relativeX = rect.left - slideRect.left;
                const relativeY = rect.top - slideRect.top;
                const maxWidth = slideWidth - relativeX;
                const maxHeight = slideHeight - relativeY;
                const clampedWidth = Math.min(rect.width, maxWidth);
                const clampedHeight = Math.min(rect.height, maxHeight);

                const elementData = {
                    type: element.tagName.toLowerCase(),
                    x: Math.max(0, relativeX),
                    y: Math.max(0, relativeY),
                    width: Math.max(0, clampedWidth),
                    height: Math.max(0, clampedHeight),
                    slideWidth: slideWidth,
                    slideHeight: slideHeight,
                    styles: styles,
                    className: element.className,
                    id: element.id,
                    zIndex: styles.zIndex
                };

                if (['div', 'span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(element.tagName.toLowerCase())) {
                    const directText = Array.from(element.childNodes)
                        .filter(node => node.nodeType === Node.TEXT_NODE)
                        .map(node => node.textContent.trim())
                        .join(' ');
                    elementData.text = directText || element.textContent.trim();
                } else if (element.tagName.toLowerCase() === 'img') {
                    elementData.src = element.src;
                    elementData.alt = element.alt || '';
                }

                slideContent.elements.push(elementData);
            });

            if (slideContent.elements.length > 0) {
                slides.push(slideContent);
            }
            processedElements.clear();
        }

        return slides;
    });

    await fs.writeFile(outputPath, JSON.stringify(slidesData, null, 2));
    await browser.close();
    console.log(`Extracted ${slidesData.length} slides with ${slidesData.reduce((sum, slide) => sum + slide.elements.length, 0)} total elements`);
    return slidesData;
}

extractSlideData('input.html', 'slides_data.json')
    .then(() => console.log('Slide data saved to slides_data.json'))
    .catch(err => console.error('Error:', err));