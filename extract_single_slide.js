const puppeteer = require('puppeteer');
const fs = require('fs').promises;

async function extractSlideData(htmlFilePath, outputPath) {
    const browser = await puppeteer.launch({ 
        headless: false, 
        devtools: false,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();
    
    await page.setViewport({ width: 1280, height: 720 });
    
    const htmlContent = await fs.readFile(htmlFilePath, 'utf-8');
    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });

    await page.waitForFunction(() => {
        const images = Array.from(document.querySelectorAll('img'));
        return images.every(img => img.complete);
    }, { timeout: 10000 }).catch(() => console.log('Some images may not have loaded'));

    const slideData = await page.evaluate(async () => {
        const slide = { 
            slideId: 1, 
            elements: [],
            slideWidth: 1280,
            slideHeight: 720
        };
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
                return element.src && element.src.trim() !== '';
            }

            return hasVisibleText;
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
                overflow: styles.overflow,
                whiteSpace: styles.whiteSpace,
                wordWrap: styles.wordWrap,
                textOverflow: styles.textOverflow
            };
        }

        function getTextContent(element) {
            const styles = window.getComputedStyle(element);
            const text = element.textContent.trim();
            
            if (styles.whiteSpace === 'nowrap') {
                return text.replace(/\s+/g, ' ');
            }
            
            return text;
        }

        function calculateTextMetrics(element, text) {
            const styles = window.getComputedStyle(element);
            const canvas = document.createElement('canvas');
            const context = canvas.getContext('2d');
            
            const fontSize = parseFloat(styles.fontSize);
            const fontFamily = styles.fontFamily.replace(/"/g, '').split(',')[0].trim();
            const fontWeight = styles.fontWeight;
            
            context.font = `${fontWeight} ${fontSize}px ${fontFamily}`;
            
            const metrics = context.measureText(text);
            const textWidth = metrics.width;
            const textHeight = fontSize * 1.2;
            
            return {
                textWidth: textWidth,
                textHeight: textHeight,
                fontSize: fontSize,
                fontFamily: fontFamily,
                fontWeight: fontWeight
            };
        }

        window.scrollTo(0, 0);
        await new Promise(resolve => setTimeout(resolve, 100));
        
        const allElements = Array.from(document.querySelectorAll('body *'));
        const elementsToProcess = allElements.filter(element => {
            const rect = element.getBoundingClientRect();
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
            
            const x = Math.max(0, Math.round(rect.left));
            const y = Math.max(0, Math.round(rect.top));
            const width = Math.max(0, Math.round(rect.width));
            const height = Math.max(0, Math.round(rect.height));

            const elementData = {
                type: element.tagName.toLowerCase(),
                x: x,
                y: y,
                width: width,
                height: height,
                slideWidth: 1280,
                slideHeight: 720,
                styles: styles,
                className: element.className,
                id: element.id,
                zIndex: parseInt(styles.zIndex) || 0
            };

            if (['div', 'span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(element.tagName.toLowerCase())) {
                const directText = Array.from(element.childNodes)
                    .filter(node => node.nodeType === Node.TEXT_NODE)
                    .map(node => node.textContent.trim())
                    .join(' ');
                
                const text = directText || getTextContent(element);
                
                if (text) {
                    elementData.text = text;
                    const textMetrics = calculateTextMetrics(element, text);
                    elementData.textMetrics = textMetrics;

                    if (textMetrics.textWidth > width && styles.whiteSpace !== 'nowrap') {
                        elementData.needsWrapping = true;
                    }
                }
            } 
            else if (element.tagName.toLowerCase() === 'img') {
                elementData.src = element.src;
                elementData.alt = element.alt || '';

                if (element.naturalWidth && element.naturalHeight) {
                    elementData.naturalWidth = element.naturalWidth;
                    elementData.naturalHeight = element.naturalHeight;
                }
            }

            slide.elements.push(elementData);
        });

        return [slide];
    });

    await fs.writeFile(outputPath, JSON.stringify(slideData, null, 2), 'utf-8');
    console.log(`Successfully extracted 1 slide to ${outputPath}`);
    console.log(`Target resolution: 1280x720 (720p)`);
    
    await browser.close();
}

const htmlFilePath = 'input.html';
const outputPath = 'slides_data.json';

extractSlideData(htmlFilePath, outputPath).catch(err => {
    console.error('Error:', err);
});