const puppeteer = require('puppeteer');
const fs = require('fs').promises;

async function extractSlideData(htmlFilePath, outputPath) {
    const browser = await puppeteer.launch({ 
        headless: false, 
        devtools: false,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();
    
    // Read HTML content first to get initial dimensions
    const htmlContent = await fs.readFile(htmlFilePath, 'utf-8');
    
    // Set initial viewport - we'll adjust this based on content
    await page.setViewport({ width: 1280, height: 720 });
    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });

    // Wait for images to load
    await page.waitForFunction(() => {
        const images = Array.from(document.querySelectorAll('img'));
        return images.every(img => img.complete);
    }, { timeout: 10000 }).catch(() => console.log('Some images may not have loaded'));

    // Calculate actual document dimensions
    const documentDimensions = await page.evaluate(() => {
        // Get the actual content dimensions
        const body = document.body;
        const html = document.documentElement;
        
        // Calculate actual width and height including all content
        const actualWidth = Math.max(
            body.scrollWidth,
            body.offsetWidth,
            html.clientWidth,
            html.scrollWidth,
            html.offsetWidth
        );
        
        const actualHeight = Math.max(
            body.scrollHeight,
            body.offsetHeight,
            html.clientHeight,
            html.scrollHeight,
            html.offsetHeight
        );

        // Also get viewport dimensions for reference
        const viewportWidth = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
        const viewportHeight = Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0);

        return {
            actualWidth,
            actualHeight,
            viewportWidth,
            viewportHeight
        };
    });

    console.log('Document dimensions:', documentDimensions);

    // Adjust viewport to match content if needed
    const targetWidth = Math.max(documentDimensions.actualWidth, 1280);
    const targetHeight = Math.max(documentDimensions.actualHeight, 720);
    
    await page.setViewport({ 
        width: targetWidth, 
        height: targetHeight 
    });


    const slideData = await page.evaluate(async (dimensions) => {
        const slide = { 
            slideId: 1, 
            elements: [],
            slideWidth: dimensions.actualWidth,
            slideHeight: dimensions.actualHeight,
            viewportWidth: dimensions.viewportWidth,
            viewportHeight: dimensions.viewportHeight
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

        function getAccuratePosition(element) {
            const rect = element.getBoundingClientRect();
            const styles = window.getComputedStyle(element);
            
            // Account for document scroll position
            const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
            const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
            
            // Get more accurate positioning
            let x = rect.left + scrollLeft;
            let y = rect.top + scrollTop;
            
            // Account for borders and padding in positioning if needed
            const borderLeft = parseFloat(styles.borderLeftWidth) || 0;
            const borderTop = parseFloat(styles.borderTopWidth) || 0;
            
            // Round to avoid sub-pixel positioning issues
            x = Math.round(x);
            y = Math.round(y);
            
            // Ensure minimum values
            x = Math.max(0, x);
            y = Math.max(0, y);
            
            return {
                x: x,
                y: y,
                width: Math.round(rect.width),
                height: Math.round(rect.height),
                // Store original rect for debugging
                originalRect: {
                    left: rect.left,
                    top: rect.top,
                    width: rect.width,
                    height: rect.height
                }
            };
        }

        // Ensure we start from top-left
        window.scrollTo(0, 0);
        await new Promise(resolve => setTimeout(resolve, 100));
        
        const allElements = Array.from(document.querySelectorAll('body *'));
        const elementsToProcess = allElements.filter(element => {
            return shouldProcessElement(element);
        });

        // Sort elements by position (top to bottom, left to right)
        elementsToProcess.sort((a, b) => {
            const rectA = a.getBoundingClientRect();
            const rectB = b.getBoundingClientRect();
            const topDiff = rectA.top - rectB.top;
            return Math.abs(topDiff) < 5 ? rectA.left - rectB.left : topDiff;
        });

        elementsToProcess.forEach(element => {
            const elementId = getElementId(element);
            processedElements.add(elementId);

            const position = getAccuratePosition(element);
            const styles = extractStyles(element);

            const elementData = {
                type: element.tagName.toLowerCase(),
                x: position.x,
                y: position.y,
                width: position.width,
                height: position.height,
                slideWidth: dimensions.actualWidth,
                slideHeight: dimensions.actualHeight,
                styles: styles,
                className: element.className,
                id: element.id,
                zIndex: parseInt(styles.zIndex) || 0,
                // Add positioning debug info
                positionInfo: {
                    originalRect: position.originalRect,
                    computedPosition: styles.position,
                    scrollOffset: {
                        x: window.pageXOffset || document.documentElement.scrollLeft,
                        y: window.pageYOffset || document.documentElement.scrollTop
                    }
                }
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

                    if (textMetrics.textWidth > position.width && styles.whiteSpace !== 'nowrap') {
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
                    
                    // Calculate scaling information
                    elementData.scaling = {
                        scaleX: position.width / element.naturalWidth,
                        scaleY: position.height / element.naturalHeight,
                        aspectRatio: element.naturalWidth / element.naturalHeight
                    };
                }
            }

            slide.elements.push(elementData);
        });

        return [slide];
    }, documentDimensions);

    await fs.writeFile(outputPath, JSON.stringify(slideData, null, 2), 'utf-8');
    console.log(`Successfully extracted 1 slide to ${outputPath}`);
    console.log(`Actual dimensions: ${documentDimensions.actualWidth}x${documentDimensions.actualHeight}`);
    console.log(`Viewport dimensions: ${documentDimensions.viewportWidth}x${documentDimensions.viewportHeight}`);
    console.log(`Elements extracted: ${slideData[0].elements.length}`);
    
    await browser.close();
}

const htmlFilePath = 'input.html';
const outputPath = 'slides_data.json';

extractSlideData(htmlFilePath, outputPath).catch(err => {
    console.error('Error:', err);
});