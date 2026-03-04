const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel,
    convertInchesToTwip
} = require('docx');

// Parse inline styles from HTML element
function parseStyle(styleStr = '') {
    const styles = {};
    if (!styleStr) return styles;
    styleStr.split(';').forEach(rule => {
        const parts = rule.split(':');
        if (parts.length >= 2) {
            const prop = parts[0].trim().toLowerCase();
            const val = parts.slice(1).join(':').trim().toLowerCase().replace(/['"]/g, '');
            styles[prop] = val;
        }
    });
    return styles;
}

// Convert hex color to docx format (remove #)
function hexColor(color = '') {
    if (!color) return undefined;
    color = color.trim();
    if (color.startsWith('#')) {
        const hex = color.slice(1).toUpperCase();
        // Expand 3-digit hex to 6-digit
        if (hex.length === 3) {
            return hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
        }
        return hex;
    }

    const named = { 'white': 'FFFFFF', 'black': '000000', 'red': 'FF0000' };
    return named[color] || undefined;
}
// Convert font-size string to half-points (docx size unit)
function fontSize(sizeStr = '') {
    const match = sizeStr.match(/(\d+(?:\.\d+)?)(px|pt|em)?/);
    if (!match) return undefined;
    const val = parseFloat(match[1]);
    const unit = match[2] || 'px';
    if (unit === 'pt') return Math.round(val * 2);
    if (unit === 'px') return Math.round((val * 0.75) * 2); // px to pt * 2
    return Math.round(val * 24); // em * 12pt * 2
}

// Parse HTML string into DOM-like structure using regex
function parseHTML(html) {
    // Simple HTML parser - extract body content
    const bodyMatch = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    const content = bodyMatch ? bodyMatch[1] : html;
    return content;
}

// Extract class styles from <style> block
function extractStyles(html) {
    const styleMap = {};
    const styleMatch = html.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
    if (!styleMatch) return styleMap;

    const cssText = styleMatch[1];
    const ruleRegex = /\.([a-zA-Z0-9_-]+)\s*\{([^}]+)\}/g;
    let match;
    while ((match = ruleRegex.exec(cssText)) !== null) {
        const className = match[1];
        const rules = parseStyle(match[2]);
        styleMap[className] = rules;
    }
    return styleMap;
}

// Convert HTML content to docx paragraphs
function htmlToDocxElements(html, classStyles) {
    const elements = [];
    const bodyContent = parseHTML(html);

    // Extract all top-level blocks
    function extractBlocks(content) {
        const result = [];
        let depth = 0;
        let i = 0;
        let currentBlock = '';

        while (i < content.length) {
            if (content[i] === '<') {
                const tagMatch = content.slice(i).match(/^<(\/?)(div|p|h[1-6]|hr|br|table|tr|td)([^>]*)>/i);
                if (tagMatch) {
                    const isClose = tagMatch[1] === '/';
                    const tag = tagMatch[2].toLowerCase();
                    const attrs = tagMatch[3];
                    const fullTag = tagMatch[0];

                    if (tag === 'hr' || tag === 'br') {
                        currentBlock += fullTag;
                        i += fullTag.length;
                        continue;
                    }

                    if (!isClose) {
                        if (depth === 0) {
                            if (currentBlock.trim()) result.push({ type: 'text', content: currentBlock });
                            currentBlock = fullTag;
                            depth = 1;
                        } else {
                            currentBlock += fullTag;
                            depth++;
                        }
                    } else {
                        depth--;
                        currentBlock += fullTag;
                        if (depth === 0) {
                            result.push({ type: 'block', content: currentBlock });
                            currentBlock = '';
                        }
                    }
                    i += fullTag.length;
                    continue;
                }
            }
            currentBlock += content[i];
            i++;
        }
        if (currentBlock.trim()) result.push({ type: 'text', content: currentBlock });
        return result;
    }

    const blocks2 = extractBlocks(bodyContent);

    for (const block of blocks2) {
        const parsed = processBlock(block.content, classStyles);
        if (parsed) elements.push(...(Array.isArray(parsed) ? parsed : [parsed]));
    }

    return elements;
}

function stripTags(html) {
    return html.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').trim();
}

function getClassAndStyle(tagStr) {
    const classMatch = tagStr.match(/class="([^"]+)"/i);
    const styleMatch = tagStr.match(/style="([^"]+)"/i);
    return {
        className: classMatch ? classMatch[1].trim() : '',
        inlineStyle: styleMatch ? styleMatch[1] : ''
    };
}

function processBlock(blockHtml, classStyles) {
    if (!blockHtml || !blockHtml.trim()) return null;

    // HR → horizontal rule paragraph
    if (/^<hr/i.test(blockHtml.trim())) {
        return new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CCCCCC', space: 1 } },
            spacing: { before: 120, after: 120 }
        });
    }

    const tagMatch = blockHtml.match(/^<(div|p|h[1-6])([^>]*)>([\s\S]*)<\/\1>$/i);
    if (!tagMatch) {
        const text = stripTags(blockHtml);
        if (!text) return null;
        return new Paragraph({ children: [new TextRun(text)] });
    }

    const tag = tagMatch[1].toLowerCase();
    const attrs = tagMatch[2];
    const innerHtml = tagMatch[3];
    const { className, inlineStyle } = getClassAndStyle(attrs);

    // Merge class styles + inline styles
    const classStyle = classStyles[className] || {};
    const inlineStyleObj = parseStyle(inlineStyle);
    const mergedStyle = { ...classStyle, ...inlineStyleObj };

    // Determine alignment
    let alignment = AlignmentType.LEFT;
    if (mergedStyle['text-align'] === 'center') alignment = AlignmentType.CENTER;
    if (mergedStyle['text-align'] === 'right') alignment = AlignmentType.RIGHT;

    // Background color for shading
    const bgColor = hexColor(mergedStyle['background-color']);
    const textColor = hexColor(mergedStyle['color']);
    const isBold = mergedStyle['font-weight'] === 'bold' || tag.match(/^h[1-6]$/);
    const fSize = fontSize(mergedStyle['font-size']);

    // Parse inner content for inline elements
    const runs = parseInlineContent(innerHtml, { color: textColor, bold: isBold, size: fSize }, classStyles);

    // Spacing
    const spacingBefore = mergedStyle['margin-top'] ? parseInt(mergedStyle['margin-top']) * 15 : 120;
    const spacingAfter = mergedStyle['margin-bottom'] ? parseInt(mergedStyle['margin-bottom']) * 15 : 80;

    const paraProps = {
        alignment,
        spacing: { before: spacingBefore, after: spacingAfter },
        children: runs
    };

    // Add shading/background as a table cell if has background color
    if (bgColor) {
        const contentWidth = 8640; // ~6 inches in DXA
        return new Table({
            width: { size: contentWidth, type: WidthType.DXA },
            columnWidths: [contentWidth],
            margins: { top: 60, bottom: 60 },
            borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE },
                insideH: { style: BorderStyle.NONE },
                insideV: { style: BorderStyle.NONE },
            },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            width: { size: contentWidth, type: WidthType.DXA },
                            shading: { fill: bgColor, type: ShadingType.CLEAR },
                            margins: { top: 100, bottom: 100, left: 180, right: 180 },
                            borders: {
                                top: { style: BorderStyle.NONE },
                                bottom: { style: BorderStyle.NONE },
                                left: { style: BorderStyle.NONE },
                                right: { style: BorderStyle.NONE },
                            },
                            children: [new Paragraph({ alignment, children: runs })]
                        })
                    ]
                })
            ]
        });
    }

    return new Paragraph(paraProps);
}

function parseInlineContent(html, defaultProps = {}, classStyles = {}) {
    const runs = [];

    if (!html || !html.trim()) return [new TextRun('')];

    // Split by inline tags
    const parts = html.split(/(<[^>]+>)/);
    let currentProps = { ...defaultProps };

    for (let i = 0; i < parts.length; i++) {
        const part = parts[i];

        if (!part) continue;

        if (part.startsWith('<')) {
            if (/^<br/i.test(part)) {
                runs.push(new TextRun({ text: '', break: 1 }));
                continue;
            }
            if (/^<b>|^<strong>/i.test(part)) {
                currentProps = { ...currentProps, bold: true };
                continue;
            }
            if (/^<\/b>|^<\/strong>/i.test(part)) {
                currentProps = { ...currentProps, bold: defaultProps.bold };
                continue;
            }
            if (/^<span/i.test(part)) {
                const styleMatch = part.match(/style="([^"]+)"/i);
                const classMatch = part.match(/class="([^"]+)"/i);
                const spanStyle = parseStyle(styleMatch ? styleMatch[1] : '');
                const spanClassStyle = classStyles[classMatch ? classMatch[1] : ''] || {};
                const merged = { ...spanClassStyle, ...spanStyle };

                if (merged['color']) currentProps = { ...currentProps, color: hexColor(merged['color']) };
                if (merged['font-weight'] === 'bold') currentProps = { ...currentProps, bold: true };
                if (merged['font-size']) currentProps = { ...currentProps, size: fontSize(merged['font-size']) };
                continue;
            }
            if (/^<\/span>/i.test(part)) {
                currentProps = { ...defaultProps };
                continue;
            }
            continue;
        }

        const text = part.replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
        if (text) {
            const runProps = { text };
            if (currentProps.bold) runProps.bold = true;
            if (currentProps.color) runProps.color = currentProps.color;
            if (currentProps.size) runProps.size = currentProps.size;
            runs.push(new TextRun(runProps));
        }
    }

    return runs.length > 0 ? runs : [new TextRun('')];
}

async function convertHtmlToDocx(html) {
    const classStyles = extractStyles(html);
    const elements = htmlToDocxElements(html, classStyles);

    const doc = new Document({
        sections: [{
            properties: {
                page: {
                    size: { width: 11906, height: 16838 }, // A4
                    margin: { top: 720, right: 720, bottom: 720, left: 720 } // 0.5 inch margins
                }
            },
            children: elements.filter(Boolean)
        }]
    });

    return await Packer.toBuffer(doc);
}

module.exports = async (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    try {
        const { html, filename = 'document.docx' } = req.body;

        if (!html) {
            return res.status(400).json({ error: 'html field is required' });
        }

        const buffer = await convertHtmlToDocx(html);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${filename.endsWith('.docx') ? filename : filename + '.docx'}"`);
        res.setHeader('Content-Length', buffer.length);

        return res.send(buffer);
    } catch (err) {
        console.error('Conversion error:', err);
        return res.status(500).json({ error: 'Conversion failed', details: err.message });
    }
};
