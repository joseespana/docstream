/**
 * Excel Spreadsheet (XLSX) Parser
 * 
 * **XLSX Format Overview:**
 * XLSX is the default format for Microsoft Excel since Office 2007, based on OOXML.
 * 
 * **File Structure:**
 * - `xl/workbook.xml` - Workbook structure and sheet list
 * - `xl/worksheets/sheet1.xml` - Individual sheet data
 * - `xl/sharedStrings.xml` - Shared string table (cell text)
 * - `xl/styles.xml` - Cell styling information
 * - `xl/drawings/*` - Charts and drawings
 * - `xl/media/*` - Embedded images
 * 
 * **Key Elements:**
 * - `<row>` - Table row with row index
 * - `<c r="A1">` - Cell with reference (A1, B2, etc.)
 * - `<v>` - Cell value (number or shared string index)
 * - `<t="s">` - Cell type (s=string, n=number, b=boolean)
 * 
 * @module ExcelParser
 * @see https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
 */

import { CellMetadata, ChartMetadata, ImageMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeParserConfig, TextFormatting, TextMetadata } from '../types';
import { extractChartData } from '../utils/chartUtils';
import { logWarning } from '../utils/errorUtils';
import { createAttachment } from '../utils/imageUtils';
import { performOcr } from '../utils/ocrUtils';
import { astToMarkdown } from '../utils/markdownUtils';
import { getElementsByTagName, parseAppMetadata, parseOfficeMetadata, parseXmlString } from '../utils/xmlUtils';
import { extractFiles } from '../utils/zipUtils';

/**
 * Converts an Excel column letter string (e.g., "A", "Z", "AA", "AZ") to a 0-based column index.
 */
const colLetterToIndex = (colStr: string): number => {
    let result = 0;
    for (let i = 0; i < colStr.length; i++) {
        result = result * 26 + (colStr.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result - 1;
};

/**
 * Parses an Excel spreadsheet (.xlsx) and extracts sheets, rows, and cells.
 * 
 * @param buffer - The XLSX file as a Buffer
 * @param config - Parser configuration
 * @returns A promise resolving to the parsed AST
 */
export const parseExcel = async (buffer: Buffer, config: OfficeParserConfig): Promise<OfficeParserAST> => {
    const sheetsRegex = /xl\/worksheets\/sheet\d+.xml/g;
    const drawingsRegex = /xl\/drawings\/drawing\d+.xml/g;
    const chartsRegex = /xl\/charts\/chart\d+.xml/g;
    const stringsFilePath = 'xl/sharedStrings.xml';
    const mediaFileRegex = /xl\/media\/.*/;
    const corePropsFileRegex = /docProps\/core\.xml/;

    const relsRegex = /xl\/worksheets\/_rels\/sheet\d+\.xml\.rels/g;
    const drawingRelsRegex = /xl\/drawings\/_rels\/drawing\d+\.xml\.rels/g;

    const files = await extractFiles(buffer, (x: string) =>
        !!x.match(sheetsRegex) ||
        !!x.match(drawingsRegex) ||
        !!x.match(chartsRegex) ||
        x === stringsFilePath ||
        x === 'xl/styles.xml' ||
        x === 'xl/workbook.xml' ||
        x === 'xl/_rels/workbook.xml.rels' ||
        !!x.match(corePropsFileRegex) ||
        x === 'docProps/app.xml' ||
        !!x.match(relsRegex) ||
        (!!config.extractAttachments && (!!x.match(mediaFileRegex) || !!x.match(drawingRelsRegex)))
    );

    const sharedStringsFile = files.find(f => f.path === stringsFilePath);
    // Updated to store structured content (rich text runs) or simple string
    const sharedStrings: (string | OfficeContentNode[])[] = [];

    if (sharedStringsFile) {
        const xml = parseXmlString(sharedStringsFile.content.toString());
        const siNodes = getElementsByTagName(xml, "si");
        for (const si of siNodes) {
            const runNodes = getElementsByTagName(si, "r");
            if (runNodes.length > 0) {
                // Rich text with runs
                const runs: OfficeContentNode[] = [];
                for (const run of runNodes) {
                    const tNode = getElementsByTagName(run, "t")[0];
                    if (tNode) {
                        const text = tNode.textContent || '';
                        // Extract run formatting 
                        const rPr = getElementsByTagName(run, "rPr")[0];
                        const formatting: TextFormatting = {};
                        if (rPr) {
                            if (getElementsByTagName(rPr, "b").length > 0) formatting.bold = true;
                            if (getElementsByTagName(rPr, "i").length > 0) formatting.italic = true;
                            if (getElementsByTagName(rPr, "u").length > 0) formatting.underline = true;
                            if (getElementsByTagName(rPr, "strike").length > 0) formatting.strikethrough = true;

                            const sz = getElementsByTagName(rPr, "sz")[0];
                            if (sz) formatting.size = sz.getAttribute("val") + 'pt';

                            const color = getElementsByTagName(rPr, "color")[0];
                            if (color) {
                                const rgb = color.getAttribute("rgb");
                                if (rgb) formatting.color = '#' + rgb.substring(2);
                            }

                            const rFont = getElementsByTagName(rPr, "rFont")[0];
                            if (rFont) formatting.font = rFont.getAttribute("val") || undefined;

                            const vertAlign = getElementsByTagName(rPr, "vertAlign")[0];
                            if (vertAlign) {
                                const val = vertAlign.getAttribute("val");
                                if (val === "subscript") formatting.subscript = true;
                                if (val === "superscript") formatting.superscript = true;
                            }
                        }
                        runs.push({
                            type: 'text',
                            text: text,
                            formatting: Object.keys(formatting).length > 0 ? formatting : undefined
                        });
                    }
                }
                sharedStrings.push(runs);
            } else {
                // Simple text case
                const tNodes = getElementsByTagName(si, "t");
                let text = '';
                for (const t of tNodes) {
                    text += t.textContent || '';
                }
                sharedStrings.push(text);
            }
        }
    }

    // Parse styles to build formatting map
    const stylesFile = files.find(f => f.path === 'xl/styles.xml');
    const cellFormatMap: Record<number, TextFormatting> = {};

    if (stylesFile) {
        const xml = parseXmlString(stylesFile.content.toString());

        // Parse fonts
        const fontsNode = getElementsByTagName(xml, "fonts")[0];
        const fonts: TextFormatting[] = [];
        if (fontsNode) {
            const fontNodes = getElementsByTagName(fontsNode, "font");
            for (const font of fontNodes) {
                const formatting: TextFormatting = {};
                if (getElementsByTagName(font, "b").length > 0) formatting.bold = true;
                if (getElementsByTagName(font, "i").length > 0) formatting.italic = true;
                if (getElementsByTagName(font, "u").length > 0) formatting.underline = true;
                if (getElementsByTagName(font, "strike").length > 0) formatting.strikethrough = true;

                const szNode = getElementsByTagName(font, "sz")[0];
                if (szNode) {
                    const val = szNode.getAttribute("val");
                    if (val) formatting.size = val + 'pt';
                }

                const colorNode = getElementsByTagName(font, "color")[0];
                if (colorNode) {
                    const rgb = colorNode.getAttribute("rgb");
                    if (rgb) formatting.color = '#' + rgb.substring(2); // Remove alpha channel
                }

                const nameNode = getElementsByTagName(font, "name")[0];
                if (nameNode) {
                    const val = nameNode.getAttribute("val");
                    if (val) formatting.font = val;
                }

                const vertAlignNode = getElementsByTagName(font, "vertAlign")[0];
                if (vertAlignNode) {
                    const val = vertAlignNode.getAttribute("val");
                    if (val === "subscript") formatting.subscript = true;
                    if (val === "superscript") formatting.superscript = true;
                }

                fonts.push(formatting);
            }
        }

        // Parse fills (for background color)
        const fillsNode = getElementsByTagName(xml, "fills")[0];
        const fills: TextFormatting[] = [];
        if (fillsNode) {
            const fillNodes = getElementsByTagName(fillsNode, "fill");
            for (const fill of fillNodes) {
                const formatting: TextFormatting = {};
                const patternFill = getElementsByTagName(fill, "patternFill")[0];
                if (patternFill) {
                    const fgColor = getElementsByTagName(patternFill, "fgColor")[0];
                    if (fgColor) {
                        const rgb = fgColor.getAttribute("rgb");
                        const theme = fgColor.getAttribute("theme");

                        if (rgb && rgb !== "00000000") { // Not default/auto
                            formatting.backgroundColor = '#' + rgb.substring(2);
                        } else if (theme) {
                            // Basic mapping for standard Office themes (Dark 1, Light 1, Dark 2, Light 2)
                            // 0: Light 1 (White), 1: Dark 1 (Black), 2: Light 2 (Tan/Gray), 3: Dark 2 (Blue/Grey)
                            const themeIdx = parseInt(theme);
                            if (themeIdx === 0) formatting.backgroundColor = '#FFFFFF';
                            else if (themeIdx === 1) formatting.backgroundColor = '#000000';
                            else if (themeIdx === 2) formatting.backgroundColor = '#EEECE1'; // Standard Light 2
                            else if (themeIdx === 3) formatting.backgroundColor = '#1F497D'; // Standard Dark 2
                        }
                    }
                }
                fills.push(formatting);
            }
        }

        // Parse cellXfs (cell format definitions)
        const cellXfsNode = getElementsByTagName(xml, "cellXfs")[0];
        if (cellXfsNode) {
            const xfNodes = getElementsByTagName(cellXfsNode, "xf");
            for (let i = 0; i < xfNodes.length; i++) {
                const xf = xfNodes[i];
                const formatting: TextFormatting = {};

                const fontId = xf.getAttribute("fontId");
                if (fontId) {
                    const fontIdx = parseInt(fontId);
                    if (fonts[fontIdx]) {
                        Object.assign(formatting, fonts[fontIdx]);
                    }
                }

                const fillId = xf.getAttribute("fillId");
                if (fillId) {
                    const fillIdx = parseInt(fillId);
                    if (fills[fillIdx] && fills[fillIdx].backgroundColor) {
                        formatting.backgroundColor = fills[fillIdx].backgroundColor;
                    }
                }

                const alignmentNode = getElementsByTagName(xf, "alignment")[0];
                if (alignmentNode) {
                    const horizontal = alignmentNode.getAttribute("horizontal");
                    if (horizontal === 'center' || horizontal === 'right' || horizontal === 'justify' || horizontal === 'left') {
                        formatting.alignment = horizontal;
                    }
                }

                cellFormatMap[i] = formatting;
            }
        }
    }

    const attachments: OfficeAttachment[] = [];
    const mediaFiles = files.filter(f => f.path.match(/xl\/media\/.*/));
    const chartFiles = files.filter(f => f.path.match(chartsRegex));

    // Map to store image details by drawing file path and relationship ID
    const drawingImageMap: Record<string, Record<string, { path: string, altText?: string }>> = {};

    if (config.extractAttachments) {
        // 1. Parse Drawing Rels to map rIds to media paths
        const drawingRelsFiles = files.filter(f => f.path.match(drawingRelsRegex));
        for (const relFile of drawingRelsFiles) {
            const drawingFilename = relFile.path.split('/').pop()?.replace('.rels', '') || '';
            const drawingPath = `xl/drawings/${drawingFilename}`;

            const relsXml = parseXmlString(relFile.content.toString());
            const relationships = getElementsByTagName(relsXml, "Relationship");

            if (!drawingImageMap[drawingPath]) {
                drawingImageMap[drawingPath] = {};
            }

            for (const rel of relationships) {
                const id = rel.getAttribute("Id");
                const target = rel.getAttribute("Target");
                if (id && target && target.includes('media/')) {
                    // Target is usually like "../media/image1.png"
                    const mediaPath = 'xl/' + target.replace('../', '');
                    drawingImageMap[drawingPath][id] = { path: mediaPath };
                }
            }
        }

        // 2. Parse Drawings to get Alt Text and link to Rels
        const drawingFiles = files.filter(f => f.path.match(drawingsRegex));
        for (const drawingFile of drawingFiles) {
            const xml = parseXmlString(drawingFile.content.toString());
            const pics = getElementsByTagName(xml, "xdr:pic"); // SpreadsheetML drawing

            const rels = drawingImageMap[drawingFile.path] || {};

            for (const pic of pics) {
                const blipFill = getElementsByTagName(pic, "xdr:blipFill")[0];
                const blip = blipFill ? getElementsByTagName(blipFill, "a:blip")[0] : null;
                const embedId = blip ? blip.getAttribute("r:embed") : null;

                const nvPicPr = getElementsByTagName(pic, "xdr:nvPicPr")[0];
                const cNvPr = nvPicPr ? getElementsByTagName(nvPicPr, "xdr:cNvPr")[0] : null;
                const altText = cNvPr ? (cNvPr.getAttribute("descr") || cNvPr.getAttribute("name")) : undefined;

                if (embedId && rels[embedId]) {
                    rels[embedId].altText = altText || '';
                }
            }
        }

        // 3. Process Media Files
        for (const media of mediaFiles) {
            const attachment = createAttachment(media.path.split('/').pop() || 'image', media.content);

            // Try to find alt text for this media
            let altText = '';
            for (const drawingPath in drawingImageMap) {
                for (const rId in drawingImageMap[drawingPath]) {
                    if (drawingImageMap[drawingPath][rId].path === media.path) {
                        altText = drawingImageMap[drawingPath][rId].altText || '';
                        break;
                    }
                }
                if (altText) break;
            }
            if (altText) attachment.altText = altText;

            attachments.push(attachment);

            if (config.ocr) {
                if (attachment.mimeType.startsWith('image/')) {
                    try {
                        const ocrText = await performOcr(media.content, config.ocrLanguage);
                        if (ocrText.trim()) {
                            attachment.ocrText = ocrText.trim();
                        }
                    } catch (e) {
                        logWarning(`OCR failed for ${attachment.name}:`, config, e);
                    }
                }
            }
        }

        for (const chart of chartFiles) {
            const attachment: OfficeAttachment = {
                type: 'chart',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                data: chart.content.toString('base64'),
                name: chart.path.split('/').pop() || '',
                extension: 'xml'
            };

            // Extract structured chart data
            try {
                const chartData = extractChartData(chart.content);
                attachment.chartData = chartData;
            } catch (e) {
                logWarning(`Failed to extract chart data from ${chart.path}:`, config, e);
            }

            attachments.push(attachment);
        }
    }

    // Build map of drawing rId -> chart attachment name for linking
    const drawingChartMap: Record<string, Record<string, string>> = {};
    if (config.extractAttachments) {
        const drawingRelsFiles = files.filter(f => f.path.match(drawingRelsRegex));
        for (const relFile of drawingRelsFiles) {
            const drawingFilename = relFile.path.split('/').pop()?.replace('.rels', '') || '';
            const drawingPath = `xl/drawings/${drawingFilename}`;

            const relsXml = parseXmlString(relFile.content.toString());
            const relationships = getElementsByTagName(relsXml, "Relationship");

            if (!drawingChartMap[drawingPath]) {
                drawingChartMap[drawingPath] = {};
            }

            for (const rel of relationships) {
                const id = rel.getAttribute("Id");
                const target = rel.getAttribute("Target");
                const type = rel.getAttribute("Type");
                if (id && target && type && type.includes('chart')) {
                    // Target is like "../charts/chart1.xml"
                    const chartName = target.split('/').pop() || '';
                    drawingChartMap[drawingPath][id] = chartName;
                }
            }
        }
    }

    // Parse workbook.xml to get sheet names and map them to sheet files
    const sheetNameMap: Record<string, string> = {};
    const workbookFile = files.find(f => f.path === 'xl/workbook.xml');
    const workbookRelsFile = files.find(f => f.path === 'xl/_rels/workbook.xml.rels');

    if (workbookFile && workbookRelsFile) {
        // Parse rels to get rId -> file mapping
        const relsXml = parseXmlString(workbookRelsFile.content.toString());
        const relationships = getElementsByTagName(relsXml, "Relationship");
        const rIdToFile: Record<string, string> = {};

        for (const rel of relationships) {
            const rId = rel.getAttribute("Id");
            const target = rel.getAttribute("Target");
            if (rId && target) {
                // Target is like "worksheets/sheet1.xml"
                const filename = target.split('/').pop() || '';
                rIdToFile[rId] = filename;
            }
        }

        // Parse workbook.xml to get sheet name -> rId mapping
        const workbookXml = parseXmlString(workbookFile.content.toString());
        const sheets = getElementsByTagName(workbookXml, "sheet");

        for (const sheet of sheets) {
            const name = sheet.getAttribute("name");
            const rId = sheet.getAttribute("r:id");
            if (name && rId && rIdToFile[rId]) {
                sheetNameMap[rIdToFile[rId]] = name;
            }
        }
    }

    const content: OfficeContentNode[] = [];
    const rawContents: string[] = [];

    for (const file of files) {
        if (file.path.match(mediaFileRegex)) continue;
        if (file.path === stringsFilePath) continue;
        if (file.path === 'xl/styles.xml') continue;
        if (file.path.match(drawingsRegex)) continue;
        if (file.path.match(chartsRegex)) continue;
        if (file.path.match(relsRegex)) continue;
        if (file.path.match(drawingRelsRegex)) continue;
        if (file.path.match(corePropsFileRegex)) continue;
        if (file.path === 'docProps/app.xml') continue;

        if (file.path.match(sheetsRegex)) {
            if (config.includeRawContent) {
                rawContents.push(file.content.toString());
            }

            // Parse merged cells for this sheet
            const sheetXmlStr = file.content.toString();
            const mergeMap: Record<string, { rowSpan: number, colSpan: number }> = {};
            const mergeCellRegex = /<mergeCell\s+ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\s*\/>/g;
            let mergeCellMatch;
            while ((mergeCellMatch = mergeCellRegex.exec(sheetXmlStr)) !== null) {
                const startCol = mergeCellMatch[1];
                const startRow = parseInt(mergeCellMatch[2]);
                const endCol = mergeCellMatch[3];
                const endRow = parseInt(mergeCellMatch[4]);
                const colSpan = colLetterToIndex(endCol) - colLetterToIndex(startCol) + 1;
                const rowSpan = endRow - startRow + 1;
                const key = startCol + startRow;
                mergeMap[key] = { rowSpan, colSpan };
            }

            const rows: OfficeContentNode[] = [];
            const rowRegex = /<row.*?>[\s\S]*?<\/row>/g;
            const rowMatches = sheetXmlStr.match(rowRegex);

            if (rowMatches) {
                for (const rowXml of rowMatches) {
                    const cells: OfficeContentNode[] = [];
                    const cRegex = /<c.*?>[\s\S]*?<\/c>/g;
                    const cMatches = rowXml.match(cRegex);

                    const rMatch = rowXml.match(/r="(\d+)"/);
                    const rowIndex = rMatch ? parseInt(rMatch[1]) - 1 : 0;

                    if (cMatches) {
                        for (const cXml of cMatches) {
                            // Extract cell value
                            const typeMatch = cXml.match(/t="([a-z]+)"/);
                            const type = typeMatch ? typeMatch[1] : 'n'; // n = number (default)

                            const vMatch = cXml.match(/<v>(.*?)<\/v>/);
                            const tMatch = cXml.match(/<t>(.*?)<\/t>/);

                            let text = '';
                            let cellNodes: OfficeContentNode[] = [];

                            if (type === 's' && vMatch) {
                                const idx = parseInt(vMatch[1]);
                                const content = sharedStrings[idx];
                                if (Array.isArray(content)) {
                                    // Rich text runs
                                    // Deep copy runs to avoid reference issues if reused
                                    cellNodes = JSON.parse(JSON.stringify(content));
                                    text = cellNodes.map(n => n.text).join('');
                                } else {
                                    text = content || '';
                                }
                            } else if (type === 'inlineStr' && tMatch) {
                                text = tMatch[1];
                            } else if (vMatch) {
                                text = vMatch[1];
                            }

                            // Parse cell coordinate
                            const coordMatch = cXml.match(/r="([A-Z]+)(\d+)"/);
                            const colStr = coordMatch ? coordMatch[1] : '';
                            const colIndex = colStr ? colLetterToIndex(colStr) : 0;

                            if (text || cellNodes.length > 0) {
                                // Extract cell style index
                                const styleMatch = cXml.match(/s="(\d+)"/);
                                const styleIdx = styleMatch ? parseInt(styleMatch[1]) : undefined;
                                const cellFormatting = (styleIdx !== undefined && cellFormatMap[styleIdx]) ? cellFormatMap[styleIdx] : {};

                                if (cellNodes.length > 0) {
                                    // If we have specific runs, merge cell styles into them if run style is missing
                                    // But usually run style overrides cell style (except maybe background)
                                    for (const node of cellNodes) {
                                        if (!node.formatting) node.formatting = {};
                                        // Cell background always applies
                                        if (cellFormatting.backgroundColor) node.formatting.backgroundColor = cellFormatting.backgroundColor;
                                        // Cell alignment always applies
                                        if (cellFormatting.alignment) node.formatting.alignment = cellFormatting.alignment;

                                        // Font defaults from cell style if not in run
                                        if (!node.formatting.font && cellFormatting.font) node.formatting.font = cellFormatting.font;
                                        if (!node.formatting.size && cellFormatting.size) node.formatting.size = cellFormatting.size;
                                    }
                                } else {
                                    // Simple text node
                                    cellNodes.push({
                                        type: 'text',
                                        text: text,
                                        formatting: cellFormatting
                                    });
                                }

                                const cellCoord = colStr + (coordMatch ? coordMatch[2] : '');
                                const cellMeta: CellMetadata = { row: rowIndex, col: colIndex };
                                if (cellCoord && mergeMap[cellCoord]) {
                                    const merge = mergeMap[cellCoord];
                                    if (merge.colSpan > 1) cellMeta.colSpan = merge.colSpan;
                                    if (merge.rowSpan > 1) cellMeta.rowSpan = merge.rowSpan;
                                }
                                const cellNode: OfficeContentNode = {
                                    type: 'cell',
                                    text: text,
                                    children: cellNodes,
                                    metadata: cellMeta
                                };
                                if (config.includeRawContent) {
                                    cellNode.rawContent = cXml;
                                }
                                cells.push(cellNode);
                            }
                        }
                    }

                    if (cells.length > 0) {
                        const rowNode: OfficeContentNode = {
                            type: 'row',
                            children: cells,
                            metadata: undefined
                        };
                        if (config.includeRawContent) {
                            rowNode.rawContent = rowXml;
                        }
                        rows.push(rowNode);
                    }
                }
            }

            // Parse and apply hyperlinks from the worksheet XML
            const sheetXml = parseXmlString(sheetXmlStr);
            const hyperlinkNodes = getElementsByTagName(sheetXml, "hyperlink");

            if (hyperlinkNodes.length > 0) {
                // Get sheet rels for resolving rIds to external URLs
                const sheetFilenameHL = file.path.split('/').pop() || '';
                const relsFilenameHL = `xl/worksheets/_rels/${sheetFilenameHL}.rels`;
                const relsFileHL = files.find(f => f.path === relsFilenameHL);
                const sheetRels: Record<string, string> = {};

                if (relsFileHL) {
                    const relsXml = parseXmlString(relsFileHL.content.toString());
                    const rels = getElementsByTagName(relsXml, "Relationship");
                    for (const rel of rels) {
                        const id = rel.getAttribute("Id");
                        const target = rel.getAttribute("Target");
                        if (id && target) sheetRels[id] = target;
                    }
                }

                // Build hyperlink map keyed by "row,col" (0-based)
                const hlMap: Record<string, { link: string, linkType: 'internal' | 'external' }> = {};

                for (const hlNode of hyperlinkNodes) {
                    const ref = hlNode.getAttribute("ref");
                    if (!ref) continue;

                    const rId = hlNode.getAttribute("r:id");
                    const location = hlNode.getAttribute("location");

                    let link: string | undefined;
                    let linkType: 'internal' | 'external' | undefined;

                    if (rId && sheetRels[rId]) {
                        link = sheetRels[rId];
                        linkType = 'external';
                    } else if (location) {
                        link = location;
                        linkType = 'internal';
                    }

                    if (link && linkType) {
                        // Convert ref like "A1" to row,col (0-based)
                        const refMatch = ref.match(/^([A-Z]+)(\d+)$/);
                        if (refMatch) {
                            const colIdx = colLetterToIndex(refMatch[1]);
                            const rowIdx = parseInt(refMatch[2]) - 1;
                            hlMap[`${rowIdx},${colIdx}`] = { link, linkType };
                        }
                    }
                }

                // Apply hyperlinks to existing cell text nodes
                for (const rowNode of rows) {
                    if (rowNode.children) {
                        for (const cellNode of rowNode.children) {
                            if (cellNode.type === 'cell' && cellNode.metadata) {
                                const meta = cellNode.metadata as CellMetadata;
                                const key = `${meta.row},${meta.col}`;
                                if (hlMap[key] && cellNode.children) {
                                    for (const child of cellNode.children) {
                                        if (child.type === 'text') {
                                            child.metadata = {
                                                ...(child.metadata || {}),
                                                link: hlMap[key].link,
                                                linkType: hlMap[key].linkType
                                            } as TextMetadata;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Handle Drawings in Sheet (images and charts)
            if (config.extractAttachments) {
                // Parse Sheet Rels to map drawing rIds
                const sheetFilename = file.path.split('/').pop() || '';
                const relsFilename = `xl/worksheets/_rels/${sheetFilename}.rels`;
                const relsFile = files.find(f => f.path === relsFilename);

                const drawingMap: Record<string, string> = {}; // rId -> drawingPath

                if (relsFile) {
                    const relsXml = parseXmlString(relsFile.content.toString());
                    const relationships = getElementsByTagName(relsXml, "Relationship");
                    for (const rel of relationships) {
                        const id = rel.getAttribute("Id");
                        const target = rel.getAttribute("Target");
                        const type = rel.getAttribute("Type");
                        if (id && target && type && type.includes('drawing')) {
                            drawingMap[id] = 'xl/drawings/' + target.replace('../drawings/', '');
                        }
                    }
                }

                const drawingMatches = file.content.toString().match(/<drawing r:id="(.*?)"/g);
                if (drawingMatches) {
                    for (const match of drawingMatches) {
                        const rIdMatch = match.match(/r:id="(.*?)"/);
                        const rId = rIdMatch ? rIdMatch[1] : null;

                        if (rId && drawingMap[rId]) {
                            const drawingPath = drawingMap[rId];

                            // Find all images in this drawing
                            const images = drawingImageMap[drawingPath];
                            if (images) {
                                for (const imgId in images) {
                                    const imgInfo = images[imgId];
                                    const attachment = attachments.find(a => a.name === imgInfo.path.split('/').pop());
                                    if (attachment) {
                                        const imageNode: OfficeContentNode = {
                                            type: 'image',
                                            text: '', // Will be populated by assignAttachmentData
                                            children: [],
                                            metadata: {
                                                attachmentName: attachment.name || 'unknown',
                                                altText: imgInfo.altText || undefined
                                            } as ImageMetadata
                                        };
                                        rows.push(imageNode);
                                    }
                                }
                            }

                            // Find all charts in this drawing
                            const charts = drawingChartMap[drawingPath];
                            if (charts) {
                                for (const chartRId in charts) {
                                    const chartName = charts[chartRId];
                                    const attachment = attachments.find(a => a.name === chartName);
                                    if (attachment) {
                                        const chartNode: OfficeContentNode = {
                                            type: 'chart',
                                            text: '', // Will be populated by assignAttachmentData
                                            children: [],
                                            metadata: {
                                                attachmentName: chartName
                                            } as ChartMetadata
                                        };
                                        rows.push(chartNode);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Get proper sheet name from workbook.xml mapping, fallback to filename
            const sheetFileName = file.path.split('/').pop() || 'Sheet';
            const sheetName = sheetNameMap[sheetFileName] || sheetFileName;

            content.push({
                type: 'sheet',
                children: rows,
                metadata: { sheetName },
                rawContent: config.includeRawContent ? file.content.toString() : undefined
            });
        }
    }

    const corePropsFile = files.find(f => f.path.match(corePropsFileRegex));
    const metadata = corePropsFile ? parseOfficeMetadata(corePropsFile.content.toString()) : {};

    const appPropsFile = files.find(f => f.path === 'docProps/app.xml');
    const appMetadata = appPropsFile ? parseAppMetadata(appPropsFile.content.toString()) : {};

    // Link OCR text and chart data to content nodes (like PPTX parser)
    const assignAttachmentData = (nodes: OfficeContentNode[]) => {
        for (const node of nodes) {
            if ('attachmentName' in (node.metadata || {})) {
                const meta = node.metadata as ImageMetadata | ChartMetadata;
                const attachment = attachments.find(a => a.name === meta.attachmentName);
                if (attachment) {
                    if (node.type === 'image') {
                        // Link OCR text to image node
                        if (attachment.ocrText) {
                            node.text = attachment.ocrText;
                        }
                        // Copy altText to attachment
                        if ((meta as ImageMetadata).altText) {
                            attachment.altText = (meta as ImageMetadata).altText;
                        }
                    }
                    if (node.type === 'chart') {
                        // Link chart data text to chart node
                        if (attachment.chartData) {
                            node.text = attachment.chartData.rawTexts.join(config.newlineDelimiter || '\n');
                        }
                    }
                }
            }
            if (node.children) {
                assignAttachmentData(node.children);
            }
        }
    };
    assignAttachmentData(content);

    return {
        type: 'xlsx',
        metadata: { ...metadata, ...appMetadata },
        content: content,
        attachments: attachments,
        toText: () => content.map(c => {
            // Recursive text extraction
            const getText = (node: OfficeContentNode): string => {
                let t = '';
                if (node.children) {
                    t += node.children.map(getText).filter(t => t != '').join(!node.children[0]?.children ? '' : config.newlineDelimiter ?? '\n');
                }
                else
                    t += node.text || '';
                return t;
            };
            return getText(c);
        }).filter(t => t != '').join(config.newlineDelimiter ?? '\n'),
        toMarkdown: () => astToMarkdown(content, config)
    };
};
