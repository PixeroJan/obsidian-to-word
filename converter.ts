import {
	Document,
	Packer,
	Paragraph,
	TextRun,
	HeadingLevel,
	AlignmentType,
	UnderlineType,
	Table,
	TableRow,
	TableCell,
	WidthType,
	BorderStyle,
	convertInchesToTwip,
	ExternalHyperlink,
	ImageRun,
	LevelFormat,
	AlignmentType as NumberAlignment,
	IStylesOptions,
	IParagraphStyleOptions,
	TabStopType,
	TableLayoutType,
} from 'docx';
import MarkdownIt from 'markdown-it';
import type Token from 'markdown-it/lib/token';
import { full as markdownItEmoji } from 'markdown-it-emoji';
import markdownItMark from 'markdown-it-mark';
import hljs from 'highlight.js';

interface ToWordSettings {
	defaultFontFamily: string;
	defaultFontSize: number;
	includeMetadata: boolean;
	preserveFormatting: boolean;
	useObsidianAppearance: boolean;
	includeFilenameAsHeader: boolean;
	pageSize: 'A4' | 'A5' | 'A3' | 'Letter' | 'Legal' | 'Tabloid';
}

interface TextStyle {
	bold?: boolean;
	italic?: boolean;
	strikethrough?: boolean;
	code?: boolean;
	highlight?: boolean;
	underline?: boolean;
	color?: string;
	superScript?: boolean;
	subScript?: boolean;
	backgroundColor?: string;
	codeBlock?: boolean;
}

interface ObsidianFontSettings {
	textFont: string;
	monospaceFont: string;
	baseFontSize: number;
	lineHeight: number;
	sizeMultiplier: number;
	headingSizes: number[];
	headingFonts: string[];
	headingColors: string[];
}

const CODE_SYNTAX_COLOR_MAP: Record<string, string> = {
	'hljs-keyword': '569CD6',
	'hljs-attr': '9CDCFE',
	'hljs-attribute': '9CDCFE',
	'hljs-symbol': 'C586C0',
	'hljs-built_in': '4EC9B0',
	'hljs-type': '4EC9B0',
	'hljs-literal': 'B5CEA8',
	'hljs-number': 'B5CEA8',
	'hljs-string': 'CE9178',
	'hljs-template-variable': '9CDCFE',
	'hljs-variable': '9CDCFE',
	'hljs-title': '4EC9B0',
	'hljs-function': 'DCDCAA',
	'hljs-comment': '6A9955',
	'hljs-meta': 'D4D4D4',
};

const DEFAULT_HYPERLINK_COLOR = '0563C1';

export class MarkdownToDocxConverter {
	private settings: ToWordSettings;
	private obsidianFonts: ObsidianFontSettings | null = null;
	private filename: string = '';
	private resourceLoader?: (link: string) => Promise<ArrayBuffer | null>;
	private footnoteDefinitions: Map<string, string> = new Map();
	private usedFootnotes: string[] = [];
	private md: MarkdownIt;

	constructor(settings: ToWordSettings) {
		this.settings = settings;
		this.md = new MarkdownIt({ html: true, linkify: false, typographer: false, breaks: false });
		this.md.use(markdownItEmoji);
		this.md.use(markdownItMark);
	}

	private createStyles(): IStylesOptions {
		if (!(this.settings.useObsidianAppearance && this.obsidianFonts)) {
			return {};
		}

		const paragraphStyles: IParagraphStyleOptions[] = [];
		const baseFont = this.obsidianFonts.textFont;
		const baseSize = this.obsidianFonts.baseFontSize * 2; // Half-points

		// Create heading styles that override Word defaults
		const headingLevels = ['Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5', 'Heading6'];
		for (let i = 0; i < 6; i++) {
			const fontSize = (this.obsidianFonts.headingSizes[i] || this.obsidianFonts.baseFontSize) * 2;
			const color = this.rgbToHex(this.obsidianFonts.headingColors[i]);
			const headingFontFamily = this.obsidianFonts.headingFonts[i] || baseFont;

			paragraphStyles.push({
				id: headingLevels[i],
				name: headingLevels[i],
				basedOn: 'Normal',
				next: 'Normal',
				run: {
					font: headingFontFamily,
					size: fontSize,
					bold: true,
					color: color,
				},
				paragraph: {
					spacing: this.getLineSpacing() || undefined,
				},
			});
		}

		// List paragraph style
		paragraphStyles.push({
			id: 'ListParagraph',
			name: 'List Paragraph',
			basedOn: 'Normal',
			run: {
				font: baseFont,
				size: baseSize,
			},
			paragraph: {
				spacing: this.getLineSpacing() || undefined,
				indent: {
					left: 0,
					hanging: 0,
				},
			},
		});

		let codeBlockFont = this.obsidianFonts.monospaceFont;
		
		if (!codeBlockFont || codeBlockFont.trim() === '' || codeBlockFont === 'undefined' || codeBlockFont === '??' || codeBlockFont.includes('??')) {
			codeBlockFont = this.getPlatformMonospaceFont();
		}
		
		// Final validation
		if (!codeBlockFont || codeBlockFont.trim() === '' || codeBlockFont === 'undefined' || codeBlockFont === '??' || codeBlockFont.includes('??')) {
			codeBlockFont = 'Courier New';
		}

		paragraphStyles.push({
			id: 'ObsidianCodeBlock',
			name: 'Obsidian Code Block',
			basedOn: 'Normal',
			run: {
				font: codeBlockFont,
				size: baseSize,
				bold: true,
			},
			paragraph: {
				spacing: this.getLineSpacing() || undefined,
			},
		});

		return {
			default: {
				document: {
					run: {
						font: baseFont,
						size: baseSize,
					},
					paragraph: {
						spacing: this.getLineSpacing() || undefined,
					},
				},
			},
			paragraphStyles,
		};
	}

	private getPageSize(): { width: number; height: number; orientation?: 'portrait' | 'landscape' } {
		// All dimensions in twips (1/20th of a point, 1440 twips = 1 inch)
		const sizes = {
			'A4': { width: convertInchesToTwip(8.27), height: convertInchesToTwip(11.69) }, // 210 x 297 mm
			'A5': { width: convertInchesToTwip(5.83), height: convertInchesToTwip(8.27) },  // 148 x 210 mm
			'A3': { width: convertInchesToTwip(11.69), height: convertInchesToTwip(16.54) }, // 297 x 420 mm
			'Letter': { width: convertInchesToTwip(8.5), height: convertInchesToTwip(11) },  // 8.5 x 11 inches
			'Legal': { width: convertInchesToTwip(8.5), height: convertInchesToTwip(14) },   // 8.5 x 14 inches
			'Tabloid': { width: convertInchesToTwip(11), height: convertInchesToTwip(17) },  // 11 x 17 inches
		};
		
		return sizes[this.settings.pageSize] || sizes['A4'];
	}

	private getPlatformMonospaceFont(): string {
		const platform = typeof navigator !== 'undefined' ? navigator.platform.toLowerCase() : '';
		if (platform.includes('mac')) {
			return 'SF Mono';
		} else if (platform.includes('win')) {
			return 'Courier New';
		} else {
			return 'Courier New';
		}
	}

	async convert(
		markdown: string,
		title: string,
		obsidianFonts?: ObsidianFontSettings | null,
		resourceLoader?: (link: string) => Promise<ArrayBuffer | null>,
	): Promise<Blob> {
		this.obsidianFonts = obsidianFonts || null;
		this.filename = title;
		this.resourceLoader = resourceLoader;

		const { content: cleanedMarkdown, definitions } = this.extractFootnotes(markdown);
		this.footnoteDefinitions = definitions;
		this.usedFootnotes = [];

		const paragraphs = await this.parseMarkdown(cleanedMarkdown);

		// Add filename as header if enabled
		if (this.settings.includeFilenameAsHeader) {
			paragraphs.unshift(this.createHeading(title, 1));
			paragraphs.unshift(new Paragraph({ children: [] })); // Add spacing
		}

		if (this.usedFootnotes.length > 0) {
			this.appendFootnotes(paragraphs);
		}

		const pageSize = this.getPageSize();
		const doc = new Document({
			styles: this.createStyles(),
			numbering: this.createNumbering(),
			sections: [{
				properties: {
					page: {
						size: {
							width: pageSize.width,
							height: pageSize.height,
						},
						margin: {
							top: convertInchesToTwip(1),
							right: convertInchesToTwip(1),
							bottom: convertInchesToTwip(1),
							left: convertInchesToTwip(1),
						},
					},
				},
				children: paragraphs,
			}],
		});

		this.resourceLoader = undefined;
		return await Packer.toBlob(doc);
	}

	private async parseMarkdown(markdown: string): Promise<(Paragraph | Table)[]> {
		const paragraphs: (Paragraph | Table)[] = [];
		const lines = markdown.split('\n');
		
		let i = 0;
		let inCodeBlock = false;
		let codeBlockContent: string[] = [];
		let codeBlockLanguage: string | null = null;
		let inTable = false;
		let tableRows: string[][] = [];
		let tableAlignments: string[] = [];

		while (i < lines.length) {
			const line = lines[i];

			// Handle code blocks
			const fenceMatch = line.trim().match(/^(```|~~~)(.*)$/);
			if (fenceMatch) {
				if (inCodeBlock) {
					// End of code block
					paragraphs.push(...this.createCodeBlock(codeBlockContent, codeBlockLanguage || undefined));
					codeBlockContent = [];
					codeBlockLanguage = null;
					inCodeBlock = false;
				} else {
					// Start of code block
					inCodeBlock = true;
					codeBlockLanguage = fenceMatch[2]?.trim() ? fenceMatch[2].trim().split(/\s+/)[0].toLowerCase() : null;
				}
				i++;
				continue;
			}

			if (inCodeBlock) {
				codeBlockContent.push(line);
				i++;
				continue;
			}

			const trimmedLine = line.trim();

			// Handle horizontal rules EARLY - before anything else can catch them
			// Explicitly check for ---, ***, ___ (with optional spaces)
			const isHorizontalRule = /^[-*_]{3,}$/.test(trimmedLine.replace(/\s/g, '')) && 
			                         /^[-*_\s]+$/.test(trimmedLine) &&
			                         trimmedLine.length >= 3;
			
			if (isHorizontalRule) {
				paragraphs.push(new Paragraph({
					border: {
						bottom: {
							color: "000000",
							space: 1,
							style: BorderStyle.SINGLE,
							size: 8,
						},
					},
					spacing: {
						before: 120,
						after: 120,
					},
					text: '',
				}));
				i++;
				continue;
			}

			// Handle tables
			if (line.trim().startsWith('|') && line.trim().endsWith('|')) {
				if (!inTable) {
					inTable = true;
					tableRows = [];
					tableAlignments = [];
				}

				const cells = line.split('|').slice(1, -1).map(cell => cell.trim());
				
				// Check if this is the alignment row
				if (cells.every(cell => /^:?-+:?$/.test(cell))) {
					tableAlignments = cells.map(cell => {
						if (cell.startsWith(':') && cell.endsWith(':')) return 'center';
						if (cell.endsWith(':')) return 'right';
						return 'left';
					});
				} else {
					tableRows.push(cells);
				}

				i++;
				continue;
			} else if (inTable) {
				// End of table
				paragraphs.push(...this.createTable(tableRows, tableAlignments));
				inTable = false;
				tableRows = [];
				tableAlignments = [];
			}

			// Handle collapsible <details> blocks
			if (/^<details/i.test(trimmedLine)) {
				const { paragraphs: detailsParagraphs, nextIndex } = await this.parseDetailsBlock(lines, i);
				paragraphs.push(...detailsParagraphs);
				i = nextIndex;
				continue;
			}

			// Handle raw HTML blocks that appear on their own line
			if (this.isHtmlBlockLine(trimmedLine)) {
				const htmlParagraphs = this.convertHtmlBlockToParagraphs(trimmedLine);
				if (htmlParagraphs.length > 0) {
					paragraphs.push(...htmlParagraphs);
				}
				i++;
				continue;
			}

			// Handle standard Markdown images ![alt](path "title")
			const standardImageMatch = trimmedLine.match(/^!\[([^\]]*)\]\((.+)\)$/);
			if (standardImageMatch) {
				const imageParagraph = await this.createStandardImageParagraph(standardImageMatch[1], standardImageMatch[2]);
				if (imageParagraph) {
					paragraphs.push(imageParagraph);
				} else {
					const fallbackLabel = standardImageMatch[1] || standardImageMatch[2];
					paragraphs.push(new Paragraph({ text: `[Image not found: ${fallbackLabel}]` }));
				}
				i++;
				continue;
			}

			// Handle Obsidian-style embedded images ![[image.png]]
			const wikiImageMatch = trimmedLine.match(/^!\[\[([^\]]+)\]\]$/);
			if (wikiImageMatch) {
				const imageParagraph = await this.createEmbeddedImageParagraph(wikiImageMatch[1]);
				if (imageParagraph) {
					paragraphs.push(imageParagraph);
				} else {
					paragraphs.push(new Paragraph({ text: `[Image not found: ${wikiImageMatch[1]}]` }));
				}
				i++;
				continue;
			}

			// Handle empty lines
			if (trimmedLine === '') {
				paragraphs.push(new Paragraph({ text: '' }));
				i++;
				continue;
			}

			// Handle headings
			const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
			if (headingMatch) {
				const level = headingMatch[1].length;
				const text = headingMatch[2];
				paragraphs.push(this.createHeading(text, level));
				i++;
				continue;
			}

			// Handle definition lists (term followed by : definition)
			if (trimmedLine !== '' && !trimmedLine.startsWith('#') && lines[i + 1] && /^:\s+/.test(lines[i + 1])) {
				const definitionMatch = lines[i + 1].match(/^:\s+(.+)$/);
				if (definitionMatch) {
					paragraphs.push(new Paragraph({
						children: [
							new TextRun({ text: `${trimmedLine}: `, bold: true }),
							...this.parseInlineFormatting(definitionMatch[1]),
						],
						spacing: this.getLineSpacing(),
					}));
					i += 2;
					continue;
				}
			}

			// Handle blockquotes (supports nesting)
			const blockquoteMatch = line.match(/^(\s*>+\s*)(.*)$/);
			if (blockquoteMatch) {
				const marker = blockquoteMatch[1];
				const depth = marker.replace(/[^>]/g, '').length;
				const quoteText = blockquoteMatch[2];

				paragraphs.push(new Paragraph({
					children: this.parseInlineFormatting(quoteText),
					indent: {
						left: convertInchesToTwip(0.3 * depth),
					},
					border: {
						left: {
							color: "CCCCCC",
							space: 1,
							style: BorderStyle.SINGLE,
							size: 12,
						},
					},
					spacing: this.getLineSpacing(),
				}));
				i++;
				continue;
			}

			// Handle GFM task lists
			const taskListMatch = line.match(/^(\s*)[-*+]\s+\[( |x|X)\]\s+(.*)$/);
			if (taskListMatch) {
				const indent = taskListMatch[1].length;
				const checked = taskListMatch[2].toLowerCase() === 'x';
				const taskText = taskListMatch[3];
				const level = Math.min(Math.floor(indent / 2), 2);
				const leftIndent = convertInchesToTwip(0.18 + level * 0.18);
				const hanging = convertInchesToTwip(0.18);

				// Use simpler checkbox symbols that render better in Word
				const checkboxSymbol = checked ? '☑' : '☐';
				
				paragraphs.push(new Paragraph({
					children: [
						new TextRun({ text: `${checkboxSymbol} ` }),
						...this.parseInlineFormatting(taskText),
					],
					style: this.settings.useObsidianAppearance ? 'ListParagraph' : undefined,
					indent: {
						left: leftIndent,
						hanging: hanging,
					},
					spacing: this.getLineSpacing(),
				}));
				i++;
				continue;
			}

			// Handle lists
			const unorderedListMatch = line.match(/^(\s*)([-*+])\s+(.+)$/);
			const orderedListMatch = line.match(/^(\s*)(\d+)\.\s+(.+)$/);
			
			if (unorderedListMatch) {
				const indent = unorderedListMatch[1].length;
				const text = unorderedListMatch[3];
				// Use 2 spaces per level to match Obsidian's visual indentation better
				const level = Math.min(Math.floor(indent / 2), 2);
				paragraphs.push(new Paragraph({
					children: this.parseInlineFormatting(text),
					numbering: {
						reference: 'obsidian-bullet',
						level: level,
					},
					style: this.settings.useObsidianAppearance ? 'ListParagraph' : undefined,
					spacing: this.getLineSpacing(),
				}));
				i++;
				continue;
			}

			if (orderedListMatch) {
				const indent = orderedListMatch[1].length;
				const text = orderedListMatch[3];
				// Use 2 spaces per level to match Obsidian's visual indentation better
				const level = Math.min(Math.floor(indent / 2), 2);
				paragraphs.push(new Paragraph({
					children: this.parseInlineFormatting(text),
					numbering: {
						reference: 'obsidian-numbered',
						level: level,
					},
					style: this.settings.useObsidianAppearance ? 'ListParagraph' : undefined,
					spacing: this.getLineSpacing(),
				}));
				i++;
				continue;
			}

			// Handle regular paragraphs
			paragraphs.push(new Paragraph({
				children: this.parseInlineFormatting(line),
				spacing: this.getLineSpacing(),
			}));

			i++;
		}

		// Handle any remaining table
		if (inTable) {
			paragraphs.push(...this.createTable(tableRows, tableAlignments));
		}

		return paragraphs;
	}

	private async parseDetailsBlock(lines: string[], startIndex: number): Promise<{ paragraphs: (Paragraph | Table)[]; nextIndex: number }> {
		let depth = 0;
		const collected: string[] = [];
		let i = startIndex;

		for (; i < lines.length; i++) {
			const current = lines[i];
			const trimmed = current.trim();
			if (/^<details/i.test(trimmed)) {
				depth++;
			}
			if (/^<\/details>/i.test(trimmed)) {
				depth--;
			}
			collected.push(current);
			if (depth === 0) {
				break;
			}
		}

		const block = collected.join('\n');
		const summaryMatch = block.match(/<summary[^>]*>([\s\S]*?)<\/summary>/i);
		const summaryText = summaryMatch ? summaryMatch[1].trim() : 'Details';
		const innerContent = block
			.replace(/<details[^>]*>/i, '')
			.replace(/<\/details>/i, '')
			.replace(summaryMatch ? summaryMatch[0] : '', '')
			.trim();

		const paragraphs: (Paragraph | Table)[] = [];
		
		// Add summary with visual indicator (using down arrow to show it's expanded)
		// Note: Word doesn't support interactive collapsing, so content is always visible
		paragraphs.push(new Paragraph({
			children: [
				new TextRun({ text: '▼ ', bold: true }),
				...this.parseInlineFormatting(summaryText),
			],
			spacing: this.getLineSpacing(),
			shading: {
				fill: 'E8E8E8',
			},
			border: {
				left: {
					color: '999999',
					space: 1,
					style: BorderStyle.SINGLE,
					size: 12,
				},
			},
		}));

		// Add inner content with indentation
		if (innerContent) {
			const innerMarkdown = innerContent.split('\n').map(line => '  ' + line).join('\n');
			const innerParagraphs = await this.parseMarkdown(innerMarkdown);
			paragraphs.push(...innerParagraphs);
		}

		// Add a blank line after the details block
		paragraphs.push(new Paragraph({ children: [] }));

		return { paragraphs, nextIndex: i + 1 };
	}

	private isHtmlBlockLine(line: string): boolean {
		if (!line.startsWith('<')) {
			return false;
		}
		if (/^<\/(\w+)/.test(line)) {
			return false;
		}
		if (/^<(!|\?)/.test(line)) {
			return false;
		}
		const tagMatch = line.match(/^<([a-zA-Z][\w:-]*)\b/);
		if (!tagMatch) {
			return false;
		}
		const tag = tagMatch[1].toLowerCase();
		if (tag === 'details' || tag === 'summary') {
			return false;
		}
		return true;
	}

	private convertHtmlBlockToParagraphs(html: string): Paragraph[] {
		const parser = new DOMParser();
		const doc = parser.parseFromString(`<wrapper>${html}</wrapper>`, 'text/html');
		const nodes = Array.from(doc.body.childNodes);
		const paragraphs = nodes.flatMap(node => this.buildParagraphsFromHtmlNode(node));
		if (paragraphs.length > 0) {
			return paragraphs;
		}
		const fallback = (doc.body.textContent ?? '').trim();
		if (!fallback) {
			return [];
		}
		return [new Paragraph({
			children: this.createTextRunsFromString(fallback, {}, true),
			spacing: this.getLineSpacing(),
		})];
	}

	private buildParagraphsFromHtmlNode(node: ChildNode): Paragraph[] {
		if (node.nodeType === Node.TEXT_NODE) {
			const text = node.textContent ?? '';
			if (!text.trim()) {
				return [];
			}
			return [new Paragraph({
				children: this.createTextRunsFromString(text.trim(), {}, true),
				spacing: this.getLineSpacing(),
			})];
		}

		if (node.nodeType !== Node.ELEMENT_NODE) {
			return [];
		}

		const element = node as HTMLElement;
		const tag = element.tagName.toLowerCase();

		if (tag === 'br') {
			return [new Paragraph({ children: [], spacing: this.getLineSpacing() })];
		}

		if (tag === 'hr') {
			return [new Paragraph({
				border: {
					bottom: {
						color: 'auto',
						space: 1,
						style: BorderStyle.SINGLE,
						size: 6,
					},
				},
				text: '',
			})];
		}

		if (tag === 'ul' || tag === 'ol') {
			const isOrdered = tag === 'ol';
			const items = Array.from(element.children).filter(child => child.tagName.toLowerCase() === 'li');
			return items.map(item => {
				const runs = this.parseHtmlNodes(item.childNodes, {}, true);
				return new Paragraph({
					children: runs.length ? runs : [this.createTextRun('', {})],
					numbering: {
						reference: isOrdered ? 'obsidian-numbered' : 'obsidian-bullet',
						level: 0,
					},
					spacing: this.getLineSpacing(),
				});
			});
		}

		if (this.isBlockElement(tag)) {
			const childParagraphs = Array.from(element.childNodes).flatMap(child => this.buildParagraphsFromHtmlNode(child));
			if (childParagraphs.length > 0) {
				return childParagraphs;
			}
		}

		const runs = this.parseHtmlNodes(element.childNodes, {}, true);
		if (runs.length === 0) {
			const fallbackText = (element.textContent ?? '').trim();
			if (!fallbackText) {
				return [];
			}
			return [new Paragraph({
				children: this.createTextRunsFromString(fallbackText, {}, true),
				spacing: this.getLineSpacing(),
			})];
		}

		return [new Paragraph({
			children: runs,
			spacing: this.getLineSpacing(),
		})];
	}

	private isBlockElement(tag: string): boolean {
		return [
			'address',
			'article',
			'aside',
			'blockquote',
			'div',
			'figcaption',
			'figure',
			'footer',
			'header',
			'li',
			'main',
			'nav',
			'p',
			'section',
			'summary',
		].includes(tag);
	}

	private getLineSpacing() {
		if (this.settings.useObsidianAppearance && this.obsidianFonts) {
			// Convert line height ratio to Word spacing
			// Word uses 240 twips per line by default (single spacing)
			// Multiply by line height ratio
			return {
				line: Math.round(240 * this.obsidianFonts.lineHeight),
				lineRule: "auto" as const,
			};
		}
		return undefined;
	}

	private rgbToHex(rgb: string): string | undefined {
		// Convert rgb(r, g, b) or rgba(r, g, b, a) to hex
		if (rgb === 'inherit' || !rgb) return undefined;
		
		const match = rgb.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
		if (!match) return undefined;
		
		const r = parseInt(match[1]);
		const g = parseInt(match[2]);
		const b = parseInt(match[3]);
		
		return ((r << 16) | (g << 8) | b).toString(16).padStart(6, '0').toUpperCase();
	}

	private createHeading(text: string, level: number): Paragraph {
		const headingLevels = [
			HeadingLevel.HEADING_1,
			HeadingLevel.HEADING_2,
			HeadingLevel.HEADING_3,
			HeadingLevel.HEADING_4,
			HeadingLevel.HEADING_5,
			HeadingLevel.HEADING_6,
		];

		let fontSize: number;
		let fontFamily: string;
		let color: string | undefined;
		const useAppearance = this.settings.useObsidianAppearance && this.obsidianFonts;

		if (useAppearance) {
			// Use actual Obsidian heading sizes, fonts, and colors
			fontSize = this.obsidianFonts!.headingSizes[level - 1] || this.obsidianFonts!.baseFontSize;
			fontFamily = this.obsidianFonts!.headingFonts[level - 1] || this.obsidianFonts!.textFont;
			const rgbColor = this.obsidianFonts!.headingColors[level - 1];
			color = this.rgbToHex(rgbColor);
			
			// If color is too light or undefined, use black for better visibility in Word
			if (!color || color === 'FFFFFF' || color === 'inherit') {
				color = '000000';
			}
		} else {
			// Standard Word heading sizes (more conservative)
			const standardHeadingSizes = [
				16, // H1: 16pt
				14, // H2: 14pt
				13, // H3: 13pt
				12, // H4: 12pt
				11, // H5: 11pt
				11, // H6: 11pt
			];
			fontSize = standardHeadingSizes[level - 1] || standardHeadingSizes[0];
			fontFamily = this.settings.defaultFontFamily;
			color = '000000'; // Black for standard headings
		}

		return new Paragraph({
			heading: headingLevels[level - 1] || HeadingLevel.HEADING_1,
			style: useAppearance ? `Heading${level}` : undefined,
			spacing: this.getLineSpacing(),
			children: [
				new TextRun({
					text,
					font: fontFamily,
					size: fontSize * 2,
					bold: true,
					color,
				}),
			],
		});
	}

	private createCodeBlock(lines: string[], language?: string): Paragraph[] {
		let bodyFontSize: number;
		let monospaceFont: string;

		if (this.settings.useObsidianAppearance && this.obsidianFonts) {
			bodyFontSize = this.obsidianFonts.baseFontSize;
			monospaceFont = this.obsidianFonts.monospaceFont;
			
			// Platform-specific fallback
			if (!monospaceFont || monospaceFont.trim() === '' || monospaceFont === 'undefined' || monospaceFont === '??' || monospaceFont.includes('??')) {
				monospaceFont = this.getPlatformMonospaceFont();
			}
		} else {
			bodyFontSize = this.settings.defaultFontSize;
			// Platform-specific monospace font
			monospaceFont = this.getPlatformMonospaceFont();
		}
		
		// Final validation
		if (!monospaceFont || monospaceFont.trim() === '' || monospaceFont === 'undefined' || monospaceFont === '??' || monospaceFont.includes('??')) {
			monospaceFont = 'Courier New';
		}

		const code = lines.join('\n');
		let highlightedLines: string[] | null = null;
		if (language && hljs.getLanguage(language)) {
			try {
				const result = hljs.highlight(code, { language });
				highlightedLines = result.value.split('\n');
			} catch {
				highlightedLines = null;
			}
		}

		const plainLines = code.split('\n');
		const effectiveLines = highlightedLines ?? plainLines;

		return effectiveLines.map((content, index) => {
			const originalText = plainLines[index] ?? '';
			const runs = highlightedLines
				? this.parseHighlightedHtmlLine(content)
				: [this.createTextRun(originalText, { code: true, codeBlock: true })];
			return new Paragraph({
				children: runs.length ? runs : [this.createTextRun('', { code: true })],
				shading: {
					fill: 'F5F5F5',
				},
				style: this.settings.useObsidianAppearance ? 'ObsidianCodeBlock' : undefined,
				spacing: this.getLineSpacing(),
			});
		});
	}

	private createPlainTextRuns(text: string): TextRun[] {
		const hasHardBreak = text.endsWith('  ');
		const content = hasHardBreak ? text.slice(0, -2) : text;
		const runs: TextRun[] = [];
		if (content.length > 0) {
			runs.push(...this.createTextRunsFromString(content, {}, true).filter((run): run is TextRun => run instanceof TextRun));
		}
		if (hasHardBreak) {
			runs.push(new TextRun({ text: '', break: 1 }));
		}
		return runs.length > 0 ? runs : [this.createTextRun('', {})];
	}

	private tokensToRuns(tokens: Token[], baseStyle: TextStyle, allowFootnotes: boolean, allowHyperlinks: boolean): (TextRun | ExternalHyperlink)[] {
		const runs: (TextRun | ExternalHyperlink)[] = [];

		for (let i = 0; i < tokens.length; i++) {
			const token = tokens[i];
			switch (token.type) {
				case 'text': {
					runs.push(...this.createTextRunsFromString(token.content, baseStyle, allowFootnotes));
					break;
				}
				case 'softbreak':
				case 'hardbreak': {
					runs.push(new TextRun({ text: '', break: 1 }));
					break;
				}
				case 'code_inline': {
					runs.push(this.createTextRun(token.content, { ...baseStyle, code: true }));
					break;
				}
				case 'html_inline': {
					runs.push(...this.parseHtmlStringToRuns(token.content, baseStyle, allowFootnotes));
					break;
				}
				case 'emoji': {
					runs.push(this.createTextRun(token.content, baseStyle));
					break;
				}
				case 'footnote_ref': {
					const label = token.meta?.label ?? token.meta?.id?.toString() ?? token.content;
					runs.push(this.createFootnoteRun(label));
					break;
				}
				case 'link_open': {
					const closeIndex = this.findClosingTokenIndex(tokens, i);
					const innerTokens = tokens.slice(i + 1, closeIndex);
					const href = token.attrGet('href') || '';
					const hyperlinkStyle: TextStyle = {
						...baseStyle,
						underline: true,
						color: baseStyle.color || DEFAULT_HYPERLINK_COLOR,
					};
					const childRuns = this.tokensToRuns(innerTokens, hyperlinkStyle, allowFootnotes, false);
					const textChildren = childRuns.filter((child): child is TextRun => child instanceof TextRun);
					const trailingRuns = childRuns.filter(child => !(child instanceof TextRun));
					if (allowHyperlinks && href && textChildren.length > 0) {
						runs.push(new ExternalHyperlink({ link: href, children: textChildren }));
						runs.push(...trailingRuns);
					} else {
						runs.push(...childRuns);
					}
					i = closeIndex;
					break;
				}
				case 'strong_open':
				case 'em_open':
				case 's_open':
				case 'mark_open':
				case 'sup_open':
				case 'sub_open': {
					const closeIndex = this.findClosingTokenIndex(tokens, i);
					const innerTokens = tokens.slice(i + 1, closeIndex);
					const nextStyle: TextStyle = { ...baseStyle };
					if (token.type === 'strong_open') nextStyle.bold = true;
					if (token.type === 'em_open') nextStyle.italic = true;
					if (token.type === 's_open') nextStyle.strikethrough = true;
					if (token.type === 'mark_open') nextStyle.highlight = true;
					if (token.type === 'sup_open') nextStyle.superScript = true;
					if (token.type === 'sub_open') nextStyle.subScript = true;
					runs.push(...this.tokensToRuns(innerTokens, nextStyle, allowFootnotes, allowHyperlinks));
					i = closeIndex;
					break;
				}
				default: {
					if (token.children && token.children.length > 0) {
						runs.push(...this.tokensToRuns(token.children, baseStyle, allowFootnotes, allowHyperlinks));
					}
					break;
				}
			}
		}

		return runs;
	}

	private findClosingTokenIndex(tokens: Token[], startIndex: number): number {
		const startToken = tokens[startIndex];
		if (!startToken || startToken.nesting !== 1) {
			return startIndex;
		}
		let depth = 1;
		for (let i = startIndex + 1; i < tokens.length; i++) {
			depth += tokens[i].nesting;
			if (depth === 0) {
				return i;
			}
		}
		return startIndex;
	}

	private createTextRunsFromString(content: string, style: TextStyle, allowFootnotes: boolean): (TextRun | ExternalHyperlink)[] {
		if (!content) {
			return [];
		}
		if (!allowFootnotes) {
			return [this.createTextRun(content, style)];
		}

		const regex = /\[\^([^\]]+)\]/g;
		const runs: (TextRun | ExternalHyperlink)[] = [];
		let lastIndex = 0;
		let match: RegExpExecArray | null;

		while ((match = regex.exec(content)) !== null) {
			if (match.index > lastIndex) {
				runs.push(this.createTextRun(content.slice(lastIndex, match.index), style));
			}
			runs.push(this.createFootnoteRun(match[1]));
			lastIndex = match.index + match[0].length;
		}

		if (lastIndex < content.length) {
			runs.push(this.createTextRun(content.slice(lastIndex), style));
		}

		return runs;
	}

	private parseHtmlStringToRuns(html: string, baseStyle: TextStyle, allowFootnotes: boolean): (TextRun | ExternalHyperlink)[] {
		if (!html) {
			return [];
		}

		const parser = new DOMParser();
		const doc = parser.parseFromString(`<wrapper>${html}</wrapper>`, 'text/html');
		const runs = this.parseHtmlNodes(doc.body.childNodes, baseStyle, allowFootnotes);
		if (runs.length === 0) {
			const fallback = doc.body.textContent ?? '';
			if (fallback.trim().length > 0) {
				return this.createTextRunsFromString(fallback, baseStyle, allowFootnotes);
			}
		}
		return runs;
	}

	private parseHtmlNodes(nodes: NodeListOf<ChildNode> | ChildNode[], baseStyle: TextStyle, allowFootnotes: boolean): (TextRun | ExternalHyperlink)[] {
		const runs: (TextRun | ExternalHyperlink)[] = [];

		nodes.forEach((node: ChildNode) => {
			if (node.nodeType === Node.TEXT_NODE) {
				runs.push(...this.createTextRunsFromString(node.textContent ?? '', baseStyle, allowFootnotes));
				return;
			}

			if (node.nodeType !== Node.ELEMENT_NODE) {
				return;
			}

			const element = node as HTMLElement;
			const tag = element.tagName.toLowerCase();
			const nextStyle: TextStyle = { ...baseStyle };

			switch (tag) {
				case 'br':
					runs.push(new TextRun({ text: '', break: 1 }));
					return;
				case 'strong':
				case 'b':
					nextStyle.bold = true;
					break;
				case 'em':
				case 'i':
					nextStyle.italic = true;
					break;
				case 's':
				case 'strike':
				case 'del':
					nextStyle.strikethrough = true;
					break;
				case 'u':
					nextStyle.underline = true;
					break;
				case 'code':
					nextStyle.code = true;
					break;
				case 'mark':
					nextStyle.highlight = true;
					break;
				case 'sup':
					nextStyle.superScript = true;
					break;
				case 'sub':
					nextStyle.subScript = true;
					break;
				case 'span': {
					const color = this.resolveSyntaxColor(element.classList);
					if (color) {
						nextStyle.color = color;
					}
					break;
				}
				case 'a': {
					const href = element.getAttribute('href') || '';
					const childRuns = this.parseHtmlNodes(element.childNodes, baseStyle, allowFootnotes);
					const textChildren = childRuns.filter((child): child is TextRun => child instanceof TextRun);
					if (href && textChildren.length > 0) {
						runs.push(new ExternalHyperlink({ link: href, children: textChildren }));
					} else {
						runs.push(...childRuns);
					}
					return;
				}
				default:
					break;
			}

			runs.push(...this.parseHtmlNodes(element.childNodes, nextStyle, allowFootnotes));
		});

		return runs;
	}

	private parseHighlightedHtmlLine(lineHtml: string): TextRun[] {
		const container = document.createElement('div');
		container.innerHTML = lineHtml || '';
		const runs = this.parseHtmlNodes(container.childNodes, { code: true, codeBlock: true }, false)
			.filter((run): run is TextRun => run instanceof TextRun);
		return runs.length > 0 ? runs : [this.createTextRun('', { code: true, codeBlock: true })];
	}

	private resolveSyntaxColor(classList: DOMTokenList): string | undefined {
		for (const cls of Array.from(classList)) {
			const color = CODE_SYNTAX_COLOR_MAP[cls];
			if (color) {
				return color;
			}
		}
		return undefined;
	}

	private createTable(rows: string[][], alignments: string[]): (Paragraph | Table)[] {
		if (rows.length === 0) return [];

		const columnCount = rows[0].length;
		const totalWidth = convertInchesToTwip(6.5);
		const columnWidth = Math.max(1, Math.floor(totalWidth / columnCount));

		const tableRows = rows.map((row, rowIndex) => {
			const cells = row.map((cell, cellIndex) => {
				const alignment = alignments[cellIndex] || 'left';
				let alignmentType: typeof AlignmentType.LEFT | typeof AlignmentType.CENTER | typeof AlignmentType.RIGHT = AlignmentType.LEFT;
				if (alignment === 'center') alignmentType = AlignmentType.CENTER;
				if (alignment === 'right') alignmentType = AlignmentType.RIGHT;

				return new TableCell({
					children: [
						new Paragraph({
							children: this.parseInlineFormatting(cell),
							alignment: alignmentType,
						}),
					],
					width: {
						size: columnWidth,
						type: WidthType.DXA,
					},
					shading: rowIndex === 0 ? { fill: 'E7E6E6' } : undefined,
				});
			});

			return new TableRow({
				children: cells,
			});
		});

		const table = new Table({
			rows: tableRows,
			width: {
				size: totalWidth,
				type: WidthType.DXA,
			},
			columnWidths: new Array(columnCount).fill(columnWidth),
			layout: TableLayoutType.FIXED,
		});

		// Return table wrapped in paragraphs for spacing
		return [new Paragraph({ children: [] }), table, new Paragraph({ children: [] })];
	}

	private parseInlineFormatting(text: string, options?: { allowFootnotes?: boolean }): (TextRun | ExternalHyperlink)[] {
		if (!this.settings.preserveFormatting) {
			return this.createPlainTextRuns(text);
		}

		const env: Record<string, unknown> = {};
		const tokens = this.md.parseInline(text, env);
		const inlineTokens = tokens.length > 0 ? tokens[0].children ?? [] : [];
		const allowFootnotes = options?.allowFootnotes !== false;

		const runs = inlineTokens.length
			? this.tokensToRuns(inlineTokens, {}, allowFootnotes, true)
			: this.createPlainTextRuns(text);

		return runs.length > 0 ? runs : this.createPlainTextRuns(text);
	}

	private createTextRun(text: string, style: TextStyle): TextRun {
		let bodyFontSize: number;
		let textFont: string;
		let monospaceFont: string;

		if (this.settings.useObsidianAppearance && this.obsidianFonts) {
			// Use actual Obsidian font settings
			bodyFontSize = this.obsidianFonts.baseFontSize;
			textFont = this.obsidianFonts.textFont;
			monospaceFont = this.obsidianFonts.monospaceFont;
			
			// Validate text font
			if (!textFont || textFont.trim() === '' || textFont === 'undefined' || textFont === '??' || textFont.includes('??')) {
				textFont = this.settings.defaultFontFamily;
			}
			
			// Platform-specific fallback if monospace font is empty
			if (!monospaceFont || monospaceFont.trim() === '' || monospaceFont === 'undefined' || monospaceFont === '??' || monospaceFont.includes('??')) {
				monospaceFont = this.getPlatformMonospaceFont();
			}
		} else {
			// Use custom settings
			bodyFontSize = this.settings.defaultFontSize;
			textFont = this.settings.defaultFontFamily;
			// Platform-specific monospace font
			monospaceFont = this.getPlatformMonospaceFont();
		}
		
		// Final validation - ensure we never have an empty or invalid font
		if (!monospaceFont || monospaceFont.trim() === '' || monospaceFont === 'undefined' || monospaceFont === '??' || monospaceFont.includes('??')) {
			monospaceFont = 'Courier New';
		}
		
		const targetFont = style.code ? monospaceFont : textFont;
		
		// Validate target font is not empty
		if (!targetFont || targetFont.trim() === '' || targetFont === 'undefined' || targetFont === '??' || targetFont.includes('??')) {
			// Use appropriate fallback based on whether it's code or text
			const finalFont = style.code ? 'Courier New' : this.settings.defaultFontFamily;
			
			const options: any = {
				text: text,
				size: bodyFontSize * 2,
				font: finalFont,
			};
			
			if (style.bold) options.bold = true;
			if (style.code) options.bold = true; // Make code bold
			if (style.italic) options.italics = true;
			if (style.strikethrough) options.strike = true;
			if (style.underline) options.underline = { type: UnderlineType.SINGLE };
			if (style.highlight) options.highlight = 'yellow';
			if (style.color) options.color = style.color;
			if (style.superScript) options.superScript = true;
			if (style.subScript) options.subScript = true;

			const shadingFill = style.backgroundColor ?? ((style.code && !style.codeBlock) ? 'F5F5F5' : undefined);
			if (shadingFill) {
				options.shading = { fill: shadingFill };
			}

			return new TextRun(options);
		}
		
		const options: any = {
			text: text,
			size: bodyFontSize * 2,
			font: targetFont,
		};

		if (style.bold) options.bold = true;
		if (style.code) options.bold = true; // Make code bold
		if (style.italic) options.italics = true;
		if (style.strikethrough) options.strike = true;
		if (style.underline) options.underline = { type: UnderlineType.SINGLE };
		if (style.highlight) options.highlight = 'yellow';
		if (style.color) options.color = style.color;
		if (style.superScript) options.superScript = true;
		if (style.subScript) options.subScript = true;

		const shadingFill = style.backgroundColor ?? ((style.code && !style.codeBlock) ? 'F5F5F5' : undefined);
		if (shadingFill) {
			options.shading = { fill: shadingFill };
		}

		return new TextRun(options);
	}

	private createFootnoteRun(label: string): TextRun {
		let index = this.usedFootnotes.indexOf(label);
		if (index === -1) {
			index = this.usedFootnotes.length;
			this.usedFootnotes.push(label);
		}
		const number = index + 1;
		return this.createTextRun(number.toString(), { superScript: true });
	}

	private createNumbering() {
		const textFont = (this.settings.useObsidianAppearance && this.obsidianFonts)
			? this.obsidianFonts.textFont
			: this.settings.defaultFontFamily;
		const baseSize = (this.settings.useObsidianAppearance && this.obsidianFonts)
			? this.obsidianFonts.baseFontSize * 2
			: this.settings.defaultFontSize * 2;

		const makeLevel = (level: number) => {
			const baseIndentInches = 0.18 + level * 0.18;
			const indentLeft = convertInchesToTwip(baseIndentInches);
			const hanging = convertInchesToTwip(0.18);
			return {
				level,
				format: LevelFormat.DECIMAL,
				restart: 1,
				text: `%${level + 1}.`,
				alignment: NumberAlignment.LEFT,
				style: {
					paragraph: {
						indent: {
							left: indentLeft,
							hanging: hanging,
						},
						tabStops: [{
							type: TabStopType.LEFT,
							position: indentLeft,
						}],
					},
					run: {
						font: textFont,
						size: baseSize,
					},
				},
			};
		};

		const makeBulletLevel = (level: number) => {
			const baseIndentInches = 0.18 + level * 0.18;
			const indentLeft = convertInchesToTwip(baseIndentInches);
			const hanging = convertInchesToTwip(0.18);
			return {
				level,
				format: LevelFormat.BULLET,
				restart: 1,
				text: '\u2022',
				alignment: NumberAlignment.LEFT,
				style: {
					paragraph: {
						indent: {
							left: indentLeft,
							hanging: hanging,
						},
						tabStops: [{
							type: TabStopType.LEFT,
							position: indentLeft,
						}],
					},
					run: {
						font: textFont,
						size: baseSize,
					},
				},
			};
		};

		return {
			config: [
				{
					reference: 'obsidian-numbered',
					levels: [0, 1, 2].map(makeLevel),
				},
				{
					reference: 'obsidian-bullet',
					levels: [0, 1, 2].map(makeBulletLevel),
				},
			],
		};
	}

	private async createEmbeddedImageParagraph(rawLink: string): Promise<Paragraph | null> {
		if (!this.resourceLoader) {
			return null;
		}

		const parsed = this.parseEmbeddedImageLink(rawLink);
		if (!parsed.target) {
			return null;
		}

		const data = await this.resourceLoader(parsed.target);
		if (!data) {
			return null;
		}

		return await this.buildImageParagraph(data, parsed.widthOverride);
	}

	private parseEmbeddedImageLink(rawLink: string): { target: string; widthOverride?: number } {
		const parts = rawLink.split('|');
		const targetPart = parts.shift()?.trim() ?? '';
		const sanitizedTarget = targetPart.split('#')[0].split('^')[0].trim();

		let widthOverride: number | undefined;
		if (parts.length > 0) {
			const remainder = parts.join('|').trim();
			const widthMatch = remainder.match(/^(\d+)(?:px)?$/i);
			if (widthMatch) {
				widthOverride = parseInt(widthMatch[1], 10);
			}
		}

		return {
			target: sanitizedTarget,
			widthOverride,
		};
	}

	private parseStandardImageTarget(rawTarget: string): { target: string; widthOverride?: number } {
		let targetPart = rawTarget.trim();

		// Remove optional title component ("title")
		const titleMatch = targetPart.match(/\s+"[^"]*"\s*$/);
		if (titleMatch) {
			targetPart = targetPart.slice(0, titleMatch.index).trim();
		}

		const segments = targetPart.split('|');
		const target = segments.shift()?.trim() ?? '';
		let widthOverride: number | undefined;

		for (const segment of segments) {
			const widthMatch = segment.trim().match(/^(\d+)(?:px)?$/i);
			if (widthMatch) {
				widthOverride = parseInt(widthMatch[1], 10);
				break;
			}
		}

		return { target, widthOverride };
	}

	private async createStandardImageParagraph(altText: string, rawTarget: string): Promise<Paragraph | null> {
		const parsed = this.parseStandardImageTarget(rawTarget);
		if (!parsed.target) {
			return null;
		}

		const data = await this.loadImageData(parsed.target);
		if (!data) {
			return null;
		}

		return await this.buildImageParagraph(data, parsed.widthOverride);
	}

	private async loadImageData(target: string): Promise<ArrayBuffer | null> {
		const trimmedTarget = target.trim();
		if (/^https?:\/\//i.test(trimmedTarget)) {
			if (typeof fetch === 'undefined') {
				return null;
			}
			try {
				const response = await fetch(trimmedTarget);
				if (!response.ok) {
					return null;
				}
				return await response.arrayBuffer();
			} catch {
				return null;
			}
		}

		if (this.resourceLoader) {
			try {
				return await this.resourceLoader(trimmedTarget);
			} catch {
				return null;
			}
		}

		return null;
	}

	private async buildImageParagraph(data: ArrayBuffer, widthOverride?: number): Promise<Paragraph> {
		const dimensions = await this.getImageDimensions(data).catch(() => null);
		let width = dimensions?.width ?? 400;
		let height = dimensions?.height ?? 300;

		if (widthOverride) {
			if (dimensions && dimensions.width > 0) {
				const ratio = dimensions.height / dimensions.width;
				width = widthOverride;
				height = Math.max(1, Math.round(widthOverride * ratio));
			} else {
				width = widthOverride;
				height = widthOverride;
			}
		} else if (dimensions && dimensions.width > 0) {
			const maxWidth = 680; // roughly 7 inches at 96 DPI
			if (width > maxWidth) {
				const ratio = dimensions.height / dimensions.width;
				width = maxWidth;
				height = Math.max(1, Math.round(maxWidth * ratio));
			}
		}

		return new Paragraph({
			alignment: AlignmentType.CENTER,
			children: [
				new ImageRun({
					data: new Uint8Array(data),
					transformation: {
						width: Math.max(1, Math.round(width)),
						height: Math.max(1, Math.round(height)),
					},
				}),
			],
		});
	}

	private async getImageDimensions(data: ArrayBuffer): Promise<{ width: number; height: number } | null> {
		if (typeof Image === 'undefined' || typeof URL === 'undefined') {
			return null;
		}

		return await new Promise((resolve) => {
			const blob = new Blob([data]);
			const url = URL.createObjectURL(blob);
			const image = new Image();
			image.onload = () => {
				const width = image.naturalWidth || image.width;
				const height = image.naturalHeight || image.height;
				URL.revokeObjectURL(url);
				resolve(width > 0 && height > 0 ? { width, height } : null);
			};
			image.onerror = () => {
				URL.revokeObjectURL(url);
				resolve(null);
			};
			image.src = url;
		});
	}

	private extractFootnotes(markdown: string): { content: string; definitions: Map<string, string> } {
		const lines = markdown.split('\n');
		const filteredLines: string[] = [];
		const definitions = new Map<string, string>();

		for (let i = 0; i < lines.length; i++) {
			const line = lines[i];
			const match = line.match(/^\[\^([^\]]+)\]:\s*(.*)$/);
			if (match) {
				const label = match[1].trim();
				const definitionParts: string[] = [];
				if (match[2]) {
					definitionParts.push(match[2].trim());
				}

				let j = i + 1;
				while (j < lines.length && /^\s{2,}.+/.test(lines[j])) {
					definitionParts.push(lines[j].trim());
					j++;
				}

				definitions.set(label, definitionParts.join(' ').trim());
				i = j - 1;
			} else {
				filteredLines.push(line);
			}
		}

		return { content: filteredLines.join('\n'), definitions };
	}

	private appendFootnotes(paragraphs: (Paragraph | Table)[]) {
		if (this.usedFootnotes.length === 0) {
			return;
		}

		paragraphs.push(new Paragraph({ children: [] }));
		paragraphs.push(this.createHeading('Footnotes', 2));

		this.usedFootnotes.forEach((label, index) => {
			const definition = this.footnoteDefinitions.get(label) || `[Missing footnote: ${label}]`;
			const children = [
				new TextRun({ text: `${index + 1}. `, bold: true }),
				...this.parseInlineFormatting(definition, { allowFootnotes: false }),
			];

			paragraphs.push(new Paragraph({
				children,
				spacing: this.getLineSpacing(),
			}));
		});
	}
}
