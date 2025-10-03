# Export to Word - Obsidian Plugin

Export your Obsidian markdown files to Microsoft Word (docx) format with full text styling and formatting preservation. Files are saved directly in your vault for easy access across all platforms.

## Features

- 📝 **Full Markdown Support**: Converts headings, lists, tables, code blocks, blockquotes, and more
- 🎨 **Text Styling**: Preserves bold, italic, strikethrough, highlights, and inline code
- 🔤 **Obsidian Appearance Matching**: Automatically uses your Obsidian fonts and sizes
- 💻 **Code Formatting**: Bold monospace font (Courier New) with syntax highlighting colors
- 📊 **Tables**: Exports markdown tables with proper formatting and alignment
- 🖼️ **Images**: Supports both standard markdown and Obsidian-style embedded images
- 🔗 **Hyperlinks**: Clickable links with proper styling
- 📐 **Page Sizes**: Choose from A4, A5, A3, Letter, Legal, or Tabloid
- 💾 **Vault Integration**: Saves files directly in your vault (not browser downloads)
- 📁 **Flexible Output**: Choose where to save - same folder, vault root, or custom folder
- ⚙️ **Configurable**: Settings panel to customize export behavior
- 🖱️ **Multiple Export Options**: 
  - Ribbon icon for quick export
  - Command palette integration
  - Right-click context menu on files

## Installation

### Manual Installation

1. Download the latest release from the releases page
2. Extract the files to your vault's plugins folder: `<vault>/.obsidian/plugins/to-word/`
3. Reload Obsidian
4. Enable the plugin in Settings → Community Plugins

### Development Installation

1. Clone this repository into your vault's plugins folder
2. Run `npm install` to install dependencies
3. Run `npm run dev` to start compilation in watch mode
4. Reload Obsidian
5. Enable the plugin in Settings → Community Plugins

## Usage

### Export Current File

1. **Using Ribbon Icon**: Click the document icon in the left ribbon
2. **Using Command Palette**: Press `Ctrl/Cmd + P` and search for "Export current file to Word"
3. **Using Context Menu**: Right-click on a markdown file and select "Export to Word"

The exported Word document will be saved in your vault according to your output location settings.

## Example Files

The `examples/` folder contains sample markdown files demonstrating various features:

- `FULL_MARKDOWN_SAMPLE.md` - Comprehensive test of all supported markdown features
- `TEST_ISSUES.md` - Simple test file for debugging
- `SAMPLE.md` - Basic example

You can use these to test the plugin and see how different markdown elements are converted to Word format.

## Settings

Access plugin settings via Settings → Export to Word:

- **Default Font Family**: Set the default font for exported documents (default: Calibri)
- **Default Font Size**: Set the default font size in points (default: 11)
- **Include Metadata**: Choose whether to include frontmatter metadata in exports
- **Preserve Formatting**: Toggle markdown formatting preservation (bold, italic, etc.)
- **Use Obsidian Appearance**: **Automatically match your Obsidian theme's appearance**
  - When enabled: Reads your actual Obsidian settings including:
    - Text font family from your theme
    - Monospace font for code blocks (with fallback to Courier New)
    - Base font size (e.g., 16pt if that's your setting)
    - Heading sizes that scale proportionally with your text size
    - Line height and spacing
  - When disabled: Uses standard Word document sizes with your custom settings
  - **Smart detection**: Adapts to theme changes and font size adjustments
- **Include Filename as Header**: Add the filename as an H1 heading at the top of the document
- **Page Size**: Choose document page size (default: A4)
  - A4 (210 × 297 mm)
  - A5 (148 × 210 mm)
  - A3 (297 × 420 mm)
  - Letter (8.5 × 11 inches)
  - Legal (8.5 × 14 inches)
  - Tabloid (11 × 17 inches)
- **Output Location**: Choose where to save exported files:
  - **Same folder as markdown file**: Keeps exports next to source files
  - **Vault root**: Saves all exports to the vault root directory
  - **Custom folder**: Saves to a specified folder within your vault
- **Custom Output Folder**: Specify the folder path when using custom folder option (e.g., "Exports" or "Documents/Word")

## Supported Markdown Features

### Text Formatting

- ✅ **Bold** (`**text**`)
- ✅ *Italic* (`*text*`)
- ✅ ***Bold Italic*** (`***text***`)
- ✅ ~~Strikethrough~~ (`~~text~~`)
- ✅ ==Highlight== (`==text==`)
- ✅ `Inline code` (`` `code` ``) - Bold Courier New with background
- ✅ Underline (via HTML `<u>`)
- ✅ Superscript and Subscript (via HTML `<sup>`, `<sub>`)

### Structure

- ✅ Headings (H1-H6) with proper styling
- ✅ Paragraphs with line spacing
- ✅ Horizontal rules (`---`, `***`, `___`)
- ✅ Line breaks (two trailing spaces)
- ✅ Blockquotes (with nesting support)

### Lists

- ✅ Ordered lists (numbered)
- ✅ Unordered lists (bullets)
- ✅ Nested lists (up to 3 levels)
- ✅ Task lists (`- [ ]` and `- [x]`) - rendered as ☐ and ☑

### Code

- ✅ Inline code with monospace font and background
- ✅ Fenced code blocks with language support
- ✅ Syntax highlighting (colors preserved from highlight.js)
- ✅ Bold Courier New font for all code

### Tables

- ✅ Standard markdown tables
- ✅ Column alignment (left, center, right)
- ✅ Header row styling
- ✅ Fixed column widths

### Links & References

- ✅ Inline links (`[text](url)`)
- ✅ Titled links (`[text](url "title")`)
- ✅ Clickable hyperlinks in Word
- ✅ Footnotes (`[^1]`) with superscript references

### Images

- ✅ Standard markdown images (`![alt](url)`)
- ✅ Obsidian embedded images (`![[image.png]]`)
- ✅ Image sizing (`![alt](url|300)` or `![[image.png|300]]`)
- ✅ Remote images (http/https URLs)
- ✅ Local vault images

### Advanced

- ✅ Emojis (`:smile:` → 😊)
- ✅ Raw HTML formatting (`<b>`, `<i>`, `<code>`, etc.)
- ✅ Collapsible sections (`<details>`) - rendered expanded with visual indicators
- ✅ Definition lists (term + `: definition`)
- ✅ Nested blockquotes

### Limitations

- ⚠️ Task list checkboxes are static (not interactive in Word)
- ⚠️ Collapsible sections are always expanded (Word doesn't support interactive collapse)
- ⚠️ Some horizontal rule variants may not render in all contexts

See [WORD_LIMITATIONS.md](WORD_LIMITATIONS.md) for detailed explanations and workarounds.

## Platform Support

- ✅ **Windows**: Fully supported - saves to vault
- ✅ **macOS**: Fully supported - saves to vault
- ✅ **Linux**: Fully supported - saves to vault
- ✅ **iOS/iPadOS**: Not supported

All supported platforms save the exported Word document directly to your vault according to your output location settings. No browser downloads!

## Development

### Building

```bash
# Install dependencies
npm install

# Development mode (watch)
npm run dev

# Production build
npm run build
```

### Project Structure

- `main.ts` - Main plugin file with Obsidian integration
- `converter.ts` - Markdown to DOCX conversion logic
- `manifest.json` - Plugin manifest
- `esbuild.config.mjs` - Build configuration

## Technologies Used

- [Obsidian API](https://github.com/obsidianmd/obsidian-api)
- [docx](https://github.com/dolanmiu/docx) - Library for creating Word documents
- [markdown-it](https://github.com/markdown-it/markdown-it) - Markdown parser with plugin support
- [markdown-it-emoji](https://github.com/markdown-it/markdown-it-emoji) - Emoji support
- [markdown-it-mark](https://github.com/markdown-it/markdown-it-mark) - Highlight support
- [highlight.js](https://highlightjs.org/) - Syntax highlighting for code blocks
- [esbuild](https://esbuild.github.io/) - Fast JavaScript bundler
- TypeScript

## License

MIT

## Support

If you encounter any issues or have feature requests, please file them in the GitHub issues section.

## Credits

Created with ❤️ for the Obsidian community.
