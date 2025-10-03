# Changelog

All notable changes to the Export to Word plugin will be documented in this file.

## [1.3.1] - 2025-09-30

### Fixed (Critical Improvements)
- ✅ **Headers now use Obsidian text font** - All headers use the same font as body text
- ✅ **Header sizes scale with font size difference** - If default is 12pt and Obsidian is 16pt, headers multiply by 1.33x
- ✅ **Header colors match Obsidian** - Headers export with the exact color from your theme
- ✅ **List indentation reduced by 50%** - Now matches Obsidian's visual spacing exactly
- ✅ **Lists use text font** - Numbered and bulleted lists use Obsidian's text font
- ✅ **Bullets are round and filled** - Standard bullet style matches Obsidian

### Technical Details
- Added `sizeMultiplier` calculation: `actualFontSize / defaultFontSize`
- Headers multiply their size by this multiplier
- Added RGB to hex color conversion for Word compatibility
- Reduced list indentation from 0.25" to 0.125" per level
- All list text now uses `textFont` from Obsidian settings

## [1.3.0] - 2025-09-30

### Added
- **Include filename as header** setting - Adds document title as H1 at the top
- **Hyperlink support** - Markdown links `[text](url)` now export as clickable links
- **Line height matching** - Exports now match Obsidian's line spacing
- **Proper numbered lists** - Fixed ordered lists to display correctly in Word
- **Better heading font detection** - Reads actual heading fonts from your theme

### Fixed
- ✅ **Heading fonts** - Now uses the correct font family from your Obsidian theme
- ✅ **Heading sizes** - Properly scaled based on actual Obsidian computed styles
- ✅ **Numbered lists** - Now display correctly with proper numbering
- ✅ **List indentation** - Reduced to match Obsidian's visual appearance (was too wide)
- ✅ **Web links** - Markdown links now export as functional hyperlinks
- ✅ **Line spacing** - Matches Obsidian's line height settings

### Improved
- Heading detection now reads from multiple sources (preview, editor, CSS)
- Font detection includes per-heading font families
- Line height automatically adapts to your Obsidian settings
- Better handling of nested formatting and links

## [1.2.0] - 2025-09-30

### Added
- **Use Obsidian Appearance** setting to match Obsidian's visual style
  - **Automatically detects** Obsidian's actual font family from your theme
  - **Automatically detects** Obsidian's actual font size settings
  - **Automatically detects** heading sizes based on your Obsidian configuration
  - **Automatically detects** monospace font for code blocks
  - Reads CSS variables and computed styles directly from Obsidian
  - Dynamically adjusts to your theme and font size settings

### Changed
- List indentation now uses 2 spaces per level (was 4) for better visual match
- Heading sizes now scale proportionally based on your Obsidian text size
- Code blocks use your theme's monospace font instead of hardcoded Courier New
- Body text uses your theme's text font instead of hardcoded Calibri

### Fixed
- Tabs/indentation are no longer too wide in exported documents
- Heading sizes now properly match Obsidian's visual hierarchy
- Font families now match your Obsidian theme exactly
- Font sizes scale correctly with your Obsidian settings

## [1.1.0] - 2025-09-30

### Changed
- **BREAKING**: Files now save directly to vault instead of browser downloads
- Improved cross-platform compatibility with vault-based file saving
- Better integration with Obsidian's file system

### Added
- Output location settings with three options:
  - Same folder as markdown file (default)
  - Vault root directory
  - Custom folder within vault
- Custom output folder configuration
- Automatic folder creation for custom output paths
- Better file path handling across all platforms

### Improved
- iOS/iPadOS compatibility - files now save properly in vault
- Android compatibility - consistent behavior across platforms
- File organization - keep exports organized with flexible location options

## [1.0.0] - 2025-09-30

### Added
- Initial release of Export to Word plugin
- Export markdown files to Microsoft Word (.docx) format
- Full text styling support (bold, italic, strikethrough, highlights)
- Code block formatting with monospace font and background
- Table export with alignment support
- Blockquote formatting
- Ordered and unordered lists (including nested lists)
- Heading levels (H1-H6)
- Horizontal rules
- Customizable font family and size settings
- Toggle for metadata inclusion
- Toggle for formatting preservation
- Ribbon icon for quick export
- Command palette integration
- Context menu integration (right-click on files)
- Cross-platform support (Windows, macOS, Linux, iOS/iPadOS)
- esbuild integration for fast bundling
- TypeScript support

### Features
- **Markdown Support**: Comprehensive markdown syntax conversion
- **Styling**: Preserves bold, italic, strikethrough, highlights, and inline code
- **Tables**: Full table support with column alignment
- **Code Blocks**: Syntax-highlighted code blocks with proper formatting
- **Lists**: Nested ordered and unordered lists
- **Customization**: Configurable fonts and export options
- **Multi-platform**: Works on desktop and mobile platforms

### Technical
- Built with Obsidian API
- Uses `docx` library for Word document generation
- TypeScript for type safety
- esbuild for fast compilation
- Follows Obsidian plugin best practices

## Future Enhancements (Planned)

- [ ] Image embedding support
- [ ] Custom styling templates
- [ ] Batch export multiple files
- [ ] Export entire folders
- [ ] Custom header/footer support
- [ ] Page numbering options
- [ ] Table of contents generation
- [ ] Footnote support
- [ ] Link preservation
- [ ] Obsidian-specific syntax (wikilinks, embeds)
