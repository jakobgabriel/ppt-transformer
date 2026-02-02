# PPTX Template Migrator

A web-based tool for migrating PowerPoint presentations to new templates at the OOXML (Open XML) level. This tool works directly with the XML structure inside `.pptx` files, ensuring all content is preserved while applying new branding, themes, and layouts.

## Features

- **Content Preservation**: All text, images, charts, tables, speaker notes, and embedded objects are preserved
- **Smart Layout Mapping**: Automatically maps source layouts to target layouts using name matching, type matching, and structural analysis
- **Custom Mapping Override**: Manually specify layout mappings when needed
- **Theme Remapping**: Automatically adapts theme-relative colors and fonts to the new template
- **Migration Report**: Detailed report of all changes, warnings, and recommendations

## How It Works

1. **Upload Files**: Provide your source presentation and target template
2. **Review Plan**: See the automatic layout mappings and adjust if needed
3. **Migrate**: Generate the migrated presentation with new branding

## Installation

```bash
# Clone or navigate to the project directory
cd ppt-transformer

# Install dependencies
npm install

# Start the server
npm start
```

The application will be available at `http://localhost:3000`

## API Endpoints

### `POST /api/analyze`
Analyzes both files and returns a migration plan.

**Form data:**
- `source`: The source .pptx file
- `template`: The target template .pptx file

**Response:**
```json
{
  "success": true,
  "source": { "slideCount": 12, "slides": [...] },
  "template": { "layoutCount": 11, "layouts": [...] },
  "migrationPlan": [...],
  "warnings": [...]
}
```

### `POST /api/migrate`
Performs the migration and returns the migrated .pptx file.

**Form data:**
- `source`: The source .pptx file
- `template`: The target template .pptx file
- `mappingInstructions` (optional): Custom layout mappings

**Response:** Binary .pptx file download

## Technical Details

### OOXML Structure

A `.pptx` file is a ZIP archive containing XML files:

```
[Content_Types].xml          - Media type registry
_rels/.rels                  - Top-level relationships
ppt/
  presentation.xml           - Slide order, sizes
  slides/slide*.xml          - Individual slide content
  slideLayouts/              - Layout definitions
  slideMasters/              - Master slide definitions
  theme/theme1.xml           - Color/font schemes
  media/                     - Images, videos
  charts/                    - Embedded charts
```

### Layout Mapping Priority

1. **User-specified mapping** - Explicit overrides
2. **Name match** - "Title Slide" ≈ "Corporate Title"
3. **Type match** - Matching placeholder types
4. **Structural match** - Similar placeholder counts
5. **Fallback** - Uses generic content layout

### Content Transfer

- **Placeholder content**: Mapped to corresponding placeholders in target layout
- **Freeform content**: Copied with position adjustments if needed
- **Media files**: Copied to output with updated relationship IDs
- **Theme references**: Preserved (automatically adapt to new theme)
- **Explicit overrides**: RGB colors and explicit fonts preserved

## Limitations

- Does not support `.pptm` (macro-enabled) files
- Encrypted presentations are not supported
- External data links in charts may need manual verification
- OLE embedded objects should be checked after migration

## Development

```bash
# Run in development mode with auto-reload
npm run dev
```

## Dependencies

- **express** - Web server framework
- **multer** - File upload handling
- **jszip** - ZIP archive manipulation
- **xml2js** - XML parsing and building
- **uuid** - Unique ID generation

## License

MIT
