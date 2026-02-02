import express from 'express';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import { PPTXMigrator } from './lib/pptx-migrator.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// Configure multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: {
    fileSize: 100 * 1024 * 1024, // 100MB max
  },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext !== '.pptx') {
      return cb(new Error('Only .pptx files are allowed'));
    }
    cb(null, true);
  }
});

// Serve static files
app.use(express.static(path.join(__dirname, '../public')));
app.use(express.json());

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Analyze endpoint - analyze files and return migration plan
app.post('/api/analyze', upload.fields([
  { name: 'source', maxCount: 1 },
  { name: 'template', maxCount: 1 }
]), async (req, res) => {
  try {
    if (!req.files?.source?.[0] || !req.files?.template?.[0]) {
      return res.status(400).json({
        error: 'Both source presentation and template are required'
      });
    }

    const sourceBuffer = req.files.source[0].buffer;
    const templateBuffer = req.files.template[0].buffer;

    const migrator = new PPTXMigrator();

    // Load and analyze both files
    const JSZip = (await import('jszip')).default;
    migrator.sourceZip = await JSZip.loadAsync(sourceBuffer);
    migrator.templateZip = await JSZip.loadAsync(templateBuffer);

    // Analyze template
    migrator.templateAnalysis = await migrator.analyzeTemplate(migrator.templateZip);

    // Analyze source
    migrator.sourceAnalysis = await migrator.analyzeSource(migrator.sourceZip);

    // Create migration plan
    migrator.migrationPlan = migrator.createMigrationPlan(null);

    res.json({
      success: true,
      source: {
        filename: req.files.source[0].originalname,
        slideCount: migrator.sourceAnalysis.slides.length,
        slideSize: migrator.sourceAnalysis.slideSize,
        slides: migrator.sourceAnalysis.slides.map(s => ({
          number: s.number,
          title: s.title,
          layoutName: s.layoutName,
          hasNotes: !!s.notes,
          placeholderCount: s.placeholderContent.length,
          freeformCount: s.freeformContent.length
        }))
      },
      template: {
        filename: req.files.template[0].originalname,
        layoutCount: migrator.templateAnalysis.layouts.length,
        slideSize: migrator.templateAnalysis.slideSize,
        layouts: migrator.templateAnalysis.layouts.map(l => ({
          name: l.name,
          type: l.type,
          placeholderCount: l.placeholders.length,
          placeholders: l.placeholders.map(p => ({ type: p.type, idx: p.idx }))
        }))
      },
      migrationPlan: migrator.migrationPlan.map(p => ({
        slideNumber: p.slideNumber,
        slideTitle: p.slideTitle,
        sourceLayout: p.sourceLayout,
        targetLayout: p.targetLayout?.name,
        targetLayoutFile: p.targetLayout?.file,
        matchType: p.matchType
      })),
      warnings: migrator.warnings
    });

  } catch (error) {
    console.error('Analysis error:', error);
    res.status(500).json({
      error: 'Failed to analyze files',
      message: error.message
    });
  }
});

// Migrate endpoint - perform the actual migration
app.post('/api/migrate', upload.fields([
  { name: 'source', maxCount: 1 },
  { name: 'template', maxCount: 1 }
]), async (req, res) => {
  try {
    if (!req.files?.source?.[0] || !req.files?.template?.[0]) {
      return res.status(400).json({
        error: 'Both source presentation and template are required'
      });
    }

    const sourceBuffer = req.files.source[0].buffer;
    const templateBuffer = req.files.template[0].buffer;
    const mappingInstructions = req.body?.mappingInstructions || null;

    const migrator = new PPTXMigrator();
    const result = await migrator.migrate(sourceBuffer, templateBuffer, mappingInstructions);

    // Generate output filename
    const sourceBasename = path.basename(
      req.files.source[0].originalname,
      '.pptx'
    );
    const outputFilename = `${sourceBasename}_migrated.pptx`;

    // Send the file
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${outputFilename}"`);
    res.setHeader('X-Migration-Report', encodeURIComponent(JSON.stringify(result.report)));

    res.send(result.buffer);

  } catch (error) {
    console.error('Migration error:', error);
    res.status(500).json({
      error: 'Migration failed',
      message: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
});

// Error handling middleware
app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({
        error: 'File too large',
        message: 'Maximum file size is 100MB'
      });
    }
    return res.status(400).json({
      error: 'Upload error',
      message: error.message
    });
  }

  if (error.message === 'Only .pptx files are allowed') {
    return res.status(400).json({
      error: 'Invalid file type',
      message: error.message
    });
  }

  console.error('Server error:', error);
  res.status(500).json({
    error: 'Internal server error',
    message: error.message
  });
});

app.listen(PORT, () => {
  console.log(`PPTX Template Migrator running at http://localhost:${PORT}`);
  console.log('');
  console.log('Endpoints:');
  console.log('  GET  /              - Web interface');
  console.log('  POST /api/analyze   - Analyze files and get migration plan');
  console.log('  POST /api/migrate   - Perform migration and download result');
  console.log('');
});

export default app;
