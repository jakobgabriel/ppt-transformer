import JSZip from 'jszip';
import { parseStringPromise, Builder } from 'xml2js';
import { v4 as uuidv4 } from 'uuid';

/**
 * PPTX Template Migration Library
 * Migrates PowerPoint presentations to new templates at the OOXML level
 */

const XML_PARSER_OPTIONS = {
  explicitArray: false,
  preserveChildrenOrder: true,
  explicitChildren: true,
  attrkey: '$',
  charkey: '_',
  trim: false,
  normalize: false,
  normalizeTags: false,
  explicitRoot: true
};

const XML_BUILDER_OPTIONS = {
  renderOpts: { pretty: true, indent: '  ', newline: '\n' },
  xmldec: { version: '1.0', encoding: 'UTF-8', standalone: true },
  headless: false
};

/**
 * Main class for PPTX template migration
 */
export class PPTXMigrator {
  constructor() {
    this.sourceZip = null;
    this.templateZip = null;
    this.outputZip = null;
    this.sourceAnalysis = null;
    this.templateAnalysis = null;
    this.migrationPlan = [];
    this.warnings = [];
    this.mediaMapping = new Map();
    this.nextSlideId = 256;
    this.nextRelId = 1;
  }

  /**
   * Main migration entry point
   */
  async migrate(sourceBuffer, templateBuffer, mappingInstructions = null) {
    // Step 4.1: Unpack both files
    this.sourceZip = await JSZip.loadAsync(sourceBuffer);
    this.templateZip = await JSZip.loadAsync(templateBuffer);

    // Start output from template copy
    this.outputZip = await JSZip.loadAsync(templateBuffer);

    // Step 4.2: Analyze target template
    this.templateAnalysis = await this.analyzeTemplate(this.templateZip);

    // Step 4.3: Analyze source presentation
    this.sourceAnalysis = await this.analyzeSource(this.sourceZip);

    // Step 4.4: Map source layouts to target layouts
    this.migrationPlan = this.createMigrationPlan(mappingInstructions);

    // Step 4.5: Prepare output package (clear template slides)
    await this.prepareOutputPackage();

    // Step 4.6: Migrate slides
    await this.migrateSlides();

    // Step 4.7: Post-migration cleanup
    await this.postMigrationCleanup();

    // Step 4.8: Generate output
    const outputBuffer = await this.outputZip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    return {
      buffer: outputBuffer,
      report: this.generateReport()
    };
  }

  /**
   * Analyze the target template
   */
  async analyzeTemplate(zip) {
    const analysis = {
      layouts: [],
      masters: [],
      theme: null,
      slideSize: null
    };

    // Parse presentation.xml for slide size
    const presentationXml = await this.parseXmlFile(zip, 'ppt/presentation.xml');
    if (presentationXml) {
      const pres = presentationXml['p:presentation'] || presentationXml;
      if (pres['p:sldSz']) {
        analysis.slideSize = {
          cx: parseInt(pres['p:sldSz'].$?.cx || 9144000),
          cy: parseInt(pres['p:sldSz'].$?.cy || 6858000),
          type: pres['p:sldSz'].$?.type || 'custom'
        };
      }
    }

    // Parse presentation.xml.rels
    const presRels = await this.parseXmlFile(zip, 'ppt/_rels/presentation.xml.rels');

    // Get all layout files
    const layoutFiles = Object.keys(zip.files).filter(f =>
      f.startsWith('ppt/slideLayouts/') && f.endsWith('.xml') && !f.includes('_rels')
    );

    for (const layoutFile of layoutFiles) {
      const layoutXml = await this.parseXmlFile(zip, layoutFile);
      if (!layoutXml) continue;

      const layout = layoutXml['p:sldLayout'] || layoutXml;
      const cSld = layout['p:cSld'] || {};

      const layoutInfo = {
        file: layoutFile.split('/').pop(),
        path: layoutFile,
        name: cSld.$?.name || 'Unnamed Layout',
        type: layout.$?.type || 'obj',
        placeholders: this.extractPlaceholders(cSld)
      };

      // Get the master reference
      const layoutRelsPath = layoutFile.replace('slideLayouts/', 'slideLayouts/_rels/') + '.rels';
      const layoutRels = await this.parseXmlFile(zip, layoutRelsPath);
      if (layoutRels) {
        const rels = this.normalizeRels(layoutRels);
        const masterRel = rels.find(r => r.Type?.includes('slideMaster'));
        if (masterRel) {
          layoutInfo.masterRef = masterRel.Target;
        }
      }

      analysis.layouts.push(layoutInfo);
    }

    // Parse theme
    const themeFile = Object.keys(zip.files).find(f =>
      f.startsWith('ppt/theme/') && f.endsWith('.xml') && !f.includes('_rels')
    );
    if (themeFile) {
      analysis.theme = await this.parseXmlFile(zip, themeFile);
    }

    return analysis;
  }

  /**
   * Analyze source presentation
   */
  async analyzeSource(zip) {
    const analysis = {
      slides: [],
      slideSize: null,
      mediaFiles: []
    };

    // Parse presentation.xml
    const presentationXml = await this.parseXmlFile(zip, 'ppt/presentation.xml');
    if (!presentationXml) {
      throw new Error('Invalid PPTX: Missing presentation.xml');
    }

    const pres = presentationXml['p:presentation'] || presentationXml;

    // Get slide size
    if (pres['p:sldSz']) {
      analysis.slideSize = {
        cx: parseInt(pres['p:sldSz'].$?.cx || 9144000),
        cy: parseInt(pres['p:sldSz'].$?.cy || 6858000),
        type: pres['p:sldSz'].$?.type || 'custom'
      };
    }

    // Get slide list
    const sldIdLst = pres['p:sldIdLst'];
    let slideIds = [];
    if (sldIdLst) {
      const sldIds = sldIdLst['p:sldId'];
      slideIds = Array.isArray(sldIds) ? sldIds : (sldIds ? [sldIds] : []);
    }

    // Parse presentation.xml.rels to map r:id to file paths
    const presRels = await this.parseXmlFile(zip, 'ppt/_rels/presentation.xml.rels');
    const relMap = new Map();
    if (presRels) {
      const rels = this.normalizeRels(presRels);
      rels.forEach(r => relMap.set(r.Id, r.Target));
    }

    // Analyze each slide
    for (let i = 0; i < slideIds.length; i++) {
      const slideId = slideIds[i];
      const rId = slideId.$?.['r:id'];
      const slidePath = relMap.get(rId);

      if (!slidePath) continue;

      const fullPath = slidePath.startsWith('ppt/') ? slidePath : `ppt/${slidePath.replace('../', '')}`;
      const slideXml = await this.parseXmlFile(zip, fullPath);

      if (!slideXml) continue;

      const slideInfo = await this.analyzeSlide(zip, fullPath, slideXml, i + 1);
      analysis.slides.push(slideInfo);
    }

    // Get media files
    const mediaFiles = Object.keys(zip.files).filter(f => f.startsWith('ppt/media/'));
    analysis.mediaFiles = mediaFiles;

    return analysis;
  }

  /**
   * Analyze a single slide
   */
  async analyzeSlide(zip, slidePath, slideXml, slideNumber) {
    const slide = slideXml['p:sld'] || slideXml;
    const cSld = slide['p:cSld'] || {};

    const slideInfo = {
      number: slideNumber,
      path: slidePath,
      xml: slideXml,
      layout: null,
      layoutName: null,
      layoutType: null,
      title: null,
      placeholderContent: [],
      freeformContent: [],
      notes: null,
      transition: slide['p:transition'] || null,
      timing: slide['p:timing'] || null
    };

    // Get layout reference from slide rels
    const slideRelsPath = slidePath.replace('slides/', 'slides/_rels/') + '.rels';
    const slideRels = await this.parseXmlFile(zip, slideRelsPath);

    if (slideRels) {
      const rels = this.normalizeRels(slideRels);
      const layoutRel = rels.find(r => r.Type?.includes('slideLayout'));
      if (layoutRel) {
        slideInfo.layout = layoutRel.Target;

        // Get layout details
        const layoutPath = layoutRel.Target.startsWith('..')
          ? `ppt/slideLayouts/${layoutRel.Target.split('/').pop()}`
          : layoutRel.Target;
        const layoutXml = await this.parseXmlFile(zip, layoutPath);
        if (layoutXml) {
          const layout = layoutXml['p:sldLayout'] || layoutXml;
          const layoutCSld = layout['p:cSld'] || {};
          slideInfo.layoutName = layoutCSld.$?.name || 'Unknown';
          slideInfo.layoutType = layout.$?.type || 'obj';
        }
      }

      // Check for notes
      const notesRel = rels.find(r => r.Type?.includes('notesSlide'));
      if (notesRel) {
        const notesPath = notesRel.Target.startsWith('..')
          ? `ppt/notesSlides/${notesRel.Target.split('/').pop()}`
          : notesRel.Target;
        slideInfo.notes = await this.parseXmlFile(zip, notesPath);
        slideInfo.notesPath = notesPath;
      }
    }

    // Extract content from spTree
    const spTree = cSld['p:spTree'];
    if (spTree) {
      this.extractSlideContent(spTree, slideInfo);
    }

    // Try to extract title
    slideInfo.title = this.extractTitle(slideInfo) || `Slide ${slideNumber}`;

    return slideInfo;
  }

  /**
   * Extract content from shape tree
   */
  extractSlideContent(spTree, slideInfo) {
    // Process shapes
    const shapes = this.getShapes(spTree, 'p:sp');
    for (const shape of shapes) {
      const nvSpPr = shape['p:nvSpPr'] || {};
      const nvPr = nvSpPr['p:nvPr'] || {};

      if (nvPr['p:ph']) {
        // This is placeholder content
        slideInfo.placeholderContent.push({
          type: 'shape',
          element: shape,
          phType: nvPr['p:ph'].$?.type || 'body',
          phIdx: nvPr['p:ph'].$?.idx
        });
      } else {
        // Freeform content
        slideInfo.freeformContent.push({
          type: 'shape',
          element: shape
        });
      }
    }

    // Process pictures
    const pictures = this.getShapes(spTree, 'p:pic');
    for (const pic of pictures) {
      const nvPicPr = pic['p:nvPicPr'] || {};
      const nvPr = nvPicPr['p:nvPr'] || {};

      if (nvPr['p:ph']) {
        slideInfo.placeholderContent.push({
          type: 'picture',
          element: pic,
          phType: nvPr['p:ph'].$?.type || 'pic',
          phIdx: nvPr['p:ph'].$?.idx
        });
      } else {
        slideInfo.freeformContent.push({
          type: 'picture',
          element: pic
        });
      }
    }

    // Process graphic frames (charts, tables, diagrams)
    const graphicFrames = this.getShapes(spTree, 'p:graphicFrame');
    for (const gf of graphicFrames) {
      slideInfo.freeformContent.push({
        type: 'graphicFrame',
        element: gf
      });
    }

    // Process groups
    const groups = this.getShapes(spTree, 'p:grpSp');
    for (const grp of groups) {
      slideInfo.freeformContent.push({
        type: 'group',
        element: grp
      });
    }
  }

  /**
   * Get shapes of a specific type from spTree
   */
  getShapes(spTree, shapeName) {
    if (!spTree) return [];
    const shapes = spTree[shapeName];
    if (!shapes) return [];
    return Array.isArray(shapes) ? shapes : [shapes];
  }

  /**
   * Extract placeholders from a layout or master
   */
  extractPlaceholders(cSld) {
    const placeholders = [];
    const spTree = cSld['p:spTree'];
    if (!spTree) return placeholders;

    const shapes = this.getShapes(spTree, 'p:sp');
    for (const shape of shapes) {
      const nvSpPr = shape['p:nvSpPr'] || {};
      const nvPr = nvSpPr['p:nvPr'] || {};

      if (nvPr['p:ph']) {
        const ph = nvPr['p:ph'].$;
        const spPr = shape['p:spPr'] || {};
        const xfrm = spPr['a:xfrm'] || {};
        const off = xfrm['a:off'] || {};
        const ext = xfrm['a:ext'] || {};

        placeholders.push({
          type: ph?.type || 'body',
          idx: ph?.idx,
          position: {
            x: parseInt(off.$?.x || 0),
            y: parseInt(off.$?.y || 0)
          },
          size: {
            cx: parseInt(ext.$?.cx || 0),
            cy: parseInt(ext.$?.cy || 0)
          }
        });
      }
    }

    return placeholders;
  }

  /**
   * Extract title from slide
   */
  extractTitle(slideInfo) {
    for (const content of slideInfo.placeholderContent) {
      const phType = content.phType;
      if (phType === 'title' || phType === 'ctrTitle') {
        const txBody = content.element['p:txBody'];
        if (txBody) {
          return this.extractTextFromTxBody(txBody);
        }
      }
    }
    return null;
  }

  /**
   * Extract plain text from txBody
   */
  extractTextFromTxBody(txBody) {
    if (!txBody) return '';

    const paragraphs = txBody['a:p'];
    const pList = Array.isArray(paragraphs) ? paragraphs : (paragraphs ? [paragraphs] : []);

    const textParts = [];
    for (const p of pList) {
      const runs = p['a:r'];
      const rList = Array.isArray(runs) ? runs : (runs ? [runs] : []);

      for (const r of rList) {
        const t = r['a:t'];
        if (t) {
          textParts.push(typeof t === 'string' ? t : (t._ || ''));
        }
      }
    }

    return textParts.join(' ').trim().substring(0, 50);
  }

  /**
   * Create migration plan mapping source slides to target layouts
   */
  createMigrationPlan(mappingInstructions) {
    const plan = [];

    for (const slide of this.sourceAnalysis.slides) {
      let targetLayout = null;
      let matchType = 'fallback';

      // Priority 1: User mapping instructions
      if (mappingInstructions) {
        const userMapping = this.findUserMapping(slide, mappingInstructions);
        if (userMapping) {
          targetLayout = userMapping;
          matchType = 'user specified';
        }
      }

      // Priority 2: Name match
      if (!targetLayout) {
        targetLayout = this.findLayoutByName(slide.layoutName);
        if (targetLayout) matchType = 'name match';
      }

      // Priority 3: Type match
      if (!targetLayout) {
        targetLayout = this.findLayoutByType(slide.layoutType);
        if (targetLayout) matchType = 'type match';
      }

      // Priority 4: Structural match
      if (!targetLayout) {
        targetLayout = this.findLayoutByStructure(slide);
        if (targetLayout) matchType = 'structural match';
      }

      // Priority 5: Fallback to first content layout
      if (!targetLayout) {
        targetLayout = this.templateAnalysis.layouts.find(l =>
          l.type === 'obj' || l.name.toLowerCase().includes('content')
        ) || this.templateAnalysis.layouts[0];

        this.warnings.push(`Slide ${slide.number}: Using fallback layout "${targetLayout?.name}"`);
      }

      plan.push({
        slideNumber: slide.number,
        slideTitle: slide.title,
        sourceLayout: slide.layoutName,
        targetLayout: targetLayout,
        matchType: matchType
      });
    }

    return plan;
  }

  /**
   * Find layout by name (fuzzy match)
   */
  findLayoutByName(sourceName) {
    if (!sourceName) return null;

    const normalized = sourceName.toLowerCase().replace(/[^a-z0-9]/g, '');

    // Exact match first
    let match = this.templateAnalysis.layouts.find(l =>
      l.name.toLowerCase().replace(/[^a-z0-9]/g, '') === normalized
    );

    if (match) return match;

    // Fuzzy match
    const fuzzyMappings = {
      'title': ['title', 'cover', 'opening'],
      'titleslide': ['titleslide', 'corporatetitle', 'cover'],
      'twocontent': ['twocontent', 'twocolumn', 'comparison'],
      'sectionheader': ['sectionheader', 'section', 'sectionbreak', 'divider'],
      'blank': ['blank', 'empty'],
      'titleandcontent': ['titleandcontent', 'titlecontent', 'content']
    };

    for (const [key, values] of Object.entries(fuzzyMappings)) {
      if (normalized.includes(key) || values.some(v => normalized.includes(v))) {
        match = this.templateAnalysis.layouts.find(l => {
          const targetNorm = l.name.toLowerCase().replace(/[^a-z0-9]/g, '');
          return values.some(v => targetNorm.includes(v));
        });
        if (match) return match;
      }
    }

    return null;
  }

  /**
   * Find layout by type attribute
   */
  findLayoutByType(sourceType) {
    if (!sourceType) return null;

    return this.templateAnalysis.layouts.find(l => l.type === sourceType);
  }

  /**
   * Find layout by structural similarity
   */
  findLayoutByStructure(slide) {
    const sourcePlaceholders = slide.placeholderContent.length;

    // Find layout with closest placeholder count
    let bestMatch = null;
    let bestDiff = Infinity;

    for (const layout of this.templateAnalysis.layouts) {
      const diff = Math.abs(layout.placeholders.length - sourcePlaceholders);
      if (diff < bestDiff) {
        bestDiff = diff;
        bestMatch = layout;
      }
    }

    return bestMatch;
  }

  /**
   * Find user-specified mapping
   */
  findUserMapping(slide, instructions) {
    if (typeof instructions === 'string') {
      // Parse simple mapping format: "Source Layout -> Target Layout"
      const lines = instructions.split('\n');
      for (const line of lines) {
        const match = line.match(/["']?([^"']+)["']?\s*[-=]>\s*["']?([^"']+)["']?/i);
        if (match) {
          const sourcePattern = match[1].toLowerCase().trim();
          const targetName = match[2].trim();

          if (slide.layoutName?.toLowerCase().includes(sourcePattern) ||
              slide.title?.toLowerCase().includes(sourcePattern)) {
            return this.templateAnalysis.layouts.find(l =>
              l.name.toLowerCase().includes(targetName.toLowerCase())
            );
          }
        }
      }
    }
    return null;
  }

  /**
   * Prepare output package by removing template slides
   */
  async prepareOutputPackage() {
    // Parse presentation.xml
    const presentationXml = await this.parseXmlFile(this.outputZip, 'ppt/presentation.xml');
    const pres = presentationXml['p:presentation'] || presentationXml;

    // Clear slide ID list
    if (pres['p:sldIdLst']) {
      pres['p:sldIdLst'] = { 'p:sldId': [] };
    }

    // Update slide size if different
    if (this.sourceAnalysis.slideSize && this.templateAnalysis.slideSize) {
      const src = this.sourceAnalysis.slideSize;
      const tgt = this.templateAnalysis.slideSize;

      if (src.cx !== tgt.cx || src.cy !== tgt.cy) {
        this.warnings.push(
          `Slide dimensions differ: Source ${src.cx}x${src.cy} vs Template ${tgt.cx}x${tgt.cy}. Using template dimensions.`
        );
      }
    }

    // Write updated presentation.xml
    await this.writeXmlFile(this.outputZip, 'ppt/presentation.xml', presentationXml);

    // Remove existing slides
    const slideFiles = Object.keys(this.outputZip.files).filter(f =>
      f.startsWith('ppt/slides/') && !f.includes('_rels')
    );
    const slideRelsFiles = Object.keys(this.outputZip.files).filter(f =>
      f.startsWith('ppt/slides/_rels/')
    );

    for (const file of [...slideFiles, ...slideRelsFiles]) {
      this.outputZip.remove(file);
    }

    // Update presentation.xml.rels to remove slide references
    const presRelsPath = 'ppt/_rels/presentation.xml.rels';
    const presRels = await this.parseXmlFile(this.outputZip, presRelsPath);
    if (presRels) {
      const rels = this.normalizeRels(presRels);
      const filteredRels = rels.filter(r => !r.Type?.includes('/slide') || r.Type?.includes('slideLayout') || r.Type?.includes('slideMaster'));
      await this.writeRelsFile(this.outputZip, presRelsPath, filteredRels);

      // Track next rel ID
      const maxId = rels.reduce((max, r) => {
        const id = parseInt(r.Id.replace('rId', ''));
        return id > max ? id : max;
      }, 0);
      this.nextRelId = maxId + 1;
    }
  }

  /**
   * Migrate all slides
   */
  async migrateSlides() {
    const slideIds = [];

    for (let i = 0; i < this.migrationPlan.length; i++) {
      const planItem = this.migrationPlan[i];
      const sourceSlide = this.sourceAnalysis.slides.find(s => s.number === planItem.slideNumber);

      const slideId = await this.migrateSlide(sourceSlide, planItem.targetLayout, i + 1);
      slideIds.push(slideId);
    }

    // Update presentation.xml with new slide list
    const presentationXml = await this.parseXmlFile(this.outputZip, 'ppt/presentation.xml');
    const pres = presentationXml['p:presentation'] || presentationXml;

    pres['p:sldIdLst'] = {
      'p:sldId': slideIds.map(sid => ({
        $: { id: sid.id.toString(), 'r:id': sid.rId }
      }))
    };

    await this.writeXmlFile(this.outputZip, 'ppt/presentation.xml', presentationXml);
  }

  /**
   * Migrate a single slide
   */
  async migrateSlide(sourceSlide, targetLayout, slideNumber) {
    const slideFileName = `slide${slideNumber}.xml`;
    const slidePath = `ppt/slides/${slideFileName}`;
    const slideRelsPath = `ppt/slides/_rels/${slideFileName}.rels`;

    // Create new slide structure based on source
    const newSlide = this.createNewSlide(sourceSlide, targetLayout);

    // Create slide relationships
    const slideRels = [];

    // Add layout relationship
    const layoutRelPath = `../slideLayouts/${targetLayout.file}`;
    slideRels.push({
      Id: 'rId1',
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
      Target: layoutRelPath
    });

    let nextRelId = 2;

    // Copy media files and update references
    await this.copyMediaForSlide(sourceSlide, newSlide, slideRels, nextRelId);

    // Handle notes
    if (sourceSlide.notes) {
      const notesSlideNumber = slideNumber;
      const notesPath = `ppt/notesSlides/notesSlide${notesSlideNumber}.xml`;
      const notesRelsPath = `ppt/notesSlides/_rels/notesSlide${notesSlideNumber}.xml.rels`;

      // Copy notes content
      await this.writeXmlFile(this.outputZip, notesPath, sourceSlide.notes);

      // Create notes rels
      const notesRels = [{
        Id: 'rId1',
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
        Target: `../slides/${slideFileName}`
      }];
      await this.writeRelsFile(this.outputZip, notesRelsPath, notesRels);

      // Add notes rel to slide
      slideRels.push({
        Id: `rId${slideRels.length + 1}`,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
        Target: `../notesSlides/notesSlide${notesSlideNumber}.xml`
      });

      // Register notes in Content_Types
      await this.registerContentType(notesPath, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
    }

    // Write slide XML
    await this.writeXmlFile(this.outputZip, slidePath, { 'p:sld': newSlide });

    // Write slide rels
    await this.writeRelsFile(this.outputZip, slideRelsPath, slideRels);

    // Register slide in Content_Types
    await this.registerContentType(slidePath, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');

    // Add to presentation.xml.rels
    const presRelsPath = 'ppt/_rels/presentation.xml.rels';
    const presRels = await this.parseXmlFile(this.outputZip, presRelsPath);
    const rels = this.normalizeRels(presRels);

    const slideRelId = `rId${this.nextRelId++}`;
    rels.push({
      Id: slideRelId,
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      Target: `slides/${slideFileName}`
    });

    await this.writeRelsFile(this.outputZip, presRelsPath, rels);

    return {
      id: this.nextSlideId++,
      rId: slideRelId
    };
  }

  /**
   * Create a new slide structure
   */
  createNewSlide(sourceSlide, targetLayout) {
    const sourceXml = sourceSlide.xml['p:sld'] || sourceSlide.xml;

    // Start with source slide structure
    const newSlide = {
      $: {
        'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
      }
    };

    // Copy cSld (common slide data)
    if (sourceXml['p:cSld']) {
      newSlide['p:cSld'] = JSON.parse(JSON.stringify(sourceXml['p:cSld']));
    }

    // Copy clrMapOvr if present
    if (sourceXml['p:clrMapOvr']) {
      newSlide['p:clrMapOvr'] = JSON.parse(JSON.stringify(sourceXml['p:clrMapOvr']));
    }

    // Copy transition if present
    if (sourceSlide.transition) {
      newSlide['p:transition'] = JSON.parse(JSON.stringify(sourceSlide.transition));
    }

    // Copy timing if present
    if (sourceSlide.timing) {
      newSlide['p:timing'] = JSON.parse(JSON.stringify(sourceSlide.timing));
    }

    return newSlide;
  }

  /**
   * Copy media files for a slide and update references
   */
  async copyMediaForSlide(sourceSlide, newSlide, slideRels, startRelId) {
    const slideRelsPath = sourceSlide.path.replace('slides/', 'slides/_rels/') + '.rels';
    const sourceRels = await this.parseXmlFile(this.sourceZip, slideRelsPath);

    if (!sourceRels) return startRelId;

    const rels = this.normalizeRels(sourceRels);
    let nextRelId = startRelId;
    const relIdMapping = new Map();

    for (const rel of rels) {
      // Skip layout relationship (we handle it separately)
      if (rel.Type?.includes('slideLayout')) continue;

      // Skip notes (handled separately)
      if (rel.Type?.includes('notesSlide')) continue;

      const targetPath = rel.Target;
      let sourcePath = '';

      if (targetPath.startsWith('../')) {
        sourcePath = `ppt/${targetPath.replace('../', '')}`;
      } else if (targetPath.startsWith('/')) {
        sourcePath = targetPath.substring(1);
      } else {
        sourcePath = `ppt/slides/${targetPath}`;
      }

      // Check if file exists in source
      const sourceFile = this.sourceZip.file(sourcePath);
      if (sourceFile) {
        // Copy to output
        const content = await sourceFile.async('nodebuffer');
        this.outputZip.file(sourcePath, content);

        // Register content type
        const ext = sourcePath.split('.').pop()?.toLowerCase();
        const mimeType = this.getMimeType(ext);
        if (mimeType) {
          await this.registerContentType(sourcePath, mimeType);
        }
      }

      // Add relationship with new ID
      const newRelId = `rId${slideRels.length + 1}`;
      relIdMapping.set(rel.Id, newRelId);

      slideRels.push({
        Id: newRelId,
        Type: rel.Type,
        Target: rel.Target,
        TargetMode: rel.TargetMode
      });
    }

    // Update r:id references in the slide content
    this.updateRelIds(newSlide, relIdMapping);

    return nextRelId;
  }

  /**
   * Update relationship IDs in slide content
   */
  updateRelIds(obj, relIdMapping) {
    if (!obj || typeof obj !== 'object') return;

    if (Array.isArray(obj)) {
      obj.forEach(item => this.updateRelIds(item, relIdMapping));
      return;
    }

    for (const [key, value] of Object.entries(obj)) {
      if (key === '$' && value && typeof value === 'object') {
        // Check for r:id, r:embed, r:link attributes
        for (const [attrKey, attrValue] of Object.entries(value)) {
          if ((attrKey === 'r:id' || attrKey === 'r:embed' || attrKey === 'r:link') &&
              typeof attrValue === 'string' && relIdMapping.has(attrValue)) {
            value[attrKey] = relIdMapping.get(attrValue);
          }
        }
      } else if (value && typeof value === 'object') {
        this.updateRelIds(value, relIdMapping);
      }
    }
  }

  /**
   * Post-migration cleanup
   */
  async postMigrationCleanup() {
    // Ensure Content_Types.xml is complete
    await this.syncContentTypes();

    // Remove orphaned media files (optional - keeping for safety)
    // await this.removeOrphanedMedia();
  }

  /**
   * Sync Content_Types.xml
   */
  async syncContentTypes() {
    const contentTypesXml = await this.parseXmlFile(this.outputZip, '[Content_Types].xml');
    if (!contentTypesXml) return;

    const types = contentTypesXml['Types'] || contentTypesXml;

    // Ensure required extensions are present
    const defaultTypes = [
      { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' },
      { Extension: 'xml', ContentType: 'application/xml' },
      { Extension: 'png', ContentType: 'image/png' },
      { Extension: 'jpg', ContentType: 'image/jpeg' },
      { Extension: 'jpeg', ContentType: 'image/jpeg' },
      { Extension: 'gif', ContentType: 'image/gif' },
      { Extension: 'svg', ContentType: 'image/svg+xml' },
      { Extension: 'emf', ContentType: 'image/x-emf' },
      { Extension: 'wmf', ContentType: 'image/x-wmf' }
    ];

    let defaults = types['Default'] || [];
    if (!Array.isArray(defaults)) defaults = [defaults];

    for (const dt of defaultTypes) {
      const exists = defaults.some(d => d.$?.Extension === dt.Extension);
      if (!exists) {
        defaults.push({ $: dt });
      }
    }

    types['Default'] = defaults;

    await this.writeXmlFile(this.outputZip, '[Content_Types].xml', { Types: types });
  }

  /**
   * Register a file in Content_Types.xml
   */
  async registerContentType(filePath, contentType) {
    const contentTypesXml = await this.parseXmlFile(this.outputZip, '[Content_Types].xml');
    if (!contentTypesXml) return;

    const types = contentTypesXml['Types'] || contentTypesXml;

    // Normalize path for comparison
    const partName = filePath.startsWith('/') ? filePath : `/${filePath}`;

    let overrides = types['Override'] || [];
    if (!Array.isArray(overrides)) overrides = overrides ? [overrides] : [];

    // Check if already registered
    const exists = overrides.some(o => o.$?.PartName === partName);
    if (!exists) {
      overrides.push({
        $: {
          PartName: partName,
          ContentType: contentType
        }
      });
      types['Override'] = overrides;
      await this.writeXmlFile(this.outputZip, '[Content_Types].xml', { Types: types });
    }
  }

  /**
   * Generate migration report
   */
  generateReport() {
    const report = {
      source: {
        slideCount: this.sourceAnalysis.slides.length,
        slideSize: this.sourceAnalysis.slideSize,
        mediaCount: this.sourceAnalysis.mediaFiles.length
      },
      template: {
        layoutCount: this.templateAnalysis.layouts.length,
        slideSize: this.templateAnalysis.slideSize
      },
      migration: {
        slidesMigrated: this.migrationPlan.length,
        mappings: this.migrationPlan.map(p => ({
          slide: p.slideNumber,
          title: p.slideTitle,
          sourceLayout: p.sourceLayout,
          targetLayout: p.targetLayout?.name,
          matchType: p.matchType
        }))
      },
      warnings: this.warnings,
      slidesWithNotes: this.sourceAnalysis.slides.filter(s => s.notes).length
    };

    return report;
  }

  // ============ Utility Methods ============

  /**
   * Parse XML file from ZIP
   */
  async parseXmlFile(zip, path) {
    const file = zip.file(path);
    if (!file) return null;

    const content = await file.async('string');
    try {
      return await parseStringPromise(content, XML_PARSER_OPTIONS);
    } catch (e) {
      console.error(`Failed to parse ${path}:`, e.message);
      return null;
    }
  }

  /**
   * Write XML file to ZIP
   */
  async writeXmlFile(zip, path, obj) {
    const builder = new Builder(XML_BUILDER_OPTIONS);
    let xml = builder.buildObject(obj);

    // Add XML declaration if missing
    if (!xml.startsWith('<?xml')) {
      xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
    }

    zip.file(path, xml);
  }

  /**
   * Write relationships file
   */
  async writeRelsFile(zip, path, rels) {
    const relsObj = {
      Relationships: {
        $: {
          'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
        },
        Relationship: rels.map(r => ({
          $: {
            Id: r.Id,
            Type: r.Type,
            Target: r.Target,
            ...(r.TargetMode ? { TargetMode: r.TargetMode } : {})
          }
        }))
      }
    };

    await this.writeXmlFile(zip, path, relsObj);
  }

  /**
   * Normalize relationships from parsed XML
   */
  normalizeRels(relsXml) {
    if (!relsXml) return [];

    const relationships = relsXml.Relationships || relsXml;
    let rels = relationships.Relationship || relationships['Relationship'] || [];

    if (!Array.isArray(rels)) rels = rels ? [rels] : [];

    return rels.map(r => ({
      Id: r.$?.Id || r.Id,
      Type: r.$?.Type || r.Type,
      Target: r.$?.Target || r.Target,
      TargetMode: r.$?.TargetMode || r.TargetMode
    })).filter(r => r.Id && r.Type);
  }

  /**
   * Get MIME type for file extension
   */
  getMimeType(ext) {
    const mimeTypes = {
      'png': 'image/png',
      'jpg': 'image/jpeg',
      'jpeg': 'image/jpeg',
      'gif': 'image/gif',
      'svg': 'image/svg+xml',
      'emf': 'image/x-emf',
      'wmf': 'image/x-wmf',
      'xml': 'application/xml',
      'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    };
    return mimeTypes[ext] || null;
  }

  /**
   * Get migration plan for preview
   */
  getMigrationPlan() {
    return this.migrationPlan;
  }
}

export default PPTXMigrator;
