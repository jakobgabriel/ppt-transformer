import JSZip from 'jszip';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

/**
 * PPTX Template Migration Library
 * Migrates PowerPoint presentations to new templates at the OOXML level
 */

const XML_PARSER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  preserveOrder: true,
  parseAttributeValue: false,
  trimValues: false,
  processEntities: true,
  parseTagValue: false
};

const XML_BUILDER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  preserveOrder: true,
  format: true,
  indentBy: '  ',
  suppressEmptyNode: false,
  processEntities: true
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
    this.parser = new XMLParser(XML_PARSER_OPTIONS);
    this.builder = new XMLBuilder(XML_BUILDER_OPTIONS);
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
      const sldSz = this.findElement(presentationXml, 'p:sldSz');
      if (sldSz) {
        const attrs = this.getAttributes(sldSz);
        analysis.slideSize = {
          cx: parseInt(attrs.cx || 9144000),
          cy: parseInt(attrs.cy || 6858000),
          type: attrs.type || 'custom'
        };
      }
    }

    // Get all layout files
    const layoutFiles = Object.keys(zip.files).filter(f =>
      f.startsWith('ppt/slideLayouts/') && f.endsWith('.xml') && !f.includes('_rels')
    );

    for (const layoutFile of layoutFiles) {
      const layoutXml = await this.parseXmlFile(zip, layoutFile);
      if (!layoutXml) continue;

      const sldLayout = this.findElement(layoutXml, 'p:sldLayout');
      const cSld = this.findElement(sldLayout, 'p:cSld');
      const layoutAttrs = this.getAttributes(sldLayout);
      const cSldAttrs = this.getAttributes(cSld);

      const layoutInfo = {
        file: layoutFile.split('/').pop(),
        path: layoutFile,
        name: cSldAttrs.name || 'Unnamed Layout',
        type: layoutAttrs.type || 'obj',
        placeholders: this.extractPlaceholders(cSld)
      };

      // Get the master reference
      const layoutRelsPath = layoutFile.replace('slideLayouts/', 'slideLayouts/_rels/') + '.rels';
      const layoutRels = await this.parseXmlFile(zip, layoutRelsPath);
      if (layoutRels) {
        const rels = this.parseRels(layoutRels);
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

    // Get slide size
    const sldSz = this.findElement(presentationXml, 'p:sldSz');
    if (sldSz) {
      const attrs = this.getAttributes(sldSz);
      analysis.slideSize = {
        cx: parseInt(attrs.cx || 9144000),
        cy: parseInt(attrs.cy || 6858000),
        type: attrs.type || 'custom'
      };
    }

    // Get slide list
    const sldIdLst = this.findElement(presentationXml, 'p:sldIdLst');
    const slideIds = sldIdLst ? this.findAllElements(sldIdLst, 'p:sldId') : [];

    // Parse presentation.xml.rels to map r:id to file paths
    const presRels = await this.parseXmlFile(zip, 'ppt/_rels/presentation.xml.rels');
    const relMap = new Map();
    if (presRels) {
      const rels = this.parseRels(presRels);
      rels.forEach(r => relMap.set(r.Id, r.Target));
    }

    // Analyze each slide
    for (let i = 0; i < slideIds.length; i++) {
      const slideId = slideIds[i];
      const attrs = this.getAttributes(slideId);
      const rId = attrs['r:id'];
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
    const slide = this.findElement(slideXml, 'p:sld');
    const cSld = this.findElement(slide, 'p:cSld');

    const slideInfo = {
      number: slideNumber,
      path: slidePath,
      rawXml: null, // Store raw XML string for copying
      layout: null,
      layoutName: null,
      layoutType: null,
      title: null,
      placeholderContent: [],
      freeformContent: [],
      notes: null,
      notesPath: null
    };

    // Store raw XML for direct copying
    const rawXmlFile = zip.file(slidePath);
    if (rawXmlFile) {
      slideInfo.rawXml = await rawXmlFile.async('string');
    }

    // Get layout reference from slide rels
    const slideRelsPath = slidePath.replace('slides/', 'slides/_rels/') + '.rels';
    const slideRels = await this.parseXmlFile(zip, slideRelsPath);
    slideInfo.rawRels = slideRels;

    if (slideRels) {
      const rels = this.parseRels(slideRels);
      const layoutRel = rels.find(r => r.Type?.includes('slideLayout'));
      if (layoutRel) {
        slideInfo.layout = layoutRel.Target;

        // Get layout details
        const layoutPath = layoutRel.Target.startsWith('..')
          ? `ppt/slideLayouts/${layoutRel.Target.split('/').pop()}`
          : layoutRel.Target;
        const layoutXml = await this.parseXmlFile(zip, layoutPath);
        if (layoutXml) {
          const layout = this.findElement(layoutXml, 'p:sldLayout');
          const layoutCSld = this.findElement(layout, 'p:cSld');
          const layoutAttrs = this.getAttributes(layout);
          const cSldAttrs = this.getAttributes(layoutCSld);
          slideInfo.layoutName = cSldAttrs.name || 'Unknown';
          slideInfo.layoutType = layoutAttrs.type || 'obj';
        }
      }

      // Check for notes
      const notesRel = rels.find(r => r.Type?.includes('notesSlide'));
      if (notesRel) {
        const notesPath = notesRel.Target.startsWith('..')
          ? `ppt/notesSlides/${notesRel.Target.split('/').pop()}`
          : notesRel.Target;
        const notesFile = zip.file(notesPath);
        if (notesFile) {
          slideInfo.notes = await notesFile.async('string');
          slideInfo.notesPath = notesPath;
        }
      }
    }

    // Extract content from spTree
    const spTree = this.findElement(cSld, 'p:spTree');
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
    const shapes = this.findAllElements(spTree, 'p:sp');
    for (const shape of shapes) {
      const nvSpPr = this.findElement(shape, 'p:nvSpPr');
      const nvPr = this.findElement(nvSpPr, 'p:nvPr');
      const ph = this.findElement(nvPr, 'p:ph');

      if (ph) {
        const phAttrs = this.getAttributes(ph);
        slideInfo.placeholderContent.push({
          type: 'shape',
          element: shape,
          phType: phAttrs.type || 'body',
          phIdx: phAttrs.idx
        });
      } else {
        slideInfo.freeformContent.push({
          type: 'shape',
          element: shape
        });
      }
    }

    // Process pictures
    const pictures = this.findAllElements(spTree, 'p:pic');
    for (const pic of pictures) {
      const nvPicPr = this.findElement(pic, 'p:nvPicPr');
      const nvPr = this.findElement(nvPicPr, 'p:nvPr');
      const ph = this.findElement(nvPr, 'p:ph');

      if (ph) {
        const phAttrs = this.getAttributes(ph);
        slideInfo.placeholderContent.push({
          type: 'picture',
          element: pic,
          phType: phAttrs.type || 'pic',
          phIdx: phAttrs.idx
        });
      } else {
        slideInfo.freeformContent.push({
          type: 'picture',
          element: pic
        });
      }
    }

    // Process graphic frames (charts, tables, diagrams)
    const graphicFrames = this.findAllElements(spTree, 'p:graphicFrame');
    for (const gf of graphicFrames) {
      slideInfo.freeformContent.push({
        type: 'graphicFrame',
        element: gf
      });
    }

    // Process groups
    const groups = this.findAllElements(spTree, 'p:grpSp');
    for (const grp of groups) {
      slideInfo.freeformContent.push({
        type: 'group',
        element: grp
      });
    }
  }

  /**
   * Extract placeholders from a layout or master
   */
  extractPlaceholders(cSld) {
    const placeholders = [];
    if (!cSld) return placeholders;

    const spTree = this.findElement(cSld, 'p:spTree');
    if (!spTree) return placeholders;

    const shapes = this.findAllElements(spTree, 'p:sp');
    for (const shape of shapes) {
      const nvSpPr = this.findElement(shape, 'p:nvSpPr');
      const nvPr = this.findElement(nvSpPr, 'p:nvPr');
      const ph = this.findElement(nvPr, 'p:ph');

      if (ph) {
        const phAttrs = this.getAttributes(ph);
        const spPr = this.findElement(shape, 'p:spPr');
        const xfrm = this.findElement(spPr, 'a:xfrm');
        const off = this.findElement(xfrm, 'a:off');
        const ext = this.findElement(xfrm, 'a:ext');
        const offAttrs = this.getAttributes(off);
        const extAttrs = this.getAttributes(ext);

        placeholders.push({
          type: phAttrs.type || 'body',
          idx: phAttrs.idx,
          position: {
            x: parseInt(offAttrs.x || 0),
            y: parseInt(offAttrs.y || 0)
          },
          size: {
            cx: parseInt(extAttrs.cx || 0),
            cy: parseInt(extAttrs.cy || 0)
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
        const txBody = this.findElement(content.element, 'p:txBody');
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

    const paragraphs = this.findAllElements(txBody, 'a:p');
    const textParts = [];

    for (const p of paragraphs) {
      const runs = this.findAllElements(p, 'a:r');
      for (const r of runs) {
        const t = this.findElement(r, 'a:t');
        if (t) {
          const text = this.getTextContent(t);
          if (text) textParts.push(text);
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
    const presXmlStr = await this.outputZip.file('ppt/presentation.xml').async('string');
    let presXml = this.parser.parse(presXmlStr);

    // Find and clear sldIdLst
    const presentation = this.findElement(presXml, 'p:presentation');
    const sldIdLst = this.findElement(presentation, 'p:sldIdLst');

    if (sldIdLst) {
      // Remove all p:sldId children
      if (Array.isArray(sldIdLst)) {
        for (const item of sldIdLst) {
          if (item['p:sldId']) {
            delete item['p:sldId'];
          }
        }
      } else if (sldIdLst['p:sldId']) {
        delete sldIdLst['p:sldId'];
      }
    }

    // Check slide size
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
    const updatedPresXml = this.buildXml(presXml);
    this.outputZip.file('ppt/presentation.xml', updatedPresXml);

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

    // Update presentation.xml.rels
    const presRelsPath = 'ppt/_rels/presentation.xml.rels';
    const presRelsStr = await this.outputZip.file(presRelsPath).async('string');
    let presRels = this.parser.parse(presRelsStr);

    const relationships = this.findElement(presRels, 'Relationships');
    if (relationships) {
      // Filter out slide relationships but keep layout and master refs
      const rels = this.findAllElements(relationships, 'Relationship');
      let maxId = 0;

      for (const rel of rels) {
        const attrs = this.getAttributes(rel);
        const id = parseInt((attrs.Id || '').replace('rId', ''));
        if (id > maxId) maxId = id;

        // Remove slide relationships
        if (attrs.Type?.includes('/slide') && !attrs.Type?.includes('slideLayout') && !attrs.Type?.includes('slideMaster')) {
          // Mark for removal by clearing
          if (rel[':@']) {
            rel[':@']['@_Id'] = '__REMOVE__';
          }
        }
      }

      this.nextRelId = maxId + 1;
    }

    // Rebuild rels without removed entries
    const updatedPresRels = this.buildXml(presRels);
    const cleanedRels = updatedPresRels.replace(/<Relationship[^>]*Id="__REMOVE__"[^>]*\/>/g, '');
    this.outputZip.file(presRelsPath, cleanedRels);
  }

  /**
   * Migrate all slides
   */
  async migrateSlides() {
    const slideIdEntries = [];

    for (let i = 0; i < this.migrationPlan.length; i++) {
      const planItem = this.migrationPlan[i];
      const sourceSlide = this.sourceAnalysis.slides.find(s => s.number === planItem.slideNumber);

      const slideId = await this.migrateSlide(sourceSlide, planItem.targetLayout, i + 1);
      slideIdEntries.push(slideId);
    }

    // Update presentation.xml with new slide list
    const presXmlStr = await this.outputZip.file('ppt/presentation.xml').async('string');

    // Use string manipulation to add slide IDs (more reliable for OOXML)
    let updatedPresXml = presXmlStr;

    // Build slide ID list XML
    const sldIdXml = slideIdEntries.map(s =>
      `<p:sldId id="${s.id}" r:id="${s.rId}"/>`
    ).join('');

    // Replace empty sldIdLst or insert slide IDs
    if (updatedPresXml.includes('<p:sldIdLst/>')) {
      updatedPresXml = updatedPresXml.replace('<p:sldIdLst/>', `<p:sldIdLst>${sldIdXml}</p:sldIdLst>`);
    } else if (updatedPresXml.includes('<p:sldIdLst>')) {
      updatedPresXml = updatedPresXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${sldIdXml}</p:sldIdLst>`);
    } else {
      // Insert after sldMasterIdLst
      updatedPresXml = updatedPresXml.replace(
        /<\/p:sldMasterIdLst>/,
        `</p:sldMasterIdLst><p:sldIdLst>${sldIdXml}</p:sldIdLst>`
      );
    }

    this.outputZip.file('ppt/presentation.xml', updatedPresXml);
  }

  /**
   * Migrate a single slide
   */
  async migrateSlide(sourceSlide, targetLayout, slideNumber) {
    const slideFileName = `slide${slideNumber}.xml`;
    const slidePath = `ppt/slides/${slideFileName}`;
    const slideRelsPath = `ppt/slides/_rels/${slideFileName}.rels`;

    // Copy the raw slide XML from source
    let slideXml = sourceSlide.rawXml;

    // Update any layout-specific references if needed
    // The slide content itself is preserved as-is

    // Write slide XML
    this.outputZip.file(slidePath, slideXml);

    // Create slide relationships
    const layoutRelPath = `../slideLayouts/${targetLayout.file}`;

    // Parse source rels to get media and other references
    let slideRelsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${layoutRelPath}"/>`;

    let nextRelId = 2;

    // Copy relationships from source (excluding layout)
    if (sourceSlide.rawRels) {
      const sourceRels = this.parseRels(sourceSlide.rawRels);

      for (const rel of sourceRels) {
        // Skip layout relationship (we use our own)
        if (rel.Type?.includes('slideLayout')) continue;

        // Copy the relationship
        const newRelId = `rId${nextRelId++}`;

        // Copy referenced file if it exists
        let targetPath = rel.Target;
        if (targetPath.startsWith('../')) {
          const sourcePath = `ppt/${targetPath.replace('../', '')}`;
          const sourceFile = this.sourceZip.file(sourcePath);
          if (sourceFile) {
            const content = await sourceFile.async('nodebuffer');
            this.outputZip.file(sourcePath, content);

            // Register in Content_Types if needed
            await this.ensureContentType(sourcePath);
          }
        }

        const targetModeAttr = rel.TargetMode ? ` TargetMode="${rel.TargetMode}"` : '';
        slideRelsContent += `\n<Relationship Id="${newRelId}" Type="${rel.Type}" Target="${rel.Target}"${targetModeAttr}/>`;

        // Update the slide XML with new rel ID if needed
        if (rel.Id !== newRelId) {
          slideXml = slideXml.replace(new RegExp(`r:id="${rel.Id}"`, 'g'), `r:id="${newRelId}"`);
          slideXml = slideXml.replace(new RegExp(`r:embed="${rel.Id}"`, 'g'), `r:embed="${newRelId}"`);
          slideXml = slideXml.replace(new RegExp(`r:link="${rel.Id}"`, 'g'), `r:link="${newRelId}"`);
        }
      }
    }

    slideRelsContent += '\n</Relationships>';

    // Update the slide file with corrected rel IDs
    this.outputZip.file(slidePath, slideXml);

    // Write slide rels
    this.outputZip.file(slideRelsPath, slideRelsContent);

    // Handle notes if present
    if (sourceSlide.notes && sourceSlide.notesPath) {
      const notesFileName = `notesSlide${slideNumber}.xml`;
      const notesPath = `ppt/notesSlides/${notesFileName}`;
      const notesRelsPath = `ppt/notesSlides/_rels/${notesFileName}.rels`;

      // Update notes to reference new slide
      let notesXml = sourceSlide.notes;
      // Notes reference the slide - update if path changed
      notesXml = notesXml.replace(/Target="[^"]*slide\d+\.xml"/g, `Target="../slides/${slideFileName}"`);

      this.outputZip.file(notesPath, notesXml);

      // Notes rels
      const notesRelsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${slideFileName}"/>
</Relationships>`;

      this.outputZip.file(notesRelsPath, notesRelsContent);

      // Add notes relationship to slide rels
      slideRelsContent = slideRelsContent.replace(
        '</Relationships>',
        `<Relationship Id="rId${nextRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/${notesFileName}"/>
</Relationships>`
      );
      this.outputZip.file(slideRelsPath, slideRelsContent);

      await this.ensureContentType(notesPath, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
    }

    // Register slide in Content_Types
    await this.ensureContentType(slidePath, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');

    // Add to presentation.xml.rels
    const presRelsPath = 'ppt/_rels/presentation.xml.rels';
    let presRelsStr = await this.outputZip.file(presRelsPath).async('string');

    const slideRelId = `rId${this.nextRelId++}`;

    // Add new relationship
    presRelsStr = presRelsStr.replace(
      '</Relationships>',
      `<Relationship Id="${slideRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/${slideFileName}"/>
</Relationships>`
    );

    this.outputZip.file(presRelsPath, presRelsStr);

    return {
      id: this.nextSlideId++,
      rId: slideRelId
    };
  }

  /**
   * Post-migration cleanup
   */
  async postMigrationCleanup() {
    await this.syncContentTypes();
  }

  /**
   * Sync Content_Types.xml
   */
  async syncContentTypes() {
    let contentTypesStr = await this.outputZip.file('[Content_Types].xml').async('string');

    // Ensure common extensions are present
    const defaultTypes = [
      { ext: 'rels', type: 'application/vnd.openxmlformats-package.relationships+xml' },
      { ext: 'xml', type: 'application/xml' },
      { ext: 'png', type: 'image/png' },
      { ext: 'jpg', type: 'image/jpeg' },
      { ext: 'jpeg', type: 'image/jpeg' },
      { ext: 'gif', type: 'image/gif' },
      { ext: 'emf', type: 'image/x-emf' },
      { ext: 'wmf', type: 'image/x-wmf' }
    ];

    for (const dt of defaultTypes) {
      if (!contentTypesStr.includes(`Extension="${dt.ext}"`)) {
        contentTypesStr = contentTypesStr.replace(
          '<Types ',
          `<Types ><Default Extension="${dt.ext}" ContentType="${dt.type}"/`
        ).replace('><Default', '>\n<Default');
      }
    }

    this.outputZip.file('[Content_Types].xml', contentTypesStr);
  }

  /**
   * Ensure a file is registered in Content_Types.xml
   */
  async ensureContentType(filePath, contentType = null) {
    let contentTypesStr = await this.outputZip.file('[Content_Types].xml').async('string');

    const partName = filePath.startsWith('/') ? filePath : `/${filePath}`;

    if (contentTypesStr.includes(`PartName="${partName}"`)) {
      return; // Already registered
    }

    if (!contentType) {
      // Guess from extension
      const ext = filePath.split('.').pop()?.toLowerCase();
      const typeMap = {
        'xml': 'application/xml',
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif': 'image/gif',
        'emf': 'image/x-emf',
        'wmf': 'image/x-wmf',
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      };
      contentType = typeMap[ext];
    }

    if (contentType) {
      contentTypesStr = contentTypesStr.replace(
        '</Types>',
        `<Override PartName="${partName}" ContentType="${contentType}"/>
</Types>`
      );
      this.outputZip.file('[Content_Types].xml', contentTypesStr);
    }
  }

  /**
   * Generate migration report
   */
  generateReport() {
    return {
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
  }

  // ============ XML Helper Methods ============

  /**
   * Parse XML file from ZIP
   */
  async parseXmlFile(zip, path) {
    const file = zip.file(path);
    if (!file) return null;

    const content = await file.async('string');
    try {
      return this.parser.parse(content);
    } catch (e) {
      console.error(`Failed to parse ${path}:`, e.message);
      return null;
    }
  }

  /**
   * Build XML string from parsed object
   */
  buildXml(obj) {
    let xml = this.builder.build(obj);
    if (!xml.startsWith('<?xml')) {
      xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
    }
    return xml;
  }

  /**
   * Find an element in parsed XML (preserveOrder format)
   * Returns the whole object containing both the element and its :@ attributes
   */
  findElement(parent, tagName) {
    if (!parent) return null;

    // If parent is already a wrapped element, search in its children
    if (parent._children !== undefined) {
      return this.findElement(parent._children, tagName);
    }

    if (Array.isArray(parent)) {
      for (const item of parent) {
        // Check if this item has the tag we're looking for
        if (item && item[tagName] !== undefined) {
          // Return an object with the children and attributes
          return { _children: item[tagName], _attrs: item[':@'] || {} };
        }
        // Recursively search in children of this item
        if (item && typeof item === 'object') {
          for (const key of Object.keys(item)) {
            if (key === ':@' || key === '#text') continue;
            const found = this.findElement(item[key], tagName);
            if (found) return found;
          }
        }
      }
      return null;
    }

    if (typeof parent === 'object' && parent !== null) {
      if (parent[tagName] !== undefined) {
        return { _children: parent[tagName], _attrs: parent[':@'] || {} };
      }

      for (const key of Object.keys(parent)) {
        if (key === ':@' || key === '#text') continue;
        const found = this.findElement(parent[key], tagName);
        if (found) return found;
      }
    }

    return null;
  }

  /**
   * Find all elements with a given tag name
   */
  findAllElements(parent, tagName) {
    const results = [];
    if (!parent) return results;

    const search = (node) => {
      if (!node) return;

      // If node is a wrapped element, search in its children
      if (node._children !== undefined) {
        search(node._children);
        return;
      }

      if (Array.isArray(node)) {
        for (const item of node) {
          if (!item || typeof item !== 'object') continue;

          if (item[tagName] !== undefined) {
            // Push an object with children and attributes
            results.push({ _children: item[tagName], _attrs: item[':@'] || {} });
          }
          // Search deeper in all children
          for (const key of Object.keys(item)) {
            if (key === ':@' || key === '#text') continue;
            search(item[key]);
          }
        }
      } else if (typeof node === 'object') {
        if (node[tagName] !== undefined) {
          results.push({ _children: node[tagName], _attrs: node[':@'] || {} });
        }
        for (const key of Object.keys(node)) {
          if (key === ':@' || key === '#text') continue;
          search(node[key]);
        }
      }
    };

    search(parent);
    return results;
  }

  /**
   * Get attributes from an element (preserveOrder format)
   */
  getAttributes(element) {
    if (!element) return {};

    // If it's our wrapped format with _attrs
    if (element._attrs) {
      const attrs = {};
      for (const [key, value] of Object.entries(element._attrs)) {
        const cleanKey = key.startsWith('@_') ? key.substring(2) : key;
        attrs[cleanKey] = value;
      }
      return attrs;
    }

    // In preserveOrder mode, attributes are in :@ property
    if (element[':@']) {
      const attrs = {};
      for (const [key, value] of Object.entries(element[':@'])) {
        const cleanKey = key.startsWith('@_') ? key.substring(2) : key;
        attrs[cleanKey] = value;
      }
      return attrs;
    }

    // Fallback: check for @_ prefixed properties directly
    const attrs = {};
    if (typeof element === 'object') {
      for (const [key, value] of Object.entries(element)) {
        if (key.startsWith('@_')) {
          attrs[key.substring(2)] = value;
        }
      }
    }
    return attrs;
  }

  /**
   * Get children of an element
   */
  getChildren(element) {
    if (!element) return [];
    if (element._children !== undefined) return element._children;
    return element;
  }

  /**
   * Get text content from an element
   */
  getTextContent(element) {
    if (!element) return '';
    if (typeof element === 'string') return element;

    const children = this.getChildren(element);

    if (Array.isArray(children)) {
      for (const item of children) {
        if (item && item['#text'] !== undefined) {
          const textNodes = item['#text'];
          if (Array.isArray(textNodes)) {
            for (const t of textNodes) {
              if (t && t['#text']) return t['#text'];
              if (typeof t === 'string') return t;
            }
          }
          return textNodes;
        }
      }
      return '';
    }

    if (children && children['#text']) return children['#text'];
    return '';
  }

  /**
   * Parse relationships from a .rels file
   */
  parseRels(relsXml) {
    if (!relsXml) return [];

    const relElements = this.findAllElements(relsXml, 'Relationship');

    return relElements.map(rel => {
      const attrs = this.getAttributes(rel);
      return {
        Id: attrs.Id,
        Type: attrs.Type,
        Target: attrs.Target,
        TargetMode: attrs.TargetMode
      };
    }).filter(r => r.Id && r.Type);
  }
}

export default PPTXMigrator;
