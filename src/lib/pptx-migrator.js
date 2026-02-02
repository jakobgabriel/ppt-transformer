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
   * Prepend title slide and apply template background
   * This mode keeps existing slides mostly intact but:
   * 1. Adds the template's title slide at the beginning
   * 2. Applies the template's background/master styling to all slides
   */
  async prependTitleSlide(sourceBuffer, templateBuffer, titleText = null) {
    this.sourceZip = await JSZip.loadAsync(sourceBuffer);
    this.templateZip = await JSZip.loadAsync(templateBuffer);

    // Start output from source (preserve original content)
    this.outputZip = await JSZip.loadAsync(sourceBuffer);

    // Analyze both
    this.templateAnalysis = await this.analyzeTemplate(this.templateZip);
    this.sourceAnalysis = await this.analyzeSource(this.sourceZip);

    // Copy template's theme, slide master, and layouts to output
    await this.copyTemplateDesign();

    // Get the title slide from template
    const titleSlideContent = await this.extractTemplateTitleSlide(titleText);

    // Insert title slide at the beginning and renumber existing slides
    await this.insertTitleSlideAtBeginning(titleSlideContent);

    // Update all existing slides to use template's slide master/layouts
    await this.applyTemplateBackgroundToSlides();

    // Sync content types
    await this.syncContentTypes();

    // Generate output
    const outputBuffer = await this.outputZip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    return {
      buffer: outputBuffer,
      report: {
        mode: 'prepend-title',
        originalSlideCount: this.sourceAnalysis.slides.length,
        newSlideCount: this.sourceAnalysis.slides.length + 1,
        titleSlideAdded: true,
        backgroundApplied: true,
        warnings: this.warnings
      }
    };
  }

  /**
   * Copy template's design elements (theme, master, layouts) to output
   * Also updates presentation.xml to reference the template's slide master
   */
  async copyTemplateDesign() {
    // First, remove existing design elements from output to avoid conflicts
    const outputFiles = Object.keys(this.outputZip.files);
    for (const file of outputFiles) {
      if (file.startsWith('ppt/slideMasters/') ||
          file.startsWith('ppt/slideLayouts/') ||
          file.startsWith('ppt/theme/')) {
        this.outputZip.remove(file);
      }
    }

    // Copy ALL template media files first (images, etc.)
    const mediaFiles = Object.keys(this.templateZip.files).filter(f =>
      f.startsWith('ppt/media/')
    );
    for (const file of mediaFiles) {
      const content = await this.templateZip.file(file).async('nodebuffer');
      this.outputZip.file(file, content);
    }

    // Copy theme folder
    const themeFiles = Object.keys(this.templateZip.files).filter(f =>
      f.startsWith('ppt/theme/')
    );
    for (const file of themeFiles) {
      const content = await this.templateZip.file(file).async('nodebuffer');
      this.outputZip.file(file, content);
    }

    // Copy slide masters (including _rels)
    const masterFiles = Object.keys(this.templateZip.files).filter(f =>
      f.startsWith('ppt/slideMasters/')
    );
    for (const file of masterFiles) {
      const content = await this.templateZip.file(file).async('nodebuffer');
      this.outputZip.file(file, content);
    }

    // Copy slide layouts (including _rels)
    const layoutFiles = Object.keys(this.templateZip.files).filter(f =>
      f.startsWith('ppt/slideLayouts/')
    );
    for (const file of layoutFiles) {
      const content = await this.templateZip.file(file).async('nodebuffer');
      this.outputZip.file(file, content);
    }

    // Get template's presentation.xml.rels
    const templatePresRels = await this.templateZip.file('ppt/_rels/presentation.xml.rels').async('string');
    let outputPresRels = await this.outputZip.file('ppt/_rels/presentation.xml.rels').async('string');

    // Remove existing slideMaster, theme, and notesMaster relationships from output
    outputPresRels = outputPresRels.replace(/<Relationship[^>]*Type="[^"]*slideMaster"[^>]*\/>\s*/g, '');
    outputPresRels = outputPresRels.replace(/<Relationship[^>]*Type="[^"]*theme"[^>]*\/>\s*/g, '');
    outputPresRels = outputPresRels.replace(/<Relationship[^>]*Type="[^"]*notesMaster"[^>]*\/>\s*/g, '');

    // Find max rId in output rels
    let maxRId = 1;
    const rIdMatches = [...outputPresRels.matchAll(/Id="rId(\d+)"/g)];
    for (const m of rIdMatches) {
      const id = parseInt(m[1]);
      if (id > maxRId) maxRId = id;
    }

    // Extract ALL master, theme, and notesMaster relationships from template
    const masterRelMatches = [...templatePresRels.matchAll(/<Relationship[^>]*Type="[^"]*slideMaster"[^>]*\/>/g)];
    const themeRelMatches = [...templatePresRels.matchAll(/<Relationship[^>]*Type="[^"]*theme"[^>]*\/>/g)];
    const notesMasterMatches = [...templatePresRels.matchAll(/<Relationship[^>]*Type="[^"]*notesMaster"[^>]*\/>/g)];

    // Add template's relationships with new IDs
    let newRelsXml = '';
    const masterRIdMap = new Map(); // Map old rId to new rId

    for (const match of masterRelMatches) {
      const relXml = match[0];
      const oldRIdMatch = relXml.match(/Id="(rId\d+)"/);
      const targetMatch = relXml.match(/Target="([^"]+)"/);

      if (oldRIdMatch && targetMatch) {
        const newRId = `rId${++maxRId}`;
        masterRIdMap.set(oldRIdMatch[1], newRId);
        newRelsXml += `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="${targetMatch[1]}"/>\n`;
      }
    }

    for (const match of themeRelMatches) {
      const relXml = match[0];
      const targetMatch = relXml.match(/Target="([^"]+)"/);

      if (targetMatch) {
        const newRId = `rId${++maxRId}`;
        newRelsXml += `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="${targetMatch[1]}"/>\n`;
      }
    }

    // Copy notesMaster if template has one
    if (notesMasterMatches.length > 0) {
      const notesMasterFiles = Object.keys(this.templateZip.files).filter(f =>
        f.startsWith('ppt/notesMasters/')
      );
      for (const file of notesMasterFiles) {
        const content = await this.templateZip.file(file).async('nodebuffer');
        this.outputZip.file(file, content);
      }

      for (const match of notesMasterMatches) {
        const relXml = match[0];
        const targetMatch = relXml.match(/Target="([^"]+)"/);
        if (targetMatch) {
          const newRId = `rId${++maxRId}`;
          newRelsXml += `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="${targetMatch[1]}"/>\n`;
        }
      }
    }

    outputPresRels = outputPresRels.replace('</Relationships>', `${newRelsXml}</Relationships>`);
    outputPresRels = outputPresRels.replace(/\n\s*\n/g, '\n');
    this.outputZip.file('ppt/_rels/presentation.xml.rels', outputPresRels);

    // Update presentation.xml's sldMasterIdLst to use template's masters
    const templatePresXml = await this.templateZip.file('ppt/presentation.xml').async('string');
    let outputPresXml = await this.outputZip.file('ppt/presentation.xml').async('string');

    // Extract sldMasterIdLst from template
    const templateMasterListMatch = templatePresXml.match(/<p:sldMasterIdLst>([\s\S]*?)<\/p:sldMasterIdLst>/);

    if (templateMasterListMatch) {
      // Update the rIds in the master list to use new IDs
      let newMasterList = templateMasterListMatch[1];

      for (const [oldRId, newRId] of masterRIdMap) {
        newMasterList = newMasterList.replace(new RegExp(`r:id="${oldRId}"`, 'g'), `r:id="${newRId}"`);
      }

      // Replace or insert sldMasterIdLst in output
      if (outputPresXml.includes('<p:sldMasterIdLst>')) {
        outputPresXml = outputPresXml.replace(
          /<p:sldMasterIdLst>[\s\S]*?<\/p:sldMasterIdLst>/,
          `<p:sldMasterIdLst>${newMasterList}</p:sldMasterIdLst>`
        );
      } else {
        // Insert after opening p:presentation tag
        outputPresXml = outputPresXml.replace(
          /(<p:presentation[^>]*>)/,
          `$1\n<p:sldMasterIdLst>${newMasterList}</p:sldMasterIdLst>`
        );
      }

      this.outputZip.file('ppt/presentation.xml', outputPresXml);
    }

    // Update Content_Types.xml
    let contentTypes = await this.outputZip.file('[Content_Types].xml').async('string');

    // Remove old master/layout/theme overrides
    contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/ppt\/slideMasters\/[^"]*"[^>]*\/>\s*/g, '');
    contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/ppt\/slideLayouts\/[^"]*"[^>]*\/>\s*/g, '');
    contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/ppt\/theme\/[^"]*"[^>]*\/>\s*/g, '');
    contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/ppt\/notesMasters\/[^"]*"[^>]*\/>\s*/g, '');

    // Add content types for template's files
    let newOverrides = '';

    for (const file of masterFiles) {
      if (file.endsWith('.xml') && !file.includes('_rels')) {
        newOverrides += `<Override PartName="/${file}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>\n`;
      }
    }

    for (const file of layoutFiles) {
      if (file.endsWith('.xml') && !file.includes('_rels')) {
        newOverrides += `<Override PartName="/${file}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>\n`;
      }
    }

    for (const file of themeFiles) {
      if (file.endsWith('.xml') && !file.includes('_rels')) {
        newOverrides += `<Override PartName="/${file}" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>\n`;
      }
    }

    // Add media content types if not present
    for (const file of mediaFiles) {
      const ext = file.split('.').pop()?.toLowerCase();
      if (ext && !contentTypes.includes(`extension="${ext}"`)) {
        const mimeTypes = {
          'png': 'image/png',
          'jpg': 'image/jpeg',
          'jpeg': 'image/jpeg',
          'gif': 'image/gif',
          'emf': 'image/x-emf',
          'wmf': 'image/x-wmf',
          'svg': 'image/svg+xml'
        };
        if (mimeTypes[ext]) {
          // Check if Default for this extension exists
          if (!contentTypes.includes(`Extension="${ext}"`)) {
            newOverrides = `<Default Extension="${ext}" ContentType="${mimeTypes[ext]}"/>\n` + newOverrides;
          }
        }
      }
    }

    contentTypes = contentTypes.replace('</Types>', `${newOverrides}</Types>`);
    contentTypes = contentTypes.replace(/\n\s*\n/g, '\n');
    this.outputZip.file('[Content_Types].xml', contentTypes);
  }

  /**
   * Extract the title slide from the template
   */
  async extractTemplateTitleSlide(customTitle) {
    // Find the first slide in template (assuming it's the title slide)
    // Or find a slide using the title layout
    const templatePresXml = await this.templateZip.file('ppt/presentation.xml').async('string');

    // Get slide list from template
    const sldIdMatch = templatePresXml.match(/<p:sldId[^>]*r:id="([^"]+)"[^>]*\/>/);
    if (!sldIdMatch) {
      throw new Error('No slides found in template');
    }

    const firstSlideRId = sldIdMatch[1];

    // Get the slide path from presentation.xml.rels
    const presRels = await this.templateZip.file('ppt/_rels/presentation.xml.rels').async('string');
    const slideRelMatch = presRels.match(new RegExp(`<Relationship[^>]*Id="${firstSlideRId}"[^>]*Target="([^"]+)"[^>]*/>`));

    if (!slideRelMatch) {
      throw new Error('Could not find title slide relationship');
    }

    const slidePath = slideRelMatch[1].startsWith('slides/')
      ? `ppt/${slideRelMatch[1]}`
      : `ppt/slides/${slideRelMatch[1].replace('../', '')}`;

    // Read the title slide XML
    let slideXml = await this.templateZip.file(slidePath).async('string');

    // If custom title provided, update the title text
    if (customTitle) {
      slideXml = this.updateTitleText(slideXml, customTitle);
    }

    // Get the slide's rels file
    const slideRelsPath = slidePath.replace('slides/', 'slides/_rels/') + '.rels';
    const slideRelsFile = this.templateZip.file(slideRelsPath);
    const slideRels = slideRelsFile ? await slideRelsFile.async('string') : null;

    return {
      xml: slideXml,
      rels: slideRels,
      originalPath: slidePath
    };
  }

  /**
   * Update the title text in a slide
   */
  updateTitleText(slideXml, newTitle) {
    // Find title placeholder and update text
    // Look for shape with <p:ph type="ctrTitle"/> or <p:ph type="title"/>
    // This is a simplified approach - just replace the first title-like text

    // Match the txBody inside a title placeholder shape
    const titleShapeRegex = /(<p:sp[^>]*>[\s\S]*?<p:ph[^>]*type="(?:ctrTitle|title)"[^>]*\/>[\s\S]*?<p:txBody>)([\s\S]*?)(<\/p:txBody>[\s\S]*?<\/p:sp>)/;

    const match = slideXml.match(titleShapeRegex);
    if (match) {
      // Create new txBody content with the custom title
      const newTxBody = `<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE" dirty="0"/><a:t>${this.escapeXml(newTitle)}</a:t></a:r></a:p>`;
      slideXml = slideXml.replace(titleShapeRegex, `$1${newTxBody}$3`);
    }

    return slideXml;
  }

  /**
   * Escape XML special characters
   */
  escapeXml(text) {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  /**
   * Insert the title slide at the beginning and renumber existing slides
   */
  async insertTitleSlideAtBeginning(titleSlideContent) {
    // Step 1: Renumber all existing slides (slide1 -> slide2, etc.)
    const existingSlides = Object.keys(this.outputZip.files).filter(f =>
      f.match(/^ppt\/slides\/slide\d+\.xml$/)
    ).sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)/)[1]);
      const numB = parseInt(b.match(/slide(\d+)/)[1]);
      return numB - numA; // Reverse order to avoid overwriting
    });

    const existingRels = Object.keys(this.outputZip.files).filter(f =>
      f.match(/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/)
    ).sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)/)[1]);
      const numB = parseInt(b.match(/slide(\d+)/)[1]);
      return numB - numA;
    });

    // Renumber slides from highest to lowest to avoid conflicts
    for (const oldPath of existingSlides) {
      const num = parseInt(oldPath.match(/slide(\d+)/)[1]);
      const newNum = num + 1;
      const newPath = `ppt/slides/slide${newNum}.xml`;

      const content = await this.outputZip.file(oldPath).async('nodebuffer');
      this.outputZip.file(newPath, content);
      this.outputZip.remove(oldPath);
    }

    for (const oldPath of existingRels) {
      const num = parseInt(oldPath.match(/slide(\d+)/)[1]);
      const newNum = num + 1;
      const newPath = `ppt/slides/_rels/slide${newNum}.xml.rels`;

      let content = await this.outputZip.file(oldPath).async('string');

      // Update notesSlide references if present
      content = content.replace(
        /Target="\.\.\/notesSlides\/notesSlide(\d+)\.xml"/g,
        (match, noteNum) => `Target="../notesSlides/notesSlide${parseInt(noteNum) + 1}.xml"`
      );

      this.outputZip.file(newPath, content);
      this.outputZip.remove(oldPath);
    }

    // Also renumber notes slides if they exist
    const existingNotes = Object.keys(this.outputZip.files).filter(f =>
      f.match(/^ppt\/notesSlides\/notesSlide\d+\.xml$/)
    ).sort((a, b) => {
      const numA = parseInt(a.match(/notesSlide(\d+)/)[1]);
      const numB = parseInt(b.match(/notesSlide(\d+)/)[1]);
      return numB - numA;
    });

    const existingNotesRels = Object.keys(this.outputZip.files).filter(f =>
      f.match(/^ppt\/notesSlides\/_rels\/notesSlide\d+\.xml\.rels$/)
    ).sort((a, b) => {
      const numA = parseInt(a.match(/notesSlide(\d+)/)[1]);
      const numB = parseInt(b.match(/notesSlide(\d+)/)[1]);
      return numB - numA;
    });

    for (const oldPath of existingNotes) {
      const num = parseInt(oldPath.match(/notesSlide(\d+)/)[1]);
      const newNum = num + 1;
      const newPath = `ppt/notesSlides/notesSlide${newNum}.xml`;

      const content = await this.outputZip.file(oldPath).async('nodebuffer');
      this.outputZip.file(newPath, content);
      this.outputZip.remove(oldPath);
    }

    for (const oldPath of existingNotesRels) {
      const num = parseInt(oldPath.match(/notesSlide(\d+)/)[1]);
      const newNum = num + 1;
      const newPath = `ppt/notesSlides/_rels/notesSlide${newNum}.xml.rels`;

      let content = await this.outputZip.file(oldPath).async('string');

      // Update slide references
      content = content.replace(
        /Target="\.\.\/slides\/slide(\d+)\.xml"/g,
        (match, slideNum) => `Target="../slides/slide${parseInt(slideNum) + 1}.xml"`
      );

      this.outputZip.file(newPath, content);
      this.outputZip.remove(oldPath);
    }

    // Step 2: Add the title slide as slide1
    this.outputZip.file('ppt/slides/slide1.xml', titleSlideContent.xml);

    // Create/copy rels for title slide
    if (titleSlideContent.rels) {
      this.outputZip.file('ppt/slides/_rels/slide1.xml.rels', titleSlideContent.rels);
    } else {
      // Create minimal rels pointing to the title layout
      const titleLayout = this.templateAnalysis.layouts.find(l =>
        l.type === 'title' || l.name.toLowerCase().includes('titel')
      ) || this.templateAnalysis.layouts[0];

      const relsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/${titleLayout.file}"/>
</Relationships>`;
      this.outputZip.file('ppt/slides/_rels/slide1.xml.rels', relsContent);
    }

    // Copy any media referenced by the title slide from template
    if (titleSlideContent.rels) {
      const relsXml = await this.parseXmlFile({ file: (path) => ({ async: async () => titleSlideContent.rels }) }, 'dummy');
      // Actually parse the rels properly
      const relMatches = titleSlideContent.rels.matchAll(/<Relationship[^>]*Target="([^"]+)"[^>]*\/>/g);
      for (const match of relMatches) {
        const target = match[1];
        if (target.includes('media/') || target.includes('image')) {
          const mediaPath = target.startsWith('../')
            ? `ppt/${target.replace('../', '')}`
            : `ppt/slides/${target}`;

          const templateFile = this.templateZip.file(mediaPath);
          if (templateFile) {
            const content = await templateFile.async('nodebuffer');
            this.outputZip.file(mediaPath, content);
          }
        }
      }
    }

    // Step 3: Update presentation.xml with new slide order
    let presXml = await this.outputZip.file('ppt/presentation.xml').async('string');

    // Get current slide ID list
    const sldIdListMatch = presXml.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
    let maxSlideId = 256;

    if (sldIdListMatch) {
      // Find max slide ID
      const idMatches = sldIdListMatch[1].matchAll(/id="(\d+)"/g);
      for (const m of idMatches) {
        const id = parseInt(m[1]);
        if (id > maxSlideId) maxSlideId = id;
      }
    }

    // Step 4: Update presentation.xml.rels - renumber slide relationships and add new one
    let presRels = await this.outputZip.file('ppt/_rels/presentation.xml.rels').async('string');

    // Find all slide relationships and renumber them
    const slideRelRegex = /<Relationship[^>]*Id="(rId\d+)"[^>]*Type="[^"]*\/slide"[^>]*Target="slides\/slide(\d+)\.xml"[^>]*\/>/g;
    const slideRels = [...presRels.matchAll(slideRelRegex)];

    // Find max rId
    let maxRId = 1;
    const rIdMatches = presRels.matchAll(/Id="rId(\d+)"/g);
    for (const m of rIdMatches) {
      const id = parseInt(m[1]);
      if (id > maxRId) maxRId = id;
    }

    // Remove old slide relationships
    presRels = presRels.replace(/<Relationship[^>]*Type="[^"]*\/slide"[^>]*\/>/g, '');

    // Add new slide relationships (title slide + renumbered existing)
    const newTitleSlideRId = `rId${maxRId + 1}`;
    let newRelsXml = `<Relationship Id="${newTitleSlideRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>`;

    for (let i = 0; i < slideRels.length; i++) {
      const newRId = `rId${maxRId + 2 + i}`;
      const newSlideNum = i + 2; // Existing slides are now 2, 3, 4, ...
      newRelsXml += `\n<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newSlideNum}.xml"/>`;
    }

    presRels = presRels.replace('</Relationships>', `${newRelsXml}\n</Relationships>`);
    presRels = presRels.replace(/\n\s*\n/g, '\n'); // Clean up empty lines

    this.outputZip.file('ppt/_rels/presentation.xml.rels', presRels);

    // Update sldIdLst in presentation.xml
    const newSlideId = maxSlideId + 1;
    let newSldIdLst = `<p:sldId id="${newSlideId}" r:id="${newTitleSlideRId}"/>`;

    for (let i = 0; i < slideRels.length; i++) {
      const newRId = `rId${maxRId + 2 + i}`;
      newSldIdLst += `<p:sldId id="${maxSlideId + 2 + i}" r:id="${newRId}"/>`;
    }

    if (sldIdListMatch) {
      presXml = presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${newSldIdLst}</p:sldIdLst>`);
    } else {
      // Insert after sldMasterIdLst
      presXml = presXml.replace('</p:sldMasterIdLst>', `</p:sldMasterIdLst><p:sldIdLst>${newSldIdLst}</p:sldIdLst>`);
    }

    this.outputZip.file('ppt/presentation.xml', presXml);

    // Update Content_Types.xml
    let contentTypes = await this.outputZip.file('[Content_Types].xml').async('string');

    // Remove old slide overrides
    contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>\s*/g, '');

    // Add new slide overrides
    let slideOverrides = '';
    for (let i = 1; i <= slideRels.length + 1; i++) {
      slideOverrides += `<Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n`;
    }

    contentTypes = contentTypes.replace('</Types>', `${slideOverrides}</Types>`);
    this.outputZip.file('[Content_Types].xml', contentTypes);
  }

  /**
   * Apply template background to all existing slides by updating their layout references
   */
  async applyTemplateBackgroundToSlides() {
    // Find a suitable content layout from template
    const contentLayout = this.templateAnalysis.layouts.find(l =>
      l.type === 'obj' || l.name.toLowerCase().includes('inhalt') || l.name.toLowerCase().includes('content')
    ) || this.templateAnalysis.layouts.find(l =>
      l.name.toLowerCase().includes('leer') || l.name.toLowerCase().includes('blank')
    ) || this.templateAnalysis.layouts[1] || this.templateAnalysis.layouts[0];

    // Update each existing slide's rels to point to template layout
    // Slides 2+ are the original slides (slide1 is now the title)
    const slideRelsFiles = Object.keys(this.outputZip.files).filter(f =>
      f.match(/^ppt\/slides\/_rels\/slide[2-9]\d*\.xml\.rels$/) ||
      f.match(/^ppt\/slides\/_rels\/slide[1-9]\d+\.xml\.rels$/)
    );

    for (const relsPath of slideRelsFiles) {
      let relsContent = await this.outputZip.file(relsPath).async('string');

      // Update layout reference to point to template layout
      const newLayoutTarget = `../slideLayouts/${contentLayout.file}`;
      relsContent = relsContent.replace(
        /(<Relationship[^>]*Type="[^"]*slideLayout"[^>]*Target=")[^"]*(")/g,
        `$1${newLayoutTarget}$2`
      );

      this.outputZip.file(relsPath, relsContent);
    }
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
   * Includes German layout names from Template_Transitionsphase_OESL.pptx
   */
  findLayoutByName(sourceName) {
    if (!sourceName) return null;

    const normalized = sourceName.toLowerCase().replace(/[^a-z0-9äöüß]/g, '');

    // Exact match first
    let match = this.templateAnalysis.layouts.find(l =>
      l.name.toLowerCase().replace(/[^a-z0-9äöüß]/g, '') === normalized
    );

    if (match) return match;

    // Fuzzy match - includes German layout names
    // Template layouts: Titelfolie, Titel und Inhalt, Abschnittsüberschrift, Titel, Titel+Text, Leer
    const fuzzyMappings = {
      // Title slide mappings
      'title': ['title', 'cover', 'opening', 'titelfolie', 'titel'],
      'titleslide': ['titleslide', 'corporatetitle', 'cover', 'titelfolie'],
      'titelfolie': ['titelfolie', 'titleslide', 'title', 'cover'],

      // Content slide mappings
      'twocontent': ['twocontent', 'twocolumn', 'comparison', 'zweiinhalte'],
      'titleandcontent': ['titleandcontent', 'titlecontent', 'content', 'titelundinhalt', 'titelinhalt'],
      'titelundinhalt': ['titelundinhalt', 'titleandcontent', 'content', 'titelinhalt'],
      'titeltext': ['titeltext', 'titletext', 'titelplustext'],

      // Section header mappings
      'sectionheader': ['sectionheader', 'section', 'sectionbreak', 'divider', 'abschnitt', 'abschnittsüberschrift'],
      'abschnitt': ['abschnitt', 'abschnittsüberschrift', 'sectionheader', 'section'],

      // Blank slide mappings
      'blank': ['blank', 'empty', 'leer'],
      'leer': ['leer', 'blank', 'empty']
    };

    for (const [key, values] of Object.entries(fuzzyMappings)) {
      if (normalized.includes(key) || values.some(v => normalized.includes(v))) {
        match = this.templateAnalysis.layouts.find(l => {
          const targetNorm = l.name.toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
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
   * Uses string manipulation to preserve XML structure
   */
  async prepareOutputPackage() {
    // Get presentation.xml as string
    let presXmlStr = await this.outputZip.file('ppt/presentation.xml').async('string');

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

    // Clear sldIdLst content using string replacement (preserve the tag structure)
    // Match <p:sldIdLst>...</p:sldIdLst> or <p:sldIdLst ...>...</p:sldIdLst>
    presXmlStr = presXmlStr.replace(
      /<p:sldIdLst[^>]*>[\s\S]*?<\/p:sldIdLst>/g,
      '<p:sldIdLst></p:sldIdLst>'
    );
    // Also handle self-closing variant
    presXmlStr = presXmlStr.replace(/<p:sldIdLst\s*\/>/g, '<p:sldIdLst></p:sldIdLst>');

    this.outputZip.file('ppt/presentation.xml', presXmlStr);

    // Remove existing template slides from the package
    const slideFiles = Object.keys(this.outputZip.files).filter(f =>
      f.startsWith('ppt/slides/') && f.endsWith('.xml') && !f.includes('_rels')
    );
    const slideRelsFiles = Object.keys(this.outputZip.files).filter(f =>
      f.startsWith('ppt/slides/_rels/')
    );

    // Also remove any template notes slides
    const notesFiles = Object.keys(this.outputZip.files).filter(f =>
      f.startsWith('ppt/notesSlides/') && f.endsWith('.xml') && !f.includes('_rels')
    );
    const notesRelsFiles = Object.keys(this.outputZip.files).filter(f =>
      f.startsWith('ppt/notesSlides/_rels/')
    );

    for (const file of [...slideFiles, ...slideRelsFiles, ...notesFiles, ...notesRelsFiles]) {
      this.outputZip.remove(file);
    }

    // Ensure the slides and notesSlides directories exist (add placeholder)
    // JSZip handles directories implicitly when files are added

    // Update presentation.xml.rels - remove slide relationships using string manipulation
    const presRelsPath = 'ppt/_rels/presentation.xml.rels';
    let presRelsStr = await this.outputZip.file(presRelsPath).async('string');

    // Find max relationship ID
    const idMatches = presRelsStr.matchAll(/Id="rId(\d+)"/g);
    let maxId = 0;
    for (const match of idMatches) {
      const id = parseInt(match[1]);
      if (id > maxId) maxId = id;
    }
    this.nextRelId = maxId + 1;

    // Remove slide relationships (but keep slideLayout and slideMaster)
    // Match relationships with Type containing /slide but not /slideLayout or /slideMaster
    presRelsStr = presRelsStr.replace(
      /<Relationship[^>]*Type="[^"]*\/slide"[^>]*\/>/g,
      ''
    );
    // Clean up any double newlines
    presRelsStr = presRelsStr.replace(/\n\s*\n/g, '\n');

    this.outputZip.file(presRelsPath, presRelsStr);
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
   * Preserves original relationship IDs to maintain all formatting and references
   */
  async migrateSlide(sourceSlide, targetLayout, slideNumber) {
    const slideFileName = `slide${slideNumber}.xml`;
    const slidePath = `ppt/slides/${slideFileName}`;
    const slideRelsPath = `ppt/slides/_rels/${slideFileName}.rels`;

    // Copy the raw slide XML and apply cleanups
    let slideXml = sourceSlide.rawXml;

    // Clean up the slide: remove footers, logos, expand text areas
    slideXml = this.cleanupSlideXml(slideXml);

    this.outputZip.file(slidePath, slideXml);

    // Get the original slide rels file as string and modify it
    const sourceSlideRelsPath = sourceSlide.path.replace('slides/', 'slides/_rels/') + '.rels';
    const sourceRelsFile = this.sourceZip.file(sourceSlideRelsPath);

    let slideRelsContent;
    let maxRelId = 1;

    if (sourceRelsFile) {
      // Copy the original rels file and only change the layout target
      slideRelsContent = await sourceRelsFile.async('string');

      // Find all relationship IDs to track the max
      const idMatches = slideRelsContent.matchAll(/Id="rId(\d+)"/g);
      for (const match of idMatches) {
        const id = parseInt(match[1]);
        if (id > maxRelId) maxRelId = id;
      }

      // Update the slideLayout relationship target to point to the new template layout
      const newLayoutTarget = `../slideLayouts/${targetLayout.file}`;
      slideRelsContent = slideRelsContent.replace(
        /(<Relationship[^>]*Type="[^"]*slideLayout"[^>]*Target=")[^"]*(")/g,
        `$1${newLayoutTarget}$2`
      );

      // Copy all referenced media files from source to output
      const sourceRels = this.parseRels(await this.parseXmlFile(this.sourceZip, sourceSlideRelsPath));
      for (const rel of sourceRels) {
        // Skip layout - we already handled it
        if (rel.Type?.includes('slideLayout')) continue;

        // Copy referenced files
        await this.copyRelatedFile(rel.Target, 'ppt/slides');
      }
    } else {
      // No source rels - create minimal rels with just layout
      const layoutRelPath = `../slideLayouts/${targetLayout.file}`;
      slideRelsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${layoutRelPath}"/>
</Relationships>`;
    }

    // Write slide rels
    this.outputZip.file(slideRelsPath, slideRelsContent);

    // Handle notes if present
    if (sourceSlide.notes && sourceSlide.notesPath) {
      const notesFileName = `notesSlide${slideNumber}.xml`;
      const notesPath = `ppt/notesSlides/${notesFileName}`;
      const notesRelsPath = `ppt/notesSlides/_rels/${notesFileName}.rels`;

      // Copy notes XML, updating the slide reference
      let notesXml = sourceSlide.notes;
      notesXml = notesXml.replace(
        /(<Relationship[^>]*Type="[^"]*\/slide"[^>]*Target=")[^"]*(")/g,
        `$1../slides/${slideFileName}$2`
      );

      this.outputZip.file(notesPath, notesXml);

      // Copy the original notes rels if it exists
      const sourceNotesRelsPath = sourceSlide.notesPath.replace('notesSlides/', 'notesSlides/_rels/') + '.rels';
      const sourceNotesRelsFile = this.sourceZip.file(sourceNotesRelsPath);

      if (sourceNotesRelsFile) {
        let notesRelsContent = await sourceNotesRelsFile.async('string');
        // Update slide reference
        notesRelsContent = notesRelsContent.replace(
          /(<Relationship[^>]*Type="[^"]*\/slide"[^>]*Target=")[^"]*(")/g,
          `$1../slides/${slideFileName}$2`
        );
        this.outputZip.file(notesRelsPath, notesRelsContent);
      } else {
        // Create minimal notes rels
        const notesRelsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${slideFileName}"/>
</Relationships>`;
        this.outputZip.file(notesRelsPath, notesRelsContent);
      }

      // Add notes relationship to slide rels if not already present
      if (!slideRelsContent.includes('notesSlide')) {
        const notesRelId = `rId${maxRelId + 1}`;
        slideRelsContent = slideRelsContent.replace(
          '</Relationships>',
          `<Relationship Id="${notesRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/${notesFileName}"/>
</Relationships>`
        );
        this.outputZip.file(slideRelsPath, slideRelsContent);
      }

      await this.ensureContentType(notesPath, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
    }

    // Register slide in Content_Types
    await this.ensureContentType(slidePath, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');

    // Add to presentation.xml.rels
    const presRelsPath = 'ppt/_rels/presentation.xml.rels';
    let presRelsStr = await this.outputZip.file(presRelsPath).async('string');

    const slideRelId = `rId${this.nextRelId++}`;

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
   * Clean up slide XML:
   * - Remove footer elements (date, slide number, footer text)
   * - Remove top-right corner images (likely logos)
   * - Expand text body areas to use more slide space
   */
  cleanupSlideXml(slideXml) {
    let xml = slideXml;

    // 1. Remove footer placeholder shapes (date, footer, slide number)
    // We need to carefully match individual <p:sp>...</p:sp> blocks and check each one
    xml = this.removeFooterShapes(xml);

    // 2. Remove top-right corner images (likely logos)
    xml = this.removeTopRightImages(xml);

    // 3. Expand body/content placeholders to use more slide width
    xml = this.expandTextAreas(xml);

    return xml;
  }

  /**
   * Remove footer placeholder shapes (date, footer, slide number)
   * Uses careful matching to avoid removing content shapes
   */
  removeFooterShapes(slideXml) {
    let xml = slideXml;

    // Footer placeholder types to remove
    const footerTypes = ['dt', 'ftr', 'sldNum'];

    // Footer-related names to remove (case insensitive check)
    const footerNames = ['footer', 'date', 'slide number', 'slidenumber', 'page number', 'pagenumber'];

    // Find all <p:sp> elements - use a function to extract them safely
    // We need to match balanced tags, so we'll process iteratively
    const shapes = this.extractShapeElements(xml);

    for (const shape of shapes) {
      let shouldRemove = false;

      // Check for footer placeholder types: <p:ph type="dt|ftr|sldNum"/>
      for (const footerType of footerTypes) {
        // Match type attribute with quotes - be strict about matching within this shape only
        const phRegex = new RegExp(`<p:ph[^>]*type=["']${footerType}["'][^>]*/>`);
        if (phRegex.test(shape)) {
          shouldRemove = true;
          break;
        }
      }

      // Check for footer-related names in cNvPr
      if (!shouldRemove) {
        const nameMatch = shape.match(/<p:cNvPr[^>]*name=["']([^"']*)["']/);
        if (nameMatch) {
          const shapeName = nameMatch[1].toLowerCase();
          for (const footerName of footerNames) {
            if (shapeName.includes(footerName)) {
              shouldRemove = true;
              break;
            }
          }
        }
      }

      if (shouldRemove) {
        xml = xml.replace(shape, '');
      }
    }

    return xml;
  }

  /**
   * Extract all <p:sp>...</p:sp> elements from XML string
   * Handles nested elements by tracking tag depth
   */
  extractShapeElements(xml) {
    const shapes = [];
    const startTag = '<p:sp';
    const endTag = '</p:sp>';

    let searchStart = 0;
    while (true) {
      const startIdx = xml.indexOf(startTag, searchStart);
      if (startIdx === -1) break;

      // Find the matching closing tag (accounting for nested p:sp is rare, but handle it)
      let depth = 1;
      let idx = startIdx + startTag.length;

      while (depth > 0 && idx < xml.length) {
        const nextStart = xml.indexOf(startTag, idx);
        const nextEnd = xml.indexOf(endTag, idx);

        if (nextEnd === -1) break; // Malformed XML

        if (nextStart !== -1 && nextStart < nextEnd) {
          // Found a nested start tag
          depth++;
          idx = nextStart + startTag.length;
        } else {
          // Found an end tag
          depth--;
          if (depth === 0) {
            const endIdx = nextEnd + endTag.length;
            shapes.push(xml.substring(startIdx, endIdx));
            searchStart = endIdx;
          } else {
            idx = nextEnd + endTag.length;
          }
        }
      }

      if (depth > 0) {
        // Couldn't find matching end tag, move past this start tag
        searchStart = startIdx + startTag.length;
      }
    }

    return shapes;
  }

  /**
   * Remove images positioned in the top-right corner (likely logos)
   */
  removeTopRightImages(slideXml) {
    // Optimized for Template_Transitionsphase_OESL.pptx
    // Slide dimensions: 18288000 x 10287000 EMUs (16:9 widescreen, 20" x 11.25")
    // Logo in template master is at x=15949416, y=268121
    // We consider top-right as: x > 14000000 (~75% from left) and y < 2000000 (~20% from top)

    const topRightThresholdX = 14000000; // ~75% from left for this wide slide
    const topRightThresholdY = 2000000;  // ~20% from top

    let xml = slideXml;

    // Find all p:pic elements
    const picRegex = /<p:pic[^>]*>[\s\S]*?<\/p:pic>/g;
    const pics = xml.match(picRegex) || [];

    for (const pic of pics) {
      // Extract position from <a:off x="..." y="..."/>
      const offMatch = pic.match(/<a:off[^>]*x="(\d+)"[^>]*y="(\d+)"/);
      if (offMatch) {
        const x = parseInt(offMatch[1]);
        const y = parseInt(offMatch[2]);

        // Check if in top-right corner (likely a logo)
        if (x > topRightThresholdX && y < topRightThresholdY) {
          // Remove this picture
          xml = xml.replace(pic, '');
        }
      }
    }

    return xml;
  }

  /**
   * Expand text/body placeholder areas to use more of the slide
   * Optimized for Template_Transitionsphase_OESL.pptx (18288000 x 10287000 EMUs)
   */
  expandTextAreas(slideXml) {
    let xml = slideXml;

    // Template dimensions: 18288000 x 10287000 EMUs
    // Template body placeholder: x=1257300, y=2738438, width=15773400, height=6527800
    // We want to expand content to use similar positioning but full width

    const slideWidth = 18288000;         // Template slide width
    const newLeftMargin = 914400;        // ~1 inch from left (consistent with template)
    const newRightMargin = 914400;       // ~1 inch from right
    const maxContentWidth = slideWidth - newLeftMargin - newRightMargin; // ~16459200

    // Match <a:xfrm> blocks within <p:sp> elements that contain body placeholders
    const spRegex = /<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g;
    const shapes = xml.match(spRegex) || [];

    for (const shape of shapes) {
      // Check if this is a body/content placeholder (not title, not footer types)
      const isTitle = /<p:ph[^>]*type="(title|ctrTitle)"/.test(shape);
      const isFooter = /<p:ph[^>]*type="(dt|ftr|sldNum)"/.test(shape);
      const isSubtitle = /<p:ph[^>]*type="subTitle"/.test(shape);

      if (isTitle || isFooter || isSubtitle) continue;

      // Check if it has a text body (indicates content shape)
      if (!/<p:txBody/.test(shape)) continue;

      // Find and modify the xfrm
      const xfrmMatch = shape.match(/<a:xfrm[^>]*>([\s\S]*?)<\/a:xfrm>/);
      if (!xfrmMatch) continue;

      const xfrmContent = xfrmMatch[0];
      const offMatch = xfrmContent.match(/<a:off[^>]*x="(\d+)"[^>]*y="(\d+)"[^>]*\/>/);
      const extMatch = xfrmContent.match(/<a:ext[^>]*cx="(\d+)"[^>]*cy="(\d+)"[^>]*\/>/);

      if (!offMatch || !extMatch) continue;

      const currentX = parseInt(offMatch[1]);
      const currentY = parseInt(offMatch[2]);
      const currentWidth = parseInt(extMatch[1]);
      const currentHeight = parseInt(extMatch[2]);

      // Only expand if width is less than 85% of max content width
      if (currentWidth > maxContentWidth * 0.85) continue;

      // Calculate new dimensions - expand to fill more space
      const newX = newLeftMargin;
      const newWidth = maxContentWidth;

      // Create new xfrm content
      const newXfrm = xfrmContent
        .replace(/<a:off[^>]*\/>/, `<a:off x="${newX}" y="${currentY}"/>`)
        .replace(/<a:ext[^>]*\/>/, `<a:ext cx="${newWidth}" cy="${currentHeight}"/>`);

      // Replace in original shape
      const newShape = shape.replace(xfrmContent, newXfrm);
      xml = xml.replace(shape, newShape);
    }

    return xml;
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

    // Remove old slide overrides from template (they were deleted)
    contentTypesStr = contentTypesStr.replace(
      /<Override[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>\s*/g,
      ''
    );

    // Remove old notesSlide overrides from template
    contentTypesStr = contentTypesStr.replace(
      /<Override[^>]*PartName="\/ppt\/notesSlides\/notesSlide\d+\.xml"[^>]*\/>\s*/g,
      ''
    );

    // Clean up extra whitespace
    contentTypesStr = contentTypesStr.replace(/\n\s*\n/g, '\n');

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
   * Copy a related file from source to output
   * Handles relative paths from relationship targets
   */
  async copyRelatedFile(targetPath, baseDir) {
    if (!targetPath) return;

    // Skip external URLs
    if (targetPath.startsWith('http://') || targetPath.startsWith('https://')) {
      return;
    }

    // Resolve the full path
    let sourcePath;
    if (targetPath.startsWith('../')) {
      // Relative path like ../media/image1.png
      sourcePath = `ppt/${targetPath.replace('../', '')}`;
    } else if (targetPath.startsWith('/')) {
      // Absolute path
      sourcePath = targetPath.substring(1);
    } else {
      // Relative to base dir
      sourcePath = `${baseDir}/${targetPath}`;
    }

    // Normalize the path
    sourcePath = sourcePath.replace(/\/+/g, '/');

    // Check if file exists in source
    const sourceFile = this.sourceZip.file(sourcePath);
    if (sourceFile) {
      try {
        const content = await sourceFile.async('nodebuffer');
        this.outputZip.file(sourcePath, content);

        // Ensure content type is registered for media files
        await this.ensureContentType(sourcePath);
      } catch (e) {
        console.error(`Failed to copy file ${sourcePath}:`, e.message);
      }
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
