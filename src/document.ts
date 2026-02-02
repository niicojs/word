import { readFile, writeFile } from 'fs/promises';
import type { MergeOptions, PageInfo } from './types';
import { readZip, writeZip, readZipText, writeZipText, ZipFiles } from './utils/zip';
import {
  parseXml,
  stringifyXml,
  findElement,
  getChildren,
  createElement,
  cloneNodes,
  cloneNode,
  XmlNode,
} from './utils/xml';

// WordprocessingML namespace
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

/**
 * Represents a Word document (.docx file)
 */
export class Document {
  private _files: ZipFiles = new Map();
  private _bodyElements: XmlNode[] = [];
  private _documentParsed: XmlNode[] = [];
  private _dirty = false;
  /** Namespace declarations from source documents (e.g., {'xmlns:w': '...', 'xmlns:w14': '...'}) */
  private _namespaces: Record<string, string> = {};
  /** Section properties from source document */
  private _sectPr: XmlNode | null = null;

  private constructor() {}

  /**
   * Create a new empty document
   */
  static create(): Document {
    const doc = new Document();
    doc._dirty = true;
    return doc;
  }

  /**
   * Load a document from a file path
   * @param path - Path to the .docx file
   */
  static async fromFile(path: string): Promise<Document> {
    const data = await readFile(path);
    return Document.fromBuffer(new Uint8Array(data));
  }

  /**
   * Load a document from a buffer
   * @param data - DOCX file as Uint8Array
   */
  static async fromBuffer(data: Uint8Array): Promise<Document> {
    const doc = new Document();
    doc._files = await readZip(data);

    // Parse document.xml
    const documentXml = readZipText(doc._files, 'word/document.xml');
    if (documentXml) {
      doc._parseDocument(documentXml);
    }

    return doc;
  }

  /**
   * Save the document to a file
   * @param path - Path to save the .docx file
   */
  async toFile(path: string): Promise<void> {
    const buffer = await this.toBuffer();
    await writeFile(path, buffer);
  }

  /**
   * Save the document to a buffer
   * @returns DOCX file as Uint8Array
   */
  async toBuffer(): Promise<Uint8Array> {
    this._updateFiles();
    return writeZip(this._files);
  }

  /**
   * Get the number of pages in the document.
   * Pages are determined by page break elements.
   */
  getPageCount(): number {
    return this.getPages().length;
  }

  /**
   * Get information about all pages in the document.
   * Pages are determined by page break elements.
   * A page break marks the end of one page and the start of another.
   */
  getPages(): PageInfo[] {
    const pages: PageInfo[] = [];
    let currentPageStart = 0;
    let pageIndex = 0;

    for (let i = 0; i < this._bodyElements.length; i++) {
      const element = this._bodyElements[i];

      // Check if this element contains a page break
      if (this._hasPageBreak(element)) {
        // End current page at this element (the page break is at the end of this page)
        pages.push({
          index: pageIndex,
          startElement: currentPageStart,
          endElement: i,
        });
        pageIndex++;
        currentPageStart = i + 1;
      }
    }

    // Add the last page (from currentPageStart to end, or empty final page after last break)
    // This handles:
    // 1. Empty document (no elements) -> 1 page
    // 2. Document with content but no breaks -> 1 page
    // 3. Document ending with a page break -> add empty final page
    // 4. Document with content after last break -> add that content as final page
    if (currentPageStart < this._bodyElements.length) {
      // There's content after the last page break
      pages.push({
        index: pageIndex,
        startElement: currentPageStart,
        endElement: this._bodyElements.length - 1,
      });
    } else if (pages.length > 0) {
      // Document ends with a page break, add empty final page
      // Use -1 for both start/end to indicate empty page
      pages.push({
        index: pageIndex,
        startElement: -1,
        endElement: -1,
      });
    } else {
      // Empty document with no elements and no breaks -> 1 empty page
      pages.push({
        index: 0,
        startElement: -1,
        endElement: -1,
      });
    }

    return pages;
  }

  /**
   * Add a page break at the end of the document
   */
  addPageBreak(): void {
    this._dirty = true;
    const pageBreakParagraph = this._createPageBreakParagraph();
    this._bodyElements.push(pageBreakParagraph);
  }

  /**
   * Remove a page and its content.
   * @param index - 0-based page index to remove
   * @throws {Error} If index is out of bounds
   */
  removePage(index: number): void {
    const pages = this.getPages();

    if (index < 0 || index >= pages.length) {
      throw new Error(`Page index out of bounds: ${index}. Document has ${pages.length} page(s).`);
    }

    this._dirty = true;
    const page = pages[index];

    // Handle empty page (startElement is -1)
    if (page.startElement === -1) {
      // This is an empty trailing page, nothing to remove in body
      // But if it's after a page break, we should remove the preceding page break
      if (index > 0 && pages[index - 1].endElement >= 0) {
        // Remove the page break from the previous page
        const prevPage = pages[index - 1];
        // The page break is the last element of the previous page
        this._bodyElements.splice(prevPage.endElement, 1);
      }
      return;
    }

    // Calculate how many elements to remove
    const removeCount = page.endElement - page.startElement + 1;

    // Remove elements for this page
    this._bodyElements.splice(page.startElement, removeCount);
  }

  /**
   * Merge pages from another document into this one.
   * @param source - Document, file path, or buffer to merge from
   * @param options - Merge options
   */
  async merge(source: Document | string | Uint8Array, options: MergeOptions = {}): Promise<void> {
    // Load source document if needed
    let sourceDoc: Document;
    if (source instanceof Document) {
      sourceDoc = source;
    } else if (typeof source === 'string') {
      sourceDoc = await Document.fromFile(source);
    } else {
      sourceDoc = await Document.fromBuffer(source);
    }

    this._dirty = true;

    // Merge namespaces from source document
    this._mergeNamespaces(sourceDoc._namespaces);

    // If this is the first merge (empty document), copy document structure from source
    const isFirstMerge = this._bodyElements.length === 0 && this._files.size === 0;
    if (isFirstMerge) {
      this._copyDocumentStructure(sourceDoc);
    }

    // Get source pages
    const sourcePages = sourceDoc.getPages();
    const addPageBreakBefore = options.addPageBreakBefore !== false;

    // Determine which pages to merge
    let pagesToMerge: PageInfo[];
    if (options.pages !== undefined) {
      // Filter to only requested pages (ignore out-of-range indices)
      pagesToMerge = options.pages
        .filter((idx) => idx >= 0 && idx < sourcePages.length)
        .map((idx) => sourcePages[idx]);
    } else {
      // Merge all pages
      pagesToMerge = sourcePages;
    }

    if (pagesToMerge.length === 0) {
      return; // Nothing to merge
    }

    // Add page break before merged content if requested and document is not empty
    if (addPageBreakBefore && this._bodyElements.length > 0) {
      this._bodyElements.push(this._createPageBreakParagraph());
    }

    // Collect elements from selected pages
    for (const page of pagesToMerge) {
      // Skip empty pages (startElement is -1)
      if (page.startElement === -1) continue;
      
      const elements = sourceDoc._bodyElements.slice(page.startElement, page.endElement + 1);
      // Deep clone elements to avoid reference issues
      const clonedElements = cloneNodes(elements);
      this._bodyElements.push(...clonedElements);
    }
  }

  /**
   * Merge namespace declarations from another document
   */
  private _mergeNamespaces(sourceNamespaces: Record<string, string>): void {
    for (const [key, value] of Object.entries(sourceNamespaces)) {
      // Only add if not already present (first wins)
      if (!(key in this._namespaces)) {
        this._namespaces[key] = value;
      }
    }
  }

  /**
   * Copy document structure (styles, fonts, etc.) from source document
   * Called only on first merge to preserve document formatting
   */
  private _copyDocumentStructure(sourceDoc: Document): void {
    // Copy section properties
    if (sourceDoc._sectPr && !this._sectPr) {
      this._sectPr = cloneNode(sourceDoc._sectPr);
    }

    // Copy all files from source EXCEPT word/document.xml (we generate that ourselves)
    for (const [path, content] of sourceDoc._files) {
      if (path === 'word/document.xml') {
        continue; // We generate this ourselves
      }
      if (!this._files.has(path)) {
        this._files.set(path, content);
      }
    }
  }

  /**
   * Parse the document.xml content
   */
  private _parseDocument(xml: string): void {
    this._documentParsed = parseXml(xml);

    // Find the document element
    const documentNode = findElement(this._documentParsed, 'w:document');
    if (!documentNode) return;

    // Extract namespace declarations from the document element
    const attrs = documentNode[':@'] as Record<string, string> | undefined;
    if (attrs) {
      for (const [key, value] of Object.entries(attrs)) {
        // Namespace attributes start with @_xmlns
        if (key.startsWith('@_xmlns')) {
          // Store without the @_ prefix (e.g., 'xmlns:w', 'xmlns:w14')
          const nsKey = key.slice(2); // Remove '@_'
          this._namespaces[nsKey] = value;
        }
      }
    }

    // Find the body element
    const documentChildren = getChildren(documentNode, 'w:document');
    const bodyNode = findElement(documentChildren, 'w:body');
    if (!bodyNode) return;

    // Get body children (paragraphs, tables, etc.)
    const bodyChildren = getChildren(bodyNode, 'w:body');
    
    // Extract sectPr (section properties) and keep it separately
    const sectPrNode = findElement(bodyChildren, 'w:sectPr');
    if (sectPrNode) {
      this._sectPr = sectPrNode;
    }
    
    // Filter out sectPr from body elements
    this._bodyElements = bodyChildren.filter((child) => !('w:sectPr' in child));
  }

  /**
   * Check if an element contains a page break
   */
  private _hasPageBreak(element: XmlNode): boolean {
    // Check if this is a paragraph with a page break
    if ('w:p' in element) {
      return this._paragraphHasPageBreak(element);
    }
    return false;
  }

  /**
   * Check if a paragraph contains a page break
   */
  private _paragraphHasPageBreak(paragraph: XmlNode): boolean {
    const children = getChildren(paragraph, 'w:p');

    for (const child of children) {
      // Check for run elements
      if ('w:r' in child) {
        const runChildren = getChildren(child, 'w:r');
        for (const runChild of runChildren) {
          // Check for break element with type="page"
          if ('w:br' in runChild) {
            const attrs = runChild[':@'] as Record<string, string> | undefined;
            if (attrs?.['@_w:type'] === 'page') {
              return true;
            }
          }
        }
      }
    }

    return false;
  }

  /**
   * Create a paragraph element with a page break
   */
  private _createPageBreakParagraph(): XmlNode {
    // Create: <w:p><w:r><w:br w:type="page"/></w:r></w:p>
    const brElement = createElement('w:br', { 'w:type': 'page' }, []);
    const runElement = createElement('w:r', {}, [brElement]);
    const paragraphElement = createElement('w:p', {}, [runElement]);
    return paragraphElement;
  }

  /**
   * Update the ZIP files with current document state
   */
  private _updateFiles(): void {
    if (!this._dirty && this._files.size > 0) {
      return;
    }

    // Ensure basic structure exists
    this._ensureBasicStructure();

    // Update document.xml
    this._updateDocumentXml();
  }

  /**
   * Ensure basic DOCX structure exists
   */
  private _ensureBasicStructure(): void {
    // [Content_Types].xml
    if (!this._files.has('[Content_Types].xml')) {
      const contentTypes = this._createContentTypes();
      writeZipText(this._files, '[Content_Types].xml', contentTypes);
    }

    // _rels/.rels
    if (!this._files.has('_rels/.rels')) {
      const rels = this._createRootRels();
      writeZipText(this._files, '_rels/.rels', rels);
    }

    // word/_rels/document.xml.rels
    if (!this._files.has('word/_rels/document.xml.rels')) {
      const docRels = this._createDocumentRels();
      writeZipText(this._files, 'word/_rels/document.xml.rels', docRels);
    }
  }

  /**
   * Create [Content_Types].xml
   */
  private _createContentTypes(): string {
    const types = createElement(
      'Types',
      { xmlns: CT_NS },
      [
        createElement('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }, []),
        createElement('Default', { Extension: 'xml', ContentType: 'application/xml' }, []),
        createElement('Override', {
          PartName: '/word/document.xml',
          ContentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        }, []),
      ]
    );
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([types])}`;
  }

  /**
   * Create _rels/.rels
   */
  private _createRootRels(): string {
    const rels = createElement(
      'Relationships',
      { xmlns: PKG_NS },
      [
        createElement('Relationship', {
          Id: 'rId1',
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
          Target: 'word/document.xml',
        }, []),
      ]
    );
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([rels])}`;
  }

  /**
   * Create word/_rels/document.xml.rels
   */
  private _createDocumentRels(): string {
    const rels = createElement('Relationships', { xmlns: PKG_NS }, []);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([rels])}`;
  }

  /**
   * Update document.xml with current body elements
   */
  private _updateDocumentXml(): void {
    // Build body content
    const bodyChildren: XmlNode[] = [...this._bodyElements];

    // Add section properties at the end (required for valid DOCX)
    if (this._sectPr) {
      // Use preserved section properties from source
      bodyChildren.push(cloneNode(this._sectPr));
    } else {
      // Create default section properties
      const sectPr = createElement('w:sectPr', {}, [
        createElement('w:pgSz', { 'w:w': '12240', 'w:h': '15840' }, []), // Letter size
        createElement('w:pgMar', {
          'w:top': '1440',
          'w:right': '1440',
          'w:bottom': '1440',
          'w:left': '1440',
          'w:header': '720',
          'w:footer': '720',
          'w:gutter': '0',
        }, []),
      ]);
      bodyChildren.push(sectPr);
    }

    const body = createElement('w:body', {}, bodyChildren);
    
    // Build namespace attributes - use collected namespaces or defaults
    const namespaceAttrs = this._getNamespaceAttributes();
    const document = createElement('w:document', namespaceAttrs, [body]);

    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${stringifyXml([document])}`;
    writeZipText(this._files, 'word/document.xml', xml);
  }

  /**
   * Get namespace attributes for the document element
   */
  private _getNamespaceAttributes(): Record<string, string> {
    // If we have collected namespaces, use them
    if (Object.keys(this._namespaces).length > 0) {
      return { ...this._namespaces };
    }
    
    // Default minimal namespaces for new documents
    return {
      'xmlns:w': W_NS,
      'xmlns:r': R_NS,
    };
  }
}
