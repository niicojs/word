# @niicojs/word

A lightweight TypeScript library for Word document (.docx) manipulation. Merge documents, manage pages, and more.

## Features

- **Document Merging** - Combine multiple DOCX files into one
- **Page Management** - Add page breaks, remove pages, get page information
- **Selective Merging** - Choose specific pages to merge from source documents
- **Templating** - Replace placeholders like `{name}` in body, tables, headers, and footers
- **Full Structure Preservation** - Maintains styles, fonts, headers, footers, and other document properties
- **Zero Native Dependencies** - Pure JavaScript/TypeScript implementation

## Installation

```bash
npm install @niicojs/word
# or
bun add @niicojs/word
# or
pnpm add @niicojs/word
```

## Quick Start

### Merge Multiple Documents

```typescript
import { Document } from '@niicojs/word';

// Create an empty output document
const output = Document.create();

// Merge documents (page breaks are automatically added between documents)
await output.merge('document1.docx');
await output.merge('document2.docx');
await output.merge('document3.docx');

// Save the merged document
await output.toFile('merged.docx');
```

### Load and Inspect a Document

```typescript
import { Document } from '@niicojs/word';

const doc = await Document.fromFile('document.docx');

console.log(`Page count: ${doc.getPageCount()}`);
console.log(`Pages:`, doc.getPages());
```

### Merge Specific Pages

```typescript
import { Document } from '@niicojs/word';

const output = Document.create();

// Merge only pages 0 and 2 (first and third pages)
await output.merge('source.docx', { pages: [0, 2] });

// Merge without adding a page break before
await output.merge('another.docx', { addPageBreakBefore: false });

await output.toFile('output.docx');
```

### Work with Buffers

```typescript
import { Document } from '@niicojs/word';

// Load from buffer
const data = await fetch('https://example.com/document.docx')
  .then(res => res.arrayBuffer())
  .then(buf => new Uint8Array(buf));

const doc = await Document.fromBuffer(data);

// Save to buffer
const outputBuffer = await doc.toBuffer();
```

### Page Management

```typescript
import { Document } from '@niicojs/word';

const doc = await Document.fromFile('document.docx');

// Add a page break at the end
doc.addPageBreak();

// Remove a specific page (0-based index)
doc.removePage(1); // Removes the second page

await doc.toFile('modified.docx');
```

### Templating

```typescript
import { Document } from '@niicojs/word';

const doc = await Document.fromFile('template.docx');

doc.render({
  name: 'Ava',
  city: 'Paris',
});

await doc.toFile('filled.docx');
```

## API Reference

### Document Class

#### Static Methods

| Method | Description |
|--------|-------------|
| `Document.create()` | Create a new empty document |
| `Document.fromFile(path)` | Load a document from a file path |
| `Document.fromBuffer(data)` | Load a document from a Uint8Array |

#### Instance Methods

| Method | Description |
|--------|-------------|
| `toFile(path)` | Save the document to a file |
| `toBuffer()` | Save the document to a Uint8Array |
| `getPageCount()` | Get the number of pages |
| `getPages()` | Get information about all pages |
| `addPageBreak()` | Add a page break at the end |
| `removePage(index)` | Remove a page by index (0-based) |
| `merge(source, options?)` | Merge content from another document |
| `render(data, options?)` | Replace template placeholders |

### MergeOptions

```typescript
interface MergeOptions {
  /** Page indices to include (0-based). If omitted, all pages are merged. */
  pages?: number[];
  
  /** Insert page break before merged content. Default: true */
  addPageBreakBefore?: boolean;
}
```

### TemplateOptions

```typescript
type TemplateValue = string | number | boolean | null | undefined;

type TemplateData = Record<string, TemplateValue>;

interface TemplateOptions {
  /** Placeholder pattern. Must contain one capture group for the key. */
  pattern?: RegExp;

  /** Remove placeholders when the key is missing. Default: false */
  removeMissing?: boolean;

  /** Optional value transformer applied before insertion. */
  transform?: (value: TemplateValue, key: string) => string;
}
```

### PageInfo

```typescript
interface PageInfo {
  /** 0-based page index */
  index: number;
  
  /** Index of first body element in this page */
  startElement: number;
  
  /** Index of last body element in this page (inclusive) */
  endElement: number;
}
```

## How It Works

This library works by:

1. Reading DOCX files as ZIP archives
2. Parsing the XML content (specifically `word/document.xml`)
3. Manipulating the document structure in memory
4. Writing back to a valid DOCX format

**Page detection** is based on explicit page break elements (`<w:br w:type="page"/>`). Note that Word may render page breaks in different positions based on content flow, fonts, and page size - this library only detects explicit page breaks.

When merging documents, the library:
- Preserves all namespace declarations from source documents
- Copies document structure (styles, fonts, headers, footers, etc.) from the first source
- Combines body content with configurable page breaks between documents

## Requirements

- Node.js >= 20
- Works with Bun and other modern JavaScript runtimes

## Dependencies

- [fflate](https://github.com/101arrowz/fflate) - Fast, lightweight ZIP compression
- [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser) - XML parsing and serialization

## License

MIT
