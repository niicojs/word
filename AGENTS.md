# AGENTS.md

Instructions for AI agents working on this codebase.

## Project Overview

`@niicojs/word` is a TypeScript library for Word/DOCX document operations. Zero-config setup using modern, high-performance tooling.

**Current Features (v0.1):**
- **Document merging** - Merge pages from multiple DOCX files
- **Page management** - Add/remove pages using page breaks

This library follows the same patterns as `@niicojs/excel` (sibling project in `../excel`).

## Tech Stack

- **Runtime**: Node.js >= 20
- **Package Manager**: Bun
- **Language**: TypeScript 5.7+ (strict mode)
- **Bundler**: tsdx (wraps Bunchee)
- **Testing**: Vitest
- **Linting**: Oxlint
- **Formatting**: Oxfmt

## Project Structure

```
src/
  index.ts          # Library entry point - ALL public APIs exported here
  types.ts          # Type definitions (MergeOptions, PageInfo)
  document.ts       # Main Document class
  utils/
    xml.ts          # XML parsing/generation
    zip.ts          # DOCX ZIP handling
test/
  document.test.ts  # Test files using Vitest
dist/               # Build output (generated, do not edit)
```

## Commands

### Build & Development

Run all test with bun, not node. 

```bash
bun install          # Install dependencies
bun run dev          # Start development mode with watch
bun run build        # Build for production (outputs ESM, CJS, and .d.ts)
```

### Testing

```bash
bun run test                              # Run all tests
bun run test:watch                        # Run tests in watch mode
bun run test -- test/document.test.ts     # Run a specific test file
bun run test -- -t "test name"            # Run tests matching a pattern
```

### Code Quality

```bash
bun run lint         # Lint code with Oxlint
bun run format       # Format code with Oxfmt
bun run format:check # Check formatting without modifying
bun run typecheck    # Run TypeScript type checking
```

### Before Committing

```bash
bun run typecheck && bun run lint && bun run test
```

## Public API

### Document Class

```typescript
import { Document } from '@niicojs/word';

// Factory methods
Document.create(): Document                           // Create empty document
Document.fromFile(path: string): Promise<Document>    // Load from file
Document.fromBuffer(data: Uint8Array): Promise<Document>  // Load from buffer

// Serialization
doc.toFile(path: string): Promise<void>               // Save to file
doc.toBuffer(): Promise<Uint8Array>                   // Save to buffer

// Page management
doc.getPageCount(): number                            // Get page count
doc.getPages(): PageInfo[]                            // Get page info array
doc.addPageBreak(): void                              // Add page break at end
doc.removePage(index: number): void                   // Remove page (0-based)

// Merging
doc.merge(source, options?): Promise<void>            // Merge from another doc
```

### Types

```typescript
interface MergeOptions {
  pages?: number[];           // Page indices to merge (0-based), omit for all
  addPageBreakBefore?: boolean;  // Add break before merged content (default: true)
}

interface PageInfo {
  index: number;              // 0-based page index
  startElement: number;       // First body element index (-1 if empty)
  endElement: number;         // Last body element index (-1 if empty)
}
```

## Code Style Guidelines

### Formatting

- **Line width**: 120 characters max
- **Quotes**: Single quotes for strings
- **Semicolons**: Required
- **Indentation**: 2 spaces

### Imports

```typescript
// Good - ordered: Node.js, external, local; types separate
import { readFile } from 'fs/promises';
import { something } from 'external-package';
import type { MergeOptions } from './types';
import { localThing } from './local';
```

### Exports

- Use named exports, avoid default exports
- Export all public APIs from `src/index.ts`
- Export types with `export type`

### Naming Conventions

- **Functions/variables**: camelCase
- **Types/Interfaces/Classes**: PascalCase
- **Private members**: Prefix with underscore (`_bodyElements`)
- **Files**: kebab-case (`shared-strings.ts`)
- **Test files**: `*.test.ts`

### Error Handling

```typescript
/**
 * Remove a page by index.
 * @throws {Error} If index is out of bounds
 */
removePage(index: number): void {
  if (index < 0 || index >= pages.length) {
    throw new Error(`Page index out of bounds: ${index}`);
  }
  // ...
}
```

### TypeScript Rules (enforced)

- `noUnusedLocals`, `noUnusedParameters`
- `noImplicitReturns`, `noFallthroughCasesInSwitch`
- `strict` mode enabled

## DOCX File Format

DOCX files are ZIP archives containing XML:

- `word/document.xml` - Main content (`<w:body>` contains paragraphs/tables)
- `word/styles.xml` - Style definitions
- `word/_rels/document.xml.rels` - Relationships
- `[Content_Types].xml` - Content type declarations

**Page breaks** are detected as: `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`

Key namespaces:
- `w:` - WordprocessingML (main content)
- `r:` - Relationships

## Testing Guidelines

```typescript
import { describe, it, expect } from 'vitest';
import { Document } from '../src';

describe('Document', () => {
  it('creates an empty document', () => {
    const doc = Document.create();
    expect(doc.getPageCount()).toBe(1);
  });

  it('adds page breaks', () => {
    const doc = Document.create();
    doc.addPageBreak();
    expect(doc.getPageCount()).toBe(2);
  });

  it('throws on invalid index', () => {
    const doc = Document.create();
    expect(() => doc.removePage(-1)).toThrow('out of bounds');
  });
});
```

## Adding New Features

1. Implement in `src/`
2. Export public APIs from `src/index.ts`
3. Add tests in `test/`
4. Add JSDoc comments
5. Run `bun run typecheck && bun run lint && bun run test`
6. Run `bun run build`

## Important Notes

- Page indices are **0-based**
- Empty documents have 1 page
- Page breaks mark page boundaries (break element ends one page, next content starts new page)
- Reference `../excel` for similar patterns
- Dependencies: `fflate` (ZIP), `fast-xml-parser` (XML)
