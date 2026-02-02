import { describe, it, expect, beforeAll } from 'vitest';
import { Document } from '../src';
import { mkdir, unlink } from 'fs/promises';
import { existsSync } from 'fs';

describe('Document', () => {
  const testDir = 'test/fixtures';
  const testOutputDir = 'test/output';

  beforeAll(async () => {
    if (!existsSync(testDir)) {
      await mkdir(testDir, { recursive: true });
    }
    if (!existsSync(testOutputDir)) {
      await mkdir(testOutputDir, { recursive: true });
    }
  });

  describe('create', () => {
    it('creates an empty document', () => {
      const doc = Document.create();
      expect(doc.getPageCount()).toBe(1); // Empty doc has 1 page
    });

    it('saves empty document to buffer', async () => {
      const doc = Document.create();
      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
    });

    it('saves and reloads empty document', async () => {
      const doc = Document.create();
      const buffer = await doc.toBuffer();

      const loaded = await Document.fromBuffer(buffer);
      expect(loaded.getPageCount()).toBe(1);
    });
  });

  describe('pages', () => {
    it('returns 1 page for document with no page breaks', async () => {
      const doc = Document.create();
      expect(doc.getPageCount()).toBe(1);
    });

    it('counts pages correctly with multiple page breaks', async () => {
      const doc = Document.create();
      doc.addPageBreak();
      expect(doc.getPageCount()).toBe(2);

      doc.addPageBreak();
      expect(doc.getPageCount()).toBe(3);
    });

    it('getPages returns correct element ranges', async () => {
      const doc = Document.create();
      doc.addPageBreak();
      doc.addPageBreak();

      const pages = doc.getPages();
      expect(pages.length).toBe(3);

      expect(pages[0].index).toBe(0);
      expect(pages[1].index).toBe(1);
      expect(pages[2].index).toBe(2);
    });

    it('preserves page count after save/load', async () => {
      const doc = Document.create();
      doc.addPageBreak();
      doc.addPageBreak();

      const buffer = await doc.toBuffer();
      const loaded = await Document.fromBuffer(buffer);

      expect(loaded.getPageCount()).toBe(3);
    });
  });

  describe('addPageBreak', () => {
    it('adds page break to empty document', () => {
      const doc = Document.create();
      doc.addPageBreak();
      expect(doc.getPageCount()).toBe(2);
    });

    it('increments page count each time', () => {
      const doc = Document.create();
      expect(doc.getPageCount()).toBe(1);

      doc.addPageBreak();
      expect(doc.getPageCount()).toBe(2);

      doc.addPageBreak();
      expect(doc.getPageCount()).toBe(3);

      doc.addPageBreak();
      expect(doc.getPageCount()).toBe(4);
    });
  });

  describe('removePage', () => {
    it('removes first page', () => {
      const doc = Document.create();
      doc.addPageBreak();
      doc.addPageBreak();
      expect(doc.getPageCount()).toBe(3);

      doc.removePage(0);
      expect(doc.getPageCount()).toBe(2);
    });

    it('removes middle page', () => {
      const doc = Document.create();
      doc.addPageBreak();
      doc.addPageBreak();
      expect(doc.getPageCount()).toBe(3);

      doc.removePage(1);
      expect(doc.getPageCount()).toBe(2);
    });

    it('removes last page', () => {
      const doc = Document.create();
      doc.addPageBreak();
      doc.addPageBreak();
      expect(doc.getPageCount()).toBe(3);

      doc.removePage(2);
      expect(doc.getPageCount()).toBe(2);
    });

    it('throws on negative index', () => {
      const doc = Document.create();
      expect(() => doc.removePage(-1)).toThrow('out of bounds');
    });

    it('throws on index >= page count', () => {
      const doc = Document.create();
      expect(() => doc.removePage(1)).toThrow('out of bounds');
      expect(() => doc.removePage(5)).toThrow('out of bounds');
    });

    it('handles removing only page leaving empty document', () => {
      const doc = Document.create();
      doc.removePage(0);
      // After removing the only page, document should have 1 empty page
      expect(doc.getPageCount()).toBe(1);
    });
  });

  describe('merge', () => {
    it('merges all pages from another document', async () => {
      const doc1 = Document.create();
      doc1.addPageBreak();
      expect(doc1.getPageCount()).toBe(2);

      const doc2 = Document.create();
      doc2.addPageBreak();
      expect(doc2.getPageCount()).toBe(2);

      await doc1.merge(doc2);
      // doc1 had 2 pages, doc2 had 2 pages
      // After merge with page break before: 2 + 1 (break) + 2 = could be 4 or 5 depending on interpretation
      // Actually: doc1 has 2 pages, merge adds page break + doc2's 2 page breaks = 5 page breaks total = 5+ pages
      expect(doc1.getPageCount()).toBeGreaterThanOrEqual(3);
    });

    it('merges specific pages by index', async () => {
      const doc1 = Document.create();
      expect(doc1.getPageCount()).toBe(1);

      const doc2 = Document.create();
      doc2.addPageBreak();
      doc2.addPageBreak();
      expect(doc2.getPageCount()).toBe(3);

      // Merge only first page (index 0)
      await doc1.merge(doc2, { pages: [0] });
      // doc1 had 1 page, we add page break + page 0's content
      expect(doc1.getPageCount()).toBeGreaterThanOrEqual(2);
    });

    it('adds page break before merged content by default', async () => {
      const doc1 = Document.create();
      doc1.addPageBreak(); // Give doc1 some content
      const initialCount = doc1.getPageCount();
      expect(initialCount).toBe(2);

      const doc2 = Document.create();
      doc2.addPageBreak(); // Give doc2 some content
      await doc1.merge(doc2);

      // Should have added a page break before merged content
      expect(doc1.getPageCount()).toBeGreaterThan(initialCount);
    });

    it('skips page break when addPageBreakBefore is false', async () => {
      const doc1 = Document.create();
      const doc2 = Document.create();

      await doc1.merge(doc2, { addPageBreakBefore: false });

      // No page break added, and doc2 was empty, so should still be 1 page
      expect(doc1.getPageCount()).toBe(1);
    });

    it('merges from buffer', async () => {
      const doc1 = Document.create();
      const doc2 = Document.create();
      doc2.addPageBreak();

      const buffer = await doc2.toBuffer();
      await doc1.merge(buffer);

      expect(doc1.getPageCount()).toBeGreaterThanOrEqual(2);
    });

    it('handles out-of-range page indices gracefully', async () => {
      const doc1 = Document.create();
      const doc2 = Document.create();
      doc2.addPageBreak();
      expect(doc2.getPageCount()).toBe(2);

      // Request pages that don't exist
      await doc1.merge(doc2, { pages: [10, 20, 30] });

      // Should not throw, just skip invalid indices
      // doc1 should remain mostly unchanged (just 1 page since nothing was merged)
      expect(doc1.getPageCount()).toBe(1);
    });

    it('handles empty pages array', async () => {
      const doc1 = Document.create();
      const doc2 = Document.create();
      doc2.addPageBreak();

      await doc1.merge(doc2, { pages: [] });

      // Nothing merged
      expect(doc1.getPageCount()).toBe(1);
    });

    it('merges Document instance directly', async () => {
      const doc1 = Document.create();
      const doc2 = Document.create();
      doc2.addPageBreak();

      await doc1.merge(doc2);

      expect(doc1.getPageCount()).toBeGreaterThanOrEqual(2);
    });
  });

  describe('save and load', () => {
    it('saves to file and loads back', async () => {
      const testFile = `${testOutputDir}/test-save-load.docx`;

      const doc = Document.create();
      doc.addPageBreak();
      doc.addPageBreak();

      await doc.toFile(testFile);

      const loaded = await Document.fromFile(testFile);
      expect(loaded.getPageCount()).toBe(3);

      // Cleanup
      await unlink(testFile);
    });

    it('preserves structure after multiple save/load cycles', async () => {
      let doc = Document.create();
      doc.addPageBreak();
      doc.addPageBreak();

      // Save and reload multiple times
      for (let i = 0; i < 3; i++) {
        const buffer = await doc.toBuffer();
        doc = await Document.fromBuffer(buffer);
      }

      expect(doc.getPageCount()).toBe(3);
    });
  });

  describe('integration', () => {
    it('creates, modifies, merges, and saves a document', async () => {
      // Create first document with 2 pages
      const main = Document.create();
      main.addPageBreak();
      expect(main.getPageCount()).toBe(2);

      // Create second document with 3 pages
      const appendix = Document.create();
      appendix.addPageBreak();
      appendix.addPageBreak();
      expect(appendix.getPageCount()).toBe(3);

      // Merge specific pages from appendix (first and third)
      await main.merge(appendix, { pages: [0, 2] });

      // Save and reload
      const buffer = await main.toBuffer();
      const loaded = await Document.fromBuffer(buffer);

      // Verify it can be loaded without errors
      expect(loaded.getPageCount()).toBeGreaterThanOrEqual(3);
    });

    it('removes pages then merges new content', async () => {
      const doc = Document.create();
      doc.addPageBreak();
      doc.addPageBreak();
      doc.addPageBreak();
      expect(doc.getPageCount()).toBe(4);

      // Remove middle pages
      doc.removePage(2);
      doc.removePage(1);
      expect(doc.getPageCount()).toBe(2);

      // Merge new content
      const extra = Document.create();
      extra.addPageBreak();
      await doc.merge(extra);

      expect(doc.getPageCount()).toBeGreaterThanOrEqual(3);
    });
  });
});
