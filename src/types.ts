/**
 * Options for merging documents
 */
export interface MergeOptions {
  /**
   * Page indices to include (0-based).
   * If omitted, all pages are merged.
   * @example [0, 2, 4] - merge first, third, and fifth pages
   */
  pages?: number[];

  /**
   * Insert page break before merged content.
   * @default true
   */
  addPageBreakBefore?: boolean;
}

/**
 * Information about a page in the document
 */
export interface PageInfo {
  /** 0-based page index */
  index: number;

  /** Index of first body element in this page */
  startElement: number;

  /** Index of last body element in this page (inclusive) */
  endElement: number;
}
