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

/**
 * Value types supported by document templating
 */
export type TemplateValue = string | number | boolean | null | undefined;

/**
 * Data used to replace template placeholders
 */
export type TemplateData = Record<string, TemplateValue>;

/**
 * Options for templating
 */
export interface TemplateOptions {
  /**
   * Placeholder pattern. Must contain one capture group for the key.
   * @default /\{([a-zA-Z0-9_.-]+)\}/g
   */
  pattern?: RegExp;

  /**
   * Remove placeholders when the key is missing.
   * @default false
   */
  removeMissing?: boolean;

  /**
   * Optional value transformer applied before insertion.
   */
  transform?: (value: TemplateValue, key: string) => string;
}
