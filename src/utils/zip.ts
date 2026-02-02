import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

export type ZipFiles = Map<string, Uint8Array>;

/**
 * Reads a ZIP file and returns a map of path -> content
 * @param data - ZIP file as Uint8Array
 * @returns Promise resolving to a map of file paths to contents
 */
export const readZip = (data: Uint8Array): Promise<ZipFiles> => {
  const result = unzipSync(data);
  const files = new Map<string, Uint8Array>();
  for (const [path, content] of Object.entries(result)) {
    // Skip directory entries
    if (!path.endsWith('/') && content && content.length > 0) {
      files.set(path, content);
    }
  }
  return Promise.resolve(files);
};

/**
 * Creates a ZIP file from a map of path -> content
 * @param files - Map of file paths to contents
 * @returns Promise resolving to ZIP file as Uint8Array
 */
export const writeZip = (files: ZipFiles): Promise<Uint8Array> => {
  const zipData: Record<string, Uint8Array> = {};
  for (const [path, content] of files) {
    // Skip directory entries (ending with /) and undefined/empty content
    if (path.endsWith('/') || !content || content.length === 0) {
      continue;
    }
    zipData[path] = content;
  }
  const result = zipSync(zipData);
  return Promise.resolve(result);
};

/**
 * Reads a file from the ZIP as a UTF-8 string
 */
export const readZipText = (files: ZipFiles, path: string): string | undefined => {
  const data = files.get(path);
  if (!data) return undefined;
  return strFromU8(data);
};

/**
 * Writes a UTF-8 string to the ZIP files map
 */
export const writeZipText = (files: ZipFiles, path: string, content: string): void => {
  files.set(path, strToU8(content));
};
