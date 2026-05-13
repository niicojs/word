import path from 'path';
import { Document } from '../src';

const folder = import.meta.dirname;

const doc = await Document.fromFile(path.join(folder, 'document.docx'));

doc.render({
  param1: 'Hello, World!',
  param2: 'This is a templated document.',
  up_left: 'Top Left Value',
  down_right: 'Bottom Right Value',
  middle: 'Center Value',
});

await doc.toFile(path.join(folder, 'output.docx'));
console.log('Saved templated document to output.docx');
