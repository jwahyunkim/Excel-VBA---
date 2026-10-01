import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const [inputPath, outputPath] = process.argv.slice(2);
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const preview = await workbook.render({sheetName: 'config', range: 'I7:M30', scale: 1.5, format: 'png'});
await fs.writeFile(outputPath, new Uint8Array(await preview.arrayBuffer()));
console.log(`Rendered ${outputPath}`);
