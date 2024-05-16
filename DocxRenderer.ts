import * as fs from "fs/promises";
import { RendererObject } from "marked";

export const docxPreamble = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
<w:body>
`.trim();
export const docxPostamble = `
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
  <w:cols w:space="720"/>
  <w:docGrid w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>
`.trim();

export const docxRenderer: RendererObject = {
  heading: (text: string, level: number) => {
    console.log("heading", { text, level });
    return wp(wpprHeading(level), wr(wt(text)));
  },
  paragraph: (text: string) => {
    console.log("paragraph", { text });
    return wp(wr(wt(text)));
  },
};

function wp(...inner: string[]): string {
  return `\n<w:p>\n${inner.join("")}\n</w:p>`;
}
function wpprHeading(level: number): string {
  return `<w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>\n`;
}
function wr(inner: string): string {
  return `<w:r>${inner}</w:r>`;
}
function wt(inner: string): string {
  return `<w:t>${inner}</w:t>`;
}

export class DocxWriter {
  constructor(private readonly file: fs.FileHandle) {}
  write(data: string) {
    return this.file.write("\n" + data);
  }
  start() {
    return this.file.write(docxPreamble);
  }
  end() {
    return this.write(docxPostamble);
  }
}
