import { Marked, RendererObject } from "marked";
import * as fs from "fs/promises";
import type { FileHandle } from "fs/promises";
import * as path from "path";
import * as zip from "cross-zip";
import { Document, HeadingLevel, Packer, Paragraph } from "docx";
import { Dirent } from "fs";
import { docxRenderer, docxPreamble, DocxWriter } from "./DocxRenderer";

interface Portfolio {
  baseDir: string;
  files: MarkdownFile[];
}

interface Portfolio2 {
  baseDir: string;
  front: Page[];
  chapters: Chapter[];
  back: Page[];
}

interface Chapter {
  number: number;
  name: string;
  scenes: Scene[];
}

interface MarkdownFile {
  number?: number;
  name: string;
  filename: string;
  relativeDir: string;
  relativePath: string;
  markdownText: string;
  renderedHtml: string;
  renderedDocx: string;
}
type Scene = MarkdownFile;
type Page = MarkdownFile;

const filenamePattern = /^\W*([0-9]+)\W+(.*)$/;

const renderer: RendererObject = {
  heading(text, level) {
    return `<h${level + 2}>${text}</h${level + 2}>\n`;
  },
};

const markedDocx = new Marked().use({
  async: true,
  pedantic: false,
  gfm: true,
  renderer: docxRenderer,
});
// const markedHtml = new Marked().use({
//   async: true,
//   pedantic: false,
//   gfm: true,
//   renderer,
// });
const markedHtml = markedDocx;

async function loadChapter(dirEntry: Dirent): Promise<Chapter> {
  const { name, number } = parseName(dirEntry);
  if (number == null) {
    throw new Error(`Chapter does not include a number: "${name}"`);
  }
  const chapter: Chapter = {
    name,
    number,
    scenes: [],
  };
  if (dirEntry.isFile()) {
    chapter.scenes.push(await loadMarkdownFile(dirEntry));
  } else {
    for (const fileEntry of await getDirEntries(
      path.join(dirEntry.path, dirEntry.name),
    )) {
      chapter.scenes.push(await loadMarkdownFile(fileEntry));
    }
  }
  return chapter;
}

function parseName(dirEntry: Dirent): { name: string; number?: number } {
  const matches = dirEntry.name.match(filenamePattern);
  if (matches == null || matches.length < 3) {
    return { name: dirEntry.name };
  } else {
    return { name: matches[2], number: Number(matches[1]) };
  }
}

async function getMarkdownFiles(baseDir: string): Promise<Array<MarkdownFile>> {
  const markdownFiles: MarkdownFile[] = [];
  for (const entry of await fs.readdir(baseDir, {
    withFileTypes: true,
    recursive: true,
  })) {
    if (entry.isFile() && path.extname(entry.name).toLowerCase() === ".md") {
      markdownFiles.push(await loadMarkdownFile(entry));
    }
  }
  return markdownFiles;
}

async function loadMarkdownFile(fileEntry: Dirent): Promise<MarkdownFile> {
  if (!fileEntry.isFile()) {
    throw new Error(
      `Expected a file but found a directory: "${fileEntry.name}"`,
    );
  }
  const { name, number } = parseName(fileEntry);
  const filename = fileEntry.name;
  const relativeDir = fileEntry.path;
  const relativePath = path.join(fileEntry.path, fileEntry.name);
  const markdownText = await fs.readFile(relativePath, "utf8");
  const renderedHtml = (await markedHtml.parse(markdownText)).trim();
  const renderedDocx = (await markedDocx.parse(markdownText)).trim();
  return {
    name,
    number,
    filename,
    relativeDir,
    relativePath,
    markdownText,
    renderedHtml,
    renderedDocx,
  };
}

const loadPortfolio = async (): Promise<Portfolio> => {
  const baseDir = "test";
  const files = await getMarkdownFiles(baseDir);
  for (const file of files) {
    console.log(file.renderedHtml);
  }
  return {
    baseDir,
    files,
  };
};

const compilePortfolio = async (portfolio: Portfolio): Promise<string> => {
  const outputFilename = path.join(portfolio.baseDir, "out", "compiled.html");
  let file: FileHandle | undefined;
  try {
    file = await fs.open(outputFilename, "w");
    for (const markdownFile of portfolio.files) {
      await file.write(markdownFile.renderedHtml + "\n");
    }
  } finally {
    if (file != null) {
      await file.close();
    }
  }
  return "ok";
  // const doc = new Document({
  //   sections: [
  //     {
  //       children: [
  //         new Paragraph({ text: "hi", heading: HeadingLevel.HEADING_1 }),
  //       ],
  //     },
  //   ],
  // });
  // const packed = await Packer.toBuffer(doc);
  // const outputPath = path.join(portfolio.baseDir, "compiled.docx");
  // await writeFile(outputPath, packed);
  // return outputPath;
};

async function getDirEntries(dir: string): Promise<Array<Dirent>> {
  return fs.readdir(dir, { withFileTypes: true });
}

const loadPortfolio2 = async (): Promise<Portfolio2> => {
  const baseDir = "test";
  const frontFiles = await getDirEntries(path.join(baseDir, "1. Front"));
  const chaptersFiles = await getDirEntries(path.join(baseDir, "2. Chapters"));
  const backFiles = await getDirEntries(path.join(baseDir, "3. Back"));
  const front = await Promise.all(frontFiles.map(loadMarkdownFile));
  const chapters = await Promise.all(chaptersFiles.map(loadChapter));
  const back = await Promise.all(backFiles.map(loadMarkdownFile));
  return {
    baseDir,
    front,
    chapters,
    back,
  };
};

const compilePortfolio2 = async (portfolio: Portfolio2): Promise<string> => {
  const outputFilename = path.join(portfolio.baseDir, "out", "compiled.html");
  let file: FileHandle | undefined;
  try {
    file = await fs.open(outputFilename, "w");
    await file.write("<header>\n");
    for (const frontMatter of portfolio.front) {
      await file.write(`\n`);
      await file.write(` <section>\n`);
      await file.write(indent(frontMatter.renderedHtml, "  ") + "\n");
      await file.write(` </section>\n`);
    }
    await file.write(`\n`);
    await file.write("</header>\n");
    await file.write(`\n`);
    await file.write("<main>\n");
    for (const chapter of portfolio.chapters) {
      await file.write(`\n`);
      await file.write(` <section>\n`);
      await file.write(
        `  <h2>Chapter ${chapter.number}<br />${chapter.name}</h2>\n`,
      );
      let articleIndex = 0;
      for (const scene of chapter.scenes) {
        if (articleIndex++ > 0) {
          await file.write(`  <p>❖</p>\n`);
        }
        await file.write(`  <article>\n`);
        await file.write(indent(scene.renderedHtml, "   ") + "\n");
        await file.write(`  </article>\n`);
      }
      await file.write(` </section>\n`);
    }
    await file.write(`\n`);
    await file.write("</main>\n");
    await file.write(`\n`);
    await file.write("<footer>\n");
    for (const backMatter of portfolio.back) {
      await file.write(`\n`);
      await file.write(` <section>\n`);
      await file.write(indent(backMatter.renderedHtml, "  ") + "\n");
      await file.write(` </section>\n`);
    }
    await file.write(`\n`);
    await file.write("</footer>\n");
  } finally {
    if (file != null) {
      await file.close();
    }
  }
  return "ok";
};

const compilePortfolio3 = async (portfolio: Portfolio2): Promise<string> => {
  const scratchPath = path.join(portfolio.baseDir, "out", "docx");
  const docxOutputPath = path.join(portfolio.baseDir, "out", "compiled.docx");
  await fs.mkdir(path.join(scratchPath, "_rels"), { recursive: true });
  await fs.mkdir(path.join(scratchPath, "word", "_rels"), { recursive: true });
  await fs.writeFile(
    path.join(scratchPath, "_rels", ".rels"),
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/>`,
  );
  await fs.writeFile(
    path.join(scratchPath, "[Content_Types].xml"),
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`,
  );
  await fs.writeFile(
    path.join(scratchPath, "word", "_rels", "document.xml.rels"),
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/>`,
  );

  const outputFilename = path.join(scratchPath, "word", "document.xml");
  let file: FileHandle | undefined;
  try {
    file = await fs.open(outputFilename, "w");
    const docx = new DocxWriter(file);
    await docx.start();
    for (const frontMatter of portfolio.front) {
      await docx.write(frontMatter.renderedDocx);
    }
    for (const chapter of portfolio.chapters) {
      // await file.write(chapter.title);
      for (const scene of chapter.scenes) {
        await docx.write(scene.renderedDocx);
      }
    }
    for (const backMatter of portfolio.back) {
      await docx.write(indent(backMatter.renderedDocx, "  ") + "\n");
    }
    await docx.end();
  } finally {
    if (file != null) {
      await file.close();
    }
  }
  console.log({ scratchPath, docxOutputPath });
  zip.zipSync(scratchPath, docxOutputPath);
  return "ok";
};

function indent(text: string, indention: string) {
  return indention + text.trim().replaceAll("\n", `\n${indention}`);
}

function dedent(text: string): string {
  const trimmed = text.replace(/^\n+/, "");
  const indentMatch = trimmed.match(/^(\s+)/) || [];
  const indention = indentMatch[1];
  return ("\n" + trimmed).replaceAll("\n" + indention, "\n").trim();
}

loadPortfolio2().then(compilePortfolio3).then(console.log);
