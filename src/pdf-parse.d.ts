type PdfParseResult = { text: string };

declare module "pdf-parse" {
  const pdfParse: (buffer: Buffer) => Promise<PdfParseResult>;
  export default pdfParse;
}

declare module "pdf-parse/lib/pdf-parse.js" {
  const pdfParse: (buffer: Buffer) => Promise<PdfParseResult>;
  export default pdfParse;
}
