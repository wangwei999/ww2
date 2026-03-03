declare module 'officegen' {
  import { Writable } from 'stream';

  interface OfficegenOptions {
    'title'?: string;
    'subject'?: string;
    'keywords'?: string;
    'description'?: string;
    'creator'?: string;
    'lastModifiedBy'?: string;
    'created'?: Date;
    'modified'?: Date;
    'category'?: string;
  }

  interface DocxObj {
    createP(options?: { align?: string }): ParagraphObj;
    createTable(data: any[][], options?: any): void;
    on(event: 'finalize', callback: (written: number) => void): void;
    on(event: 'error', callback: (err: Error) => void): void;
    generate(output: Writable): void;
  }

  interface ParagraphObj {
    addText(text: string, options?: any): void;
    addHorizontalLine(): void;
  }

  function officegen(type: 'docx', options?: OfficegenOptions): DocxObj;

  export = officegen;
}
