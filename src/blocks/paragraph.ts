
    /**
     * Paragraph builder.
     * @module
     */
    import { escapeXML, ptToHalfPoints } from "../core/utils.ts";
    import type { ParagraphOptions } from "../core/docxaur.ts";
    import { Element } from "./section.ts";

    interface TextRun { text: string; style?: ParagraphOptions; }
    type ParagraphOperation = () => void;

    export class Paragraph extends Element {
      private runs: TextRun[] = [];
      private options: ParagraphOptions;
      private operations: ParagraphOperation[] = [];
      private isBuilt = false;

      constructor(options: ParagraphOptions = {}) { super(); this.options = options; }

      text(text: string, style?: ParagraphOptions): this { this.operations.push(() => this.runs.push({ text, style })); return this; }
      tab(): this { this.operations.push(() => this.runs.push({ text: "	" })); return this; }
      lineBreak(count: number = 1): this { this.operations.push(() => { for (let i = 0; i < count; i++) this.runs.push({ text: "
" }); }); return this; }
      pageBreak(count: number = 1): this { this.operations.push(() => { for (let i = 0; i < count; i++) this.runs.push({ text: "[PAGE_BREAK]" }); }); return this; }

      /** @deprecated Prefer explicit method calls instead of apply(). */
      apply(...operations: ((builder: this) => this)[]): this {
        console.warn("Paragraph.apply() is deprecated. Use direct method calls instead.");
        for (const op of operations) op(this); return this;
      }

      private hasRunProperties(style: ParagraphOptions): boolean {
        return !!(style.bold || style.italic || style.underline || style.fontSize || style.fontColor || style.fontName);
      }
      private build(): void { if (this.isBuilt) return; this.isBuilt = true; for (const op of this.operations) op(); }

      toXML(): string {
        this.build();
        const align = this.options.align ?? "left";
        const breaksBefore = this.options.breakBefore ?? 0; const breaksAfter = this.options.breakAfter ?? 0;
        let xml = "  <w:p>
";
        xml += "    <w:pPr>
";
        if (align !== "left") { const jc = align === "justify" ? "both" : align; xml += `      <w:jc w:val="${jc}"/>
`; }
        const spacing = this.options.spacing; const before = spacing?.before ? ptToHalfPoints(spacing.before) * 20 : 0; const after = spacing?.after ? ptToHalfPoints(spacing.after) * 20 : 0; const line = spacing?.line ? Math.round(spacing.line * 240) : 240;
        xml += `      <w:spacing w:after="${after}" w:before="${before}" w:line="${line}" w:lineRule="auto"/>
`;
        xml += "    </w:pPr>
";
        for (let i = 0; i < breaksBefore; i++) xml += "    <w:r><w:br/></w:r>
";
        for (const run of this.runs) {
          if (run.text === "	") xml += "    <w:r><w:tab/></w:r>
";
          else if (run.text === "
") xml += "    <w:r><w:br/></w:r>
";
          else if (run.text === "[PAGE_BREAK]") xml += '    <w:r><w:br w:type="page"/></w:r>
';
          else {
            xml += "    <w:r>
";
            const style = run.style;
            if (style && this.hasRunProperties(style)) {
              xml += "      <w:rPr>
";
              if (style.bold)      xml += "        <w:b/>
";
              if (style.italic)    xml += "        <w:i/>
";
              if (style.underline) xml += '        <w:u w:val="single"/>
';
              if (style.fontSize)  xml += `        <w:sz w:val="${ptToHalfPoints(style.fontSize)}"/>
`;
              if (style.fontColor) xml += `        <w:color w:val="${style.fontColor}"/>
`;
              if (style.fontName)  xml += `        <w:rFonts w:ascii="${style.fontName}" w:hAnsi="${style.fontName}"/>
`;
              xml += "      </w:rPr>
";
            }
            xml += `      <w:t xml:space="preserve">${escapeXML(run.text)}</w:t>
`;
            xml += "    </w:r>
";
          }
        }
        for (let i = 0; i < breaksAfter; i++) xml += "    <w:r><w:br/></w:r>
";
        xml += "  </w:p>"; return xml;
      }
    }
