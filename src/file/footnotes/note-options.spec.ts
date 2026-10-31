import { describe, expect, it } from "vitest";

import { Formatter } from "@export/formatter";
import { Packer } from "@export/packer/packer";
import { EndnoteRefRun, EndnoteReferenceRun } from "@file/endnotes";
import { File } from "@file/file";
import { Paragraph, TextRun } from "@file/paragraph";
import { textOf, unzip } from "tests/zip";

import { FootnoteRefRun, FootnoteReferenceRun, SeperatorRun } from "./footnote/run";

/** The parts of a document, packed */
const partsOf = async (document: File): Promise<(path: string) => string> => {
    const files = unzip(await Packer.toBuffer(document));
    return (path) => textOf(files, path)!;
};

const documentWith = (children: readonly (TextRun | FootnoteReferenceRun | EndnoteReferenceRun)[], options = {}): File =>
    new File({ ...options, sections: [{ children: [new Paragraph({ children: [...children] })] }] });

describe("note references with run options", () => {
    it("should format the reference run, keeping its style unless it names another", () => {
        expect(new Formatter().format(new FootnoteReferenceRun(1, { bold: true }))).to.deep.equal({
            "w:r": [
                { "w:rPr": [{ "w:rStyle": { _attr: { "w:val": "FootnoteReference" } } }, { "w:b": {} }, { "w:bCs": {} }] },
                { "w:footnoteReference": { _attr: { "w:id": 1 } } },
            ],
        });
        expect(new Formatter().format(new EndnoteReferenceRun(2, { style: "Mark" }))).to.deep.equal({
            "w:r": [{ "w:rPr": [{ "w:rStyle": { _attr: { "w:val": "Mark" } } }] }, { "w:endnoteReference": { _attr: { "w:id": 2 } } }],
        });
    });

    it("should write a custom mark as the run's text after the reference", () => {
        expect(new Formatter().format(new FootnoteReferenceRun(1, { customMark: "†" }))).to.deep.equal({
            "w:r": [
                { "w:rPr": [{ "w:rStyle": { _attr: { "w:val": "FootnoteReference" } } }] },
                { "w:footnoteReference": { _attr: { "w:id": 1, "w:customMarkFollows": 1 } } },
                { "w:t": [{ _attr: { "xml:space": "preserve" } }, "†"] },
            ],
        });
    });

    it("should write the custom mark of a reference with its note's content", async () => {
        const read = await partsOf(
            documentWith([
                new TextRun("Text"),
                new FootnoteReferenceRun({ children: [new Paragraph("Note")] }, { customMark: "*" }),
                new EndnoteReferenceRun({ children: [new Paragraph("End")] }, { customMark: "§" }),
            ]),
        );
        const document = read("word/document.xml");
        expect(document).to.contain(`<w:footnoteReference w:id="1" w:customMarkFollows="1"/><w:t xml:space="preserve">*</w:t>`);
        expect(document).to.contain(`<w:endnoteReference w:id="1" w:customMarkFollows="1"/><w:t xml:space="preserve">§</w:t>`);
    });
});

describe("notes with their own mark", () => {
    it("should not give the first paragraph a mark when the paragraphs hold it", async () => {
        const read = await partsOf(
            documentWith([
                new FootnoteReferenceRun({
                    ownMark: true,
                    children: [new Paragraph({ children: [new TextRun("Before "), new FootnoteRefRun({ italics: true })] })],
                }),
                new EndnoteReferenceRun({ ownMark: true, children: [new Paragraph({ children: [new EndnoteRefRun()] })] }),
            ]),
        );
        const footnotes = read("word/footnotes.xml");
        expect(footnotes.match(/<w:footnoteRef\/>/g)).to.have.length(1);
        expect(footnotes).to.contain(
            `<w:t xml:space="preserve">Before </w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:i/><w:iCs/></w:rPr><w:footnoteRef/>`,
        );
        expect(read("word/endnotes.xml").match(/<w:endnoteRef\/>/g)).to.have.length(1);
    });

    it("should give a note without one its mark first", async () => {
        const read = await partsOf(documentWith([new FootnoteReferenceRun({ children: [new Paragraph("Note")] })]));
        expect(read("word/footnotes.xml")).to.match(
            /<w:footnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteRef\/>/,
        );
    });
});

describe("separators", () => {
    it("should have no mark in the default separators", async () => {
        const read = await partsOf(documentWith([new FootnoteReferenceRun({ children: [new Paragraph("Note")] })]));
        const separators = read("word/footnotes.xml").match(/<w:footnote w:type="[^"]+" w:id="-?\d+">.*?<\/w:footnote>/g)!;
        expect(separators).to.have.length(2);
        for (const separator of separators) {
            expect(separator).not.to.contain("footnoteRef");
        }
    });

    it("should write the separators given in place of the default ones, and a continuation notice", async () => {
        const read = await partsOf(
            documentWith([new TextRun("Text"), new FootnoteReferenceRun({ children: [new Paragraph("Note")] })], {
                footnoteSeparators: {
                    separator: [new Paragraph({ style: "Separator", children: [new SeperatorRun({ size: 16 })] })],
                    continuationNotice: [new Paragraph("Continued")],
                },
                endnoteSeparators: { continuationSeparator: [new Paragraph("Endnotes go on")] },
            }),
        );
        const footnotes = read("word/footnotes.xml");
        expect(footnotes).to.contain(
            `<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:pStyle w:val="Separator"/></w:pPr><w:r><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:separator/></w:r></w:p></w:footnote>`,
        );
        expect(footnotes).to.contain(`<w:footnote w:type="continuationSeparator" w:id="0">`);
        expect(footnotes).to.contain(
            `<w:footnote w:type="continuationNotice" w:id="1"><w:p><w:r><w:t xml:space="preserve">Continued</w:t></w:r></w:p></w:footnote>`,
        );
        // the notice before the notes, which take the ids after it
        expect(footnotes.indexOf("continuationNotice")).to.be.lessThan(footnotes.indexOf(`w:id="2"`));
        expect(read("word/document.xml")).to.contain(`<w:footnoteReference w:id="2"/>`);
        expect(read("word/endnotes.xml")).to.contain(
            `<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:t xml:space="preserve">Endnotes go on</w:t></w:r></w:p></w:endnote>`,
        );
    });
});

describe("endnotePosition", () => {
    it("should write the endnotes' position in the settings, before the compatibility settings", async () => {
        const read = await partsOf(documentWith([new TextRun("Text")], { endnotePosition: "sectEnd" }));
        expect(read("word/settings.xml")).to.match(/<w:endnotePr><w:pos w:val="sectEnd"\/><\/w:endnotePr><w:compat>/);
    });

    it("should write none by default", async () => {
        const read = await partsOf(documentWith([new TextRun("Text")]));
        expect(read("word/settings.xml")).not.to.contain("endnotePr");
    });
});
