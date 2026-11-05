import { describe, expect, it } from "vitest";

import { Formatter } from "@export/formatter";
import { TextRun } from "@file/paragraph/run";
import { BuilderElement } from "@file/xml-components";

import { DeletedRuns } from "./deleted-runs";

const tree = (children: readonly ConstructorParameters<typeof DeletedRuns>[0]["children"][number][]) =>
    new Formatter().format(new DeletedRuns({ id: 1, author: "Ann", date: "2024-01-01T00:00:00Z", children }));

describe("DeletedRuns", () => {
    it("writes its runs in one w:del, with their text as w:delText", () => {
        expect(tree([new TextRun("one"), new TextRun({ text: "two", bold: true })])).to.deep.equal({
            "w:del": [
                { _attr: { "w:id": 1, "w:author": "Ann", "w:date": "2024-01-01T00:00:00Z" } },
                { "w:r": [{ "w:delText": [{ _attr: { "xml:space": "preserve" } }, "one"] }] },
                { "w:r": [{ "w:rPr": [{ "w:b": {} }, { "w:bCs": {} }] }, { "w:delText": [{ _attr: { "xml:space": "preserve" } }, "two"] }] },
            ],
        });
    });

    it("writes a field's instructions as w:delInstrText", () => {
        const instruction = new BuilderElement({ name: "w:instrText" });
        const run = new TextRun({ children: [instruction as never] });
        expect(JSON.stringify(tree([run]))).to.include("w:delInstrText").and.not.include('"w:instrText"');
    });

    it("leaves out the date when there is none", () => {
        const xml = new Formatter().format(new DeletedRuns({ id: 1, author: "Ann", children: [] }));
        expect(xml).to.deep.equal({ "w:del": { _attr: { "w:id": 1, "w:author": "Ann" } } });
    });

    it("leaves a w:t attribute named like the element alone", () => {
        const element = new BuilderElement({ name: "w:other", attributes: { value: { key: "w:t", value: "1" } } });
        expect(tree([element])).to.deep.include({
            "w:del": [{ _attr: { "w:id": 1, "w:author": "Ann", "w:date": "2024-01-01T00:00:00Z" } }, { "w:other": { _attr: { "w:t": "1" } } }],
        });
    });
});
