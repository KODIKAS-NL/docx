import { describe, expect, it } from "vitest";

import { Formatter } from "@export/formatter";
import { BorderStyle } from "@file/border";
import { WidthType } from "@file/table";

import { StyleForTable } from "./table-style";

describe("StyleForTable", () => {
    it("should write a table style with its table, paragraph and run properties", () => {
        const style = new StyleForTable({
            id: "PlainGrid",
            name: "Plain Grid",
            basedOn: "TableNormal",
            semiHidden: true,
            table: {
                indent: { size: 0, type: WidthType.DXA },
                borders: { top: { style: BorderStyle.SINGLE, size: 4, color: "auto" } },
                cellMargin: { left: 108, right: 108 },
            },
            paragraph: { spacing: { after: 0 } },
            run: { size: 20 },
        });
        const tree = new Formatter().format(style);
        expect(tree).to.deep.equal({
            "w:style": [
                { _attr: { "w:type": "table", "w:styleId": "PlainGrid" } },
                { "w:name": { _attr: { "w:val": "Plain Grid" } } },
                { "w:basedOn": { _attr: { "w:val": "TableNormal" } } },
                { "w:semiHidden": {} },
                { "w:pPr": [{ "w:spacing": { _attr: { "w:after": 0 } } }] },
                { "w:rPr": [{ "w:sz": { _attr: { "w:val": 20 } } }, { "w:szCs": { _attr: { "w:val": 20 } } }] },
                {
                    "w:tblPr": [
                        { "w:tblInd": { _attr: { "w:type": "dxa", "w:w": 0 } } },
                        { "w:tblBorders": [{ "w:top": { _attr: { "w:val": "single", "w:sz": 4, "w:color": "auto" } } }] },
                        {
                            "w:tblCellMar": [
                                { "w:left": { _attr: { "w:type": "dxa", "w:w": 108 } } },
                                { "w:right": { _attr: { "w:type": "dxa", "w:w": 108 } } },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should write only the flags and properties given", () => {
        const tree = new Formatter().format(new StyleForTable({ id: "Empty", name: "Empty" }));
        expect(tree).to.deep.equal({
            "w:style": [{ _attr: { "w:type": "table", "w:styleId": "Empty" } }, { "w:name": { _attr: { "w:val": "Empty" } } }],
        });
    });
});
