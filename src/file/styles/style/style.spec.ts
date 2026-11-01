import { describe, expect, it } from "vitest";

import { Formatter } from "@export/formatter";

import { Style } from "./style";

describe("Style", () => {
    describe("#constructor()", () => {
        it("should set the given properties", () => {
            const style = new Style(
                {
                    type: "paragraph",
                    styleId: "myStyleId",
                    default: true,
                },
                {},
            );
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": { _attr: { "w:type": "paragraph", "w:styleId": "myStyleId", "w:default": true } },
            });
        });

        it("should set the name of the style, if given", () => {
            const style = new Style(
                {
                    type: "paragraph",
                    styleId: "myStyleId",
                },
                { name: "Style Name" },
            );
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:type": "paragraph", "w:styleId": "myStyleId" } },
                    { "w:name": { _attr: { "w:val": "Style Name" } } },
                ],
            });
        });

        it("should write the aliases after the name", () => {
            const style = new Style(
                { type: "paragraph", styleId: "Heading1" },
                { name: "heading 1", aliases: "Chapter", basedOn: "Normal" },
            );
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:type": "paragraph", "w:styleId": "Heading1" } },
                    { "w:name": { _attr: { "w:val": "heading 1" } } },
                    { "w:aliases": { _attr: { "w:val": "Chapter" } } },
                    { "w:basedOn": { _attr: { "w:val": "Normal" } } },
                ],
            });
        });

        it("should take default and customStyle from the options", () => {
            const style = new Style({ type: "paragraph", styleId: "Body" }, { name: "Body", default: true, customStyle: true });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:type": "paragraph", "w:styleId": "Body", "w:default": true, "w:customStyle": "1" } },
                    { "w:name": { _attr: { "w:val": "Body" } } },
                ],
            });
        });

        it("should leave out default and customStyle when they are off", () => {
            const style = new Style({ type: "paragraph", styleId: "Body" }, { default: false, customStyle: false });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({ "w:style": { _attr: { "w:type": "paragraph", "w:styleId": "Body" } } });
        });
    });
});
