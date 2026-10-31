import { describe, expect, it } from "vitest";

import { Formatter } from "@export/formatter";

import { createBlip } from "./blip";

describe("createBlip", () => {
    const mediaData = {
        type: "png" as const,
        data: new Uint8Array([1]),
        fileName: "test.png",
        transformation: { pixels: { x: 1, y: 1 }, emus: { x: 9525, y: 9525 } },
    };

    it("should write no effect without a luminance", () => {
        expect(new Formatter().format(createBlip(mediaData))).to.deep.equal({
            "a:blip": { _attr: { "r:embed": "rId{test.png}", cstate: "none" } },
        });
    });

    it("should write the luminance in thousandths of a percent", () => {
        expect(new Formatter().format(createBlip(mediaData, { brightness: 70, contrast: -70 }))).to.deep.equal({
            "a:blip": [
                { _attr: { "r:embed": "rId{test.png}", cstate: "none" } },
                { "a:lum": { _attr: { bright: 70000, contrast: -70000 } } },
            ],
        });
    });

    it("should leave out a luminance of no change, and a part of it that is 0", () => {
        expect(new Formatter().format(createBlip(mediaData, { brightness: 0, contrast: 0 }))).to.deep.equal({
            "a:blip": { _attr: { "r:embed": "rId{test.png}", cstate: "none" } },
        });
        expect(new Formatter().format(createBlip(mediaData, { brightness: 30 }))).to.deep.equal({
            "a:blip": [{ _attr: { "r:embed": "rId{test.png}", cstate: "none" } }, { "a:lum": { _attr: { bright: 30000 } } }],
        });
    });
});
