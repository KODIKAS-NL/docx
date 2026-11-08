import { describe, expect, it } from "vitest";

import { Formatter } from "@export/formatter";

import { People } from "./people";

describe("People", () => {
    it("writes each person with the presence info when it has a provider and a user id", () => {
        const tree = new Formatter().format(
            new People([
                { author: "Ann Author", providerId: "None", userId: "ann" },
                { author: "Bob", providerId: "AD" },
            ]),
        );
        expect(tree).to.deep.equal({
            "w15:people": [
                {
                    _attr: {
                        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                        "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
                        "mc:Ignorable": "w15",
                    },
                },
                {
                    "w15:person": [
                        { _attr: { "w15:author": "Ann Author" } },
                        { "w15:presenceInfo": { _attr: { "w15:providerId": "None", "w15:userId": "ann" } } },
                    ],
                },
                { "w15:person": { _attr: { "w15:author": "Bob" } } },
            ],
        });
    });
});
