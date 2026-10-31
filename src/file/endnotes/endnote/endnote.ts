import { FirstNoteParagraph } from "@file/footnotes/footnote/note-paragraph";
import type { Paragraph } from "@file/paragraph";
import { XmlComponent } from "@file/xml-components";

import { EndnoteAttributes } from "./endnote-attributes";
import { EndnoteRefRun } from "./run/endnote-ref-run";

export const EndnoteType = {
    SEPARATOR: "separator",

    CONTINUATION_SEPARATOR: "continuationSeparator",

    CONTINUATION_NOTICE: "continuationNotice",
} as const;

export type IEndnoteOptions = {
    readonly id: number;
    readonly type?: (typeof EndnoteType)[keyof typeof EndnoteType];
    readonly children: readonly Paragraph[];
    /** The paragraphs hold the endnote's mark themselves (an {@link EndnoteRefRun}), so the first isn't given one */
    readonly ownMark?: boolean;
};

export class Endnote extends XmlComponent {
    public constructor(options: IEndnoteOptions) {
        super("w:endnote");
        this.root.push(
            new EndnoteAttributes({
                type: options.type,
                id: options.id,
            }),
        );

        // A separator or a continuation notice has no mark
        const marked = options.type === undefined && !options.ownMark;
        for (const [index, child] of options.children.entries()) {
            this.root.push(index === 0 && marked ? new FirstNoteParagraph(child, new EndnoteRefRun()) : child);
        }
    }
}
