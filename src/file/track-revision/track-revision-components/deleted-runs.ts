/**
 * Deleted runs module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtrackChanges.php
 *
 * @module
 */
import { type IContext, type IXmlableObject, XmlComponent } from "@file/xml-components";

import { ChangeAttributes, type IChangedAttributesProperties } from "../track-revision";

const DELETED_NAMES: Readonly<Record<string, string>> = { "w:t": "w:delText", "w:instrText": "w:delInstrText" };

/** The XML with its text as deleted text, at any depth: a text box in a deleted run has its text deleted with it. */
const asDeleted = (xml: unknown): unknown => {
    if (Array.isArray(xml)) {
        return xml.map(asDeleted);
    }
    if (typeof xml !== "object" || xml === null) {
        return xml;
    }
    return Object.fromEntries(
        Object.entries(xml).map(([name, value]) => [DELETED_NAMES[name] ?? name, name === "_attr" ? value : asDeleted(value)]),
    );
};

/**
 * Options for {@link DeletedRuns}.
 */
export type IDeletedRunsOptions = Omit<IChangedAttributesProperties, "date"> & {
    /** Date and time when the change was made (ISO 8601 format) */
    readonly date?: string;
    /** The runs that were deleted, of any kind: text, fields, pictures, or hyperlinks */
    readonly children: readonly XmlComponent[];
};

/**
 * Runs deleted as a tracked change, written in one `w:del`.
 *
 * Where {@link DeletedTextRun} makes a single deleted run from run options, this takes runs as they are. Their text is
 * written as `w:delText` and their field instructions as `w:delInstrText`, as the schema has it inside a deletion.
 *
 * @example
 * ```typescript
 * new DeletedRuns({ id: 1, author: "Ann", date: "2024-01-01T00:00:00Z", children: [new TextRun("gone")] });
 * ```
 */
export class DeletedRuns extends XmlComponent {
    public constructor({ id, author, date, children }: IDeletedRunsOptions) {
        super("w:del");
        this.root.push(new ChangeAttributes({ id, author, date: date as string }));
        for (const child of children) {
            this.root.push(child);
        }
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        return asDeleted(super.prepForXml(context)) as IXmlableObject | undefined;
    }
}
