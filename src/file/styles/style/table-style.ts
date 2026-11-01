/**
 * Table style module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPstyleTableStyles.php
 *
 * @module
 */
import { type IParagraphStylePropertiesOptions, type IRunStylePropertiesOptions, ParagraphProperties } from "@file/paragraph";
import { RunProperties } from "@file/paragraph/run/properties";
import { type ITablePropertiesOptionsBase, TableProperties } from "@file/table/table-properties";

import { type IStyleOptions, Style } from "./style";

/**
 * Options for a table style.
 *
 * @property id - Unique identifier of the table style
 * @property table - The table properties of tables in the style (`w:tblPr`): their width, alignment, indent, borders,
 * shading, layout and cell margins
 * @property paragraph - Paragraph properties of the text in the style's tables (`w:pPr`)
 * @property run - Run properties of the text in the style's tables (`w:rPr`)
 */
export type ITableStyleOptions = {
    /** Unique identifier of the table style */
    readonly id: string;
    /** The table properties of tables in the style */
    readonly table?: Omit<ITablePropertiesOptionsBase, "style" | "float" | "tableLook">;
    /** Paragraph properties of the text in the style's tables */
    readonly paragraph?: IParagraphStylePropertiesOptions;
    /** Run properties of the text in the style's tables */
    readonly run?: IRunStylePropertiesOptions;
} & IStyleOptions;

/**
 * A table style (`w:style w:type="table"`): the table properties, and the paragraph and run properties of the text in
 * the tables that use it. Its borders are the sides given only.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Style">
 *   <xsd:sequence>
 *     ...
 *     <xsd:element name="pPr" type="CT_PPrGeneral" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="rPr" type="CT_RPr" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblPr" type="CT_TblPrBase" minOccurs="0" maxOccurs="1"/>
 *     ...
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * new StyleForTable({
 *   id: "PlainGrid",
 *   name: "Plain Grid",
 *   basedOn: "TableNormal",
 *   table: { borders: { top: { style: BorderStyle.SINGLE, size: 4, color: "auto" } } },
 *   run: { size: 20 },
 * });
 * ```
 */
export class StyleForTable extends Style {
    public constructor(options: ITableStyleOptions) {
        super({ type: "table", styleId: options.id }, options);

        if (options.paragraph) {
            this.root.push(new ParagraphProperties(options.paragraph, { implicitListParagraphStyle: false, inListOrder: false }));
        }
        if (options.run) {
            this.root.push(new RunProperties(options.run));
        }
        this.root.push(new TableProperties(options.table ?? {}, true));
    }
}
