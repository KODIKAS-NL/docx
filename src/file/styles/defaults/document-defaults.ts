/**
 * Document defaults module for WordprocessingML styles.
 *
 * Reference: http://officeopenxml.com/WPstyles.php
 *
 * @module
 */
import type { IParagraphStylePropertiesOptions } from "@file/paragraph/properties";
import type { IRunStylePropertiesOptions } from "@file/paragraph/run/properties";
import { XmlComponent } from "@file/xml-components";

import { ParagraphPropertiesDefaults } from "./paragraph-properties";
import { RunPropertiesDefaults } from "./run-properties";

/**
 * Options for configuring document-wide default formatting.
 *
 * @property paragraph - Default paragraph properties applied to all paragraphs
 * @property run - Default run properties applied to all text runs
 */
export type IDocumentDefaultsOptions = {
    /** Default paragraph properties applied to all paragraphs */
    readonly paragraph?: IParagraphStylePropertiesOptions;
    /** Default run properties applied to all text runs; `null` leaves out `w:rPrDefault`, which Word treats otherwise than an empty one */
    readonly run?: IRunStylePropertiesOptions | null;
};

/**
 * Represents document-wide default formatting in a WordprocessingML document.
 *
 * Document defaults define the base formatting properties that apply to all
 * paragraphs and runs in a document unless overridden by a style or direct formatting.
 *
 * Reference: http://officeopenxml.com/WPstyles.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_DocDefaults">
 *   <xsd:sequence>
 *     <xsd:element name="rPrDefault" type="CT_RPrDefault" minOccurs="0"/>
 *     <xsd:element name="pPrDefault" type="CT_PPrDefault" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Set default font and spacing for the document
 * new DocumentDefaults({
 *   run: {
 *     font: "Calibri",
 *     size: 22
 *   },
 *   paragraph: {
 *     spacing: { after: 200, line: 276 }
 *   }
 * });
 * ```
 */
export class DocumentDefaults extends XmlComponent {
    private readonly paragraphPropertiesDefaults: ParagraphPropertiesDefaults;

    public constructor(options: IDocumentDefaultsOptions) {
        super("w:docDefaults");

        this.paragraphPropertiesDefaults = new ParagraphPropertiesDefaults(options.paragraph);

        if (options.run !== null) {
            this.root.push(new RunPropertiesDefaults(options.run));
        }
        this.root.push(this.paragraphPropertiesDefaults);
    }
}
