/**
 * Numbering module for WordprocessingML documents.
 *
 * Numbering provides support for numbered and bulleted lists.
 *
 * Reference: http://officeopenxml.com/WPnumbering.php
 *
 * @see https://stackoverflow.com/questions/58622437/purpose-of-abstractnum-and-numberinginstance
 *
 * @module
 */
import { AlignmentType } from "@file/paragraph";
import { type IContext, type IXmlableObject, XmlComponent } from "@file/xml-components";
import { abstractNumUniqueNumericIdGen, concreteNumUniqueNumericIdGen, convertInchesToTwip } from "@util/convenience-functions";
import { decimalNumber } from "@util/values";

import { AbstractNumbering } from "./abstract-numbering";
import { type ILevelsOptions, LevelFormat } from "./level";
import { listKey } from "./list-key";
import { ConcreteNumbering } from "./num";
import { DocumentAttributes } from "../document/document-attributes";

/**
 * A list's definition: its reference and its levels.
 *
 * A document's `numbering` option has a definition for each of its lists. A paragraph can also be given a list's
 * definition as its `numbering`'s `reference`, for a list the document's `numbering` option doesn't have, which is added
 * to the document's numbering when the paragraph is written.
 *
 * @see {@link Numbering}
 */
export type NumberingDefinition = {
    /** Array of level definitions for this numbering configuration. */
    readonly levels: readonly ILevelsOptions[];
    /** Unique reference name for this numbering configuration. */
    readonly reference: string;
    /**
     * The start overrides (`w:lvlOverride/w:startOverride`) of its instances, by instance: an instance listed here gets exactly
     * these, none for an empty list; one not listed starts over at every configured level.
     */
    readonly instances?: Readonly<Record<number, readonly { readonly level: number; readonly start: number }[]>>;
};

/**
 * Options for configuring numbering definitions.
 *
 * @property config - Array of numbering configurations
 *
 * @see {@link Numbering}
 */
export type INumberingOptions = {
    /** Array of numbering configurations, each with levels and a reference name. */
    readonly config: readonly NumberingDefinition[];
    /**
     * Whether the document gets the default bullet list (`default-bullet-numbering`, `w:numId` 1) that `bullet` paragraphs use;
     * true by default. A document without `bullet` paragraphs can leave it out, so it has no list nothing uses.
     */
    readonly defaultBullets?: boolean;
};

/**
 * Represents the numbering definitions in a WordprocessingML document.
 *
 * The numbering element contains abstract numbering definitions and their
 * concrete instances, which are referenced by paragraphs to create lists.
 * Each numbering configuration includes a default bullet list and any
 * custom numbering schemes defined by the user.
 *
 * Reference: http://officeopenxml.com/WPnumbering.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="numbering" type="CT_Numbering"/>
 *
 * <xsd:complexType name="CT_Numbering">
 *   <xsd:sequence>
 *     <xsd:element name="numPicBullet" type="CT_NumPicBullet" minOccurs="0" maxOccurs="unbounded"/>
 *     <xsd:element name="abstractNum" type="CT_AbstractNum" minOccurs="0" maxOccurs="unbounded"/>
 *     <xsd:element name="num" type="CT_Num" minOccurs="0" maxOccurs="unbounded"/>
 *     <xsd:element name="numIdMacAtCleanup" type="CT_DecimalNumber" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Create numbering with custom decimal list
 * const numbering = new Numbering({
 *   config: [
 *     {
 *       reference: "my-decimal-list",
 *       levels: [
 *         {
 *           level: 0,
 *           format: LevelFormat.DECIMAL,
 *           text: "%1.",
 *           alignment: AlignmentType.LEFT,
 *           start: 1,
 *           style: {
 *             paragraph: {
 *               indent: { left: 720, hanging: 360 },
 *             },
 *           },
 *         },
 *         {
 *           level: 1,
 *           format: LevelFormat.LOWER_LETTER,
 *           text: "%2)",
 *           alignment: AlignmentType.LEFT,
 *           style: {
 *             paragraph: {
 *               indent: { left: 1440, hanging: 360 },
 *             },
 *           },
 *         },
 *       ],
 *     },
 *   ],
 * });
 * ```
 */
export class Numbering extends XmlComponent {
    private readonly abstractNumberingMap = new Map<string, AbstractNumbering>();
    private readonly concreteNumberingMap = new Map<string, ConcreteNumbering>();
    private readonly referenceConfigMap = new Map<string, readonly ILevelsOptions[]>();
    private readonly referenceInstancesMap = new Map<string, NonNullable<NumberingDefinition["instances"]>>();
    private readonly abstractNumUniqueNumericId = abstractNumUniqueNumericIdGen();
    private readonly concreteNumUniqueNumericId = concreteNumUniqueNumericIdGen();

    /**
     * Creates a new numbering definition collection.
     *
     * Initializes the numbering with a default bullet list configuration and
     * any custom numbering configurations provided in the options.
     *
     * @param options - Configuration options for numbering definitions
     */
    public constructor(options: INumberingOptions) {
        super("w:numbering");
        this.root.push(
            new DocumentAttributes(
                ["wpc", "mc", "o", "r", "m", "v", "wp14", "wp", "w10", "w", "w14", "w15", "wpg", "wpi", "wne", "wps"],
                "w14 w15 wp14",
            ),
        );

        const abstractNumbering = new AbstractNumbering(this.abstractNumUniqueNumericId(), [
            {
                level: 0,
                format: LevelFormat.BULLET,
                text: "\u25CF",
                alignment: AlignmentType.LEFT,
                style: {
                    paragraph: {
                        indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.25) },
                    },
                },
            },
            {
                level: 1,
                format: LevelFormat.BULLET,
                text: "\u25CB",
                alignment: AlignmentType.LEFT,
                style: {
                    paragraph: {
                        indent: { left: convertInchesToTwip(1), hanging: convertInchesToTwip(0.25) },
                    },
                },
            },
            {
                level: 2,
                format: LevelFormat.BULLET,
                text: "\u25A0",
                alignment: AlignmentType.LEFT,
                style: {
                    paragraph: {
                        indent: { left: 2160, hanging: convertInchesToTwip(0.25) },
                    },
                },
            },
            {
                level: 3,
                format: LevelFormat.BULLET,
                text: "\u25CF",
                alignment: AlignmentType.LEFT,
                style: {
                    paragraph: {
                        indent: { left: 2880, hanging: convertInchesToTwip(0.25) },
                    },
                },
            },
            {
                level: 4,
                format: LevelFormat.BULLET,
                text: "\u25CB",
                alignment: AlignmentType.LEFT,
                style: {
                    paragraph: {
                        indent: { left: 3600, hanging: convertInchesToTwip(0.25) },
                    },
                },
            },
            {
                level: 5,
                format: LevelFormat.BULLET,
                text: "\u25A0",
                alignment: AlignmentType.LEFT,
                style: {
                    paragraph: {
                        indent: { left: 4320, hanging: convertInchesToTwip(0.25) },
                    },
                },
            },
            {
                level: 6,
                format: LevelFormat.BULLET,
                text: "\u25CF",
                alignment: AlignmentType.LEFT,
                style: {
                    paragraph: {
                        indent: { left: 5040, hanging: convertInchesToTwip(0.25) },
                    },
                },
            },
            {
                level: 7,
                format: LevelFormat.BULLET,
                text: "\u25CF",
                alignment: AlignmentType.LEFT,
                style: {
                    paragraph: {
                        indent: { left: 5760, hanging: convertInchesToTwip(0.25) },
                    },
                },
            },
            {
                level: 8,
                format: LevelFormat.BULLET,
                text: "\u25CF",
                alignment: AlignmentType.LEFT,
                style: {
                    paragraph: {
                        indent: { left: 6480, hanging: convertInchesToTwip(0.25) },
                    },
                },
            },
        ]);

        if (options.defaultBullets !== false) {
            this.addDefaultBullets(abstractNumbering);
        }

        for (const con of options.config) {
            // eslint-disable-next-line functional/immutable-data
            this.abstractNumberingMap.set(con.reference, new AbstractNumbering(this.abstractNumUniqueNumericId(), con.levels));
            // eslint-disable-next-line functional/immutable-data
            this.referenceConfigMap.set(con.reference, con.levels);
            if (con.instances) {
                // eslint-disable-next-line functional/immutable-data
                this.referenceInstancesMap.set(con.reference, con.instances);
            }
        }
    }

    /** Adds the default bullet list, `w:numId` 1, which `bullet` paragraphs refer to. */
    private addDefaultBullets(abstractNumbering: AbstractNumbering): void {
        // eslint-disable-next-line functional/immutable-data
        this.concreteNumberingMap.set(
            "default-bullet-numbering",
            new ConcreteNumbering({
                numId: 1,
                abstractNumId: abstractNumbering.id,
                reference: "default-bullet-numbering",
                instance: 0,
                overrideLevels: [
                    {
                        num: 0,
                        start: 1,
                    },
                ],
            }),
        );

        // eslint-disable-next-line functional/immutable-data
        this.abstractNumberingMap.set("default-bullet-numbering", abstractNumbering);
    }

    /**
     * Prepares the numbering definitions for XML serialization.
     *
     * Adds all abstract and concrete numbering definitions to the XML tree.
     *
     * @param context - The XML context
     * @returns The prepared XML object
     */
    public prepForXml(context: IContext): IXmlableObject | undefined {
        // Replace the definitions written before, after the attributes, so a document packed again doesn't repeat them
        this.root.splice(1, this.root.length - 1, ...this.abstractNumberingMap.values(), ...this.concreteNumberingMap.values());
        return super.prepForXml(context);
    }

    /**
     * Creates a concrete numbering instance from an abstract numbering definition.
     *
     * This method creates a new concrete numbering instance that references an
     * abstract numbering definition. It's used internally when paragraphs reference
     * numbering configurations.
     *
     * @param reference - The reference name of the abstract numbering definition
     * @param instance - The instance number for this concrete numbering
     * @param definition - The list's definition, for a list a paragraph gives its own definition of. It is added with
     * its reference, unless the numbering already has a list with that reference, such as one of the document's
     * `numbering` option, which is kept
     * @param restartedListId - For a list a paragraph starts again with `restart`, the id of that restart. The list is
     * then that restart's own, not the `instance`'s
     */
    public createConcreteNumberingInstance(
        reference: string,
        instance: number,
        definition?: NumberingDefinition,
        restartedListId?: number,
    ): void {
        if (definition && !this.abstractNumberingMap.has(definition.reference)) {
            // eslint-disable-next-line functional/immutable-data
            this.abstractNumberingMap.set(
                definition.reference,
                new AbstractNumbering(this.abstractNumUniqueNumericId(), definition.levels),
            );
            // eslint-disable-next-line functional/immutable-data
            this.referenceConfigMap.set(definition.reference, definition.levels);
            if (definition.instances) {
                // eslint-disable-next-line functional/immutable-data
                this.referenceInstancesMap.set(definition.reference, definition.instances);
            }
        }

        const abstractNumbering = this.abstractNumberingMap.get(reference);

        if (!abstractNumbering) {
            return;
        }

        const fullReference = listKey(reference, instance, restartedListId);

        if (this.concreteNumberingMap.has(fullReference)) {
            return;
        }

        const referenceConfigLevels = this.referenceConfigMap.get(reference) ?? [{ level: 0 }];

        // Every configured level starts over in a new instance, not only the first, so a list restarts at each of its levels.
        // Each starts at the number its level is written with, as `Level` writes it
        const concreteNumberingSettings = {
            numId: this.concreteNumUniqueNumericId(),
            abstractNumId: abstractNumbering.id,
            reference,
            instance,
            restartedListId,
            // an instance with start overrides of its own gets those (a restarted list keeps starting over)
            overrideLevels:
                (restartedListId === undefined ? this.referenceInstancesMap.get(reference)?.[instance] : undefined)?.map(
                    ({ level, start }) => ({ num: level, start: decimalNumber(start) }),
                ) ??
                referenceConfigLevels.map(({ level, start = 1 }) => ({
                    num: level,
                    start: decimalNumber(start),
                })),
        };

        // eslint-disable-next-line functional/immutable-data
        this.concreteNumberingMap.set(fullReference, new ConcreteNumbering(concreteNumberingSettings));
    }

    /**
     * Gets all concrete numbering instances.
     *
     * @returns An array of all concrete numbering instances
     */
    public get ConcreteNumbering(): readonly ConcreteNumbering[] {
        return Array.from(this.concreteNumberingMap.values());
    }

    /**
     * Gets all reference configurations.
     *
     * @returns An array of all numbering reference configurations
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public get ReferenceConfig(): readonly Record<string, any>[] {
        return Array.from(this.referenceConfigMap.values());
    }
}
