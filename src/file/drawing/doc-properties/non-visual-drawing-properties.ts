/**
 * Non-visual drawing properties (`CT_NonVisualDrawingProps`) of the shapes, pictures and groups inside a drawing:
 * their id, name and alternative text, a link followed when they are clicked, and whether they are decorative.
 *
 * @module
 */
import { TargetModeType } from "@file/relationships/relationship/relationship";
import { BuilderElement, type IContext, type IXmlableObject, NextAttributeComponent, XmlComponent } from "@file/xml-components";
import { uniqueId } from "@util/convenience-functions";

import { createHyperlinkClick } from "./doc-properties-children";

/**
 * A link and accessibility settings shared by drawings and the shapes in them.
 */
export type DrawingLinkOptions = {
    /** A web address followed when the drawing is clicked (with Ctrl in Word) */
    readonly link?: string;
    /** Marks the drawing as decorative, so screen readers skip it */
    readonly decorative?: boolean;
};

export type NonVisualDrawingPropertiesOptions = DrawingLinkOptions & {
    readonly id: number;
    readonly name: string;
    readonly description?: string;
    readonly title?: string;
};

// The extension Word writes for "Mark as decorative"
const DECORATIVE_EXTENSION_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

/**
 * Creates the extension list that marks a drawing as decorative, as Word's "Mark as decorative" does.
 *
 * ```xml
 * <a:extLst>
 *   <a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">
 *     <adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/>
 *   </a:ext>
 * </a:extLst>
 * ```
 *
 * @param declareNamespace - Declares the DrawingML namespace, for elements outside `a:graphic` such as `wp:docPr`
 */
export const createDecorativeExtensionList = (declareNamespace: boolean): XmlComponent => createExtensionList(declareNamespace, [], true);

const createDecorativeExtension = (): XmlComponent =>
    new BuilderElement<{ readonly uri: string }>({
        name: "a:ext",
        attributes: { uri: { key: "uri", value: DECORATIVE_EXTENSION_URI } },
        children: [
            new BuilderElement<{ readonly namespace: string; readonly value: number }>({
                name: "adec:decorative",
                attributes: {
                    namespace: { key: "xmlns:adec", value: "http://schemas.microsoft.com/office/drawing/2017/decorative" },
                    value: { key: "val", value: 1 },
                },
            }),
        ],
    });

/**
 * An extension of an application's own in a drawing's extension list: an `a:ext` with its `uri` (a GUID that names the extension) and
 * one element in the application's namespace. Word keeps extensions it doesn't know when it saves a document.
 *
 * ```xml
 * <a:ext uri="{00000000-0000-0000-0000-000000000000}">
 *   <x:mark xmlns:x="urn:example:mark" id="42"/>
 * </a:ext>
 * ```
 */
export type DrawingExtension = {
    /** The extension's GUID, in braces */
    readonly uri: string;
    /** The element's qualified name, with its prefix, such as `x:mark` */
    readonly name: string;
    /** The namespace the element's prefix stands for */
    readonly namespace: string;
    /** The element's attributes, by name */
    readonly attributes?: Readonly<Record<string, string>>;
};

const createExtension = ({ uri, name, namespace, attributes = {} }: DrawingExtension): XmlComponent => {
    const prefix = name.includes(":") ? name.slice(0, name.indexOf(":")) : undefined;
    const declared =
        prefix === undefined
            ? { namespace: { key: "xmlns", value: namespace } }
            : { namespace: { key: `xmlns:${prefix}`, value: namespace } };
    const own = Object.fromEntries(Object.entries(attributes).map(([key, value]) => [key, { key, value }]));
    return new BuilderElement<{ readonly uri: string }>({
        name: "a:ext",
        attributes: { uri: { key: "uri", value: uri } },
        children: [new BuilderElement<Record<string, string>>({ name, attributes: { ...declared, ...own } })],
    });
};

/**
 * Creates a drawing's extension list (`a:extLst`) with extensions of an application's own, after the decorative mark when there is one.
 *
 * @param declareNamespace - Declares the DrawingML namespace, for elements outside `a:graphic` such as `wp:docPr`
 */
export const createExtensionList = (declareNamespace: boolean, extensions: readonly DrawingExtension[], decorative = false): XmlComponent =>
    new BuilderElement<{ readonly namespace?: string }>({
        name: "a:extLst",
        attributes: {
            namespace: {
                key: "xmlns:a",
                value: declareNamespace ? "http://schemas.openxmlformats.org/drawingml/2006/main" : undefined,
            },
        },
        children: [...(decorative ? [createDecorativeExtension()] : []), ...extensions.map(createExtension)],
    });

/**
 * A link from a drawing to a web address: the `a:hlinkClick` element, and the relationship that holds the address.
 *
 * The relationship is added to the part the drawing is written in (the document, a header or a footer) when it is written.
 */
export class DrawingLink {
    public readonly linkId = uniqueId();
    // The relationship lists the link has been added to, so writing a document twice doesn't add it twice
    private readonly addedTo = new WeakSet<object>();

    public constructor(private readonly link: string) {}

    public createClick(declareNamespace: boolean): XmlComponent {
        return createHyperlinkClick(this.linkId, declareNamespace);
    }

    public addRelationship(context: IContext): void {
        const relationships = context.viewWrapper.Relationships;
        if (this.addedTo.has(relationships)) {
            return;
        }
        this.addedTo.add(relationships);
        relationships.addRelationship(
            this.linkId,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            this.link,
            TargetModeType.EXTERNAL,
        );
    }
}

/**
 * The non-visual drawing properties of a shape, picture or group inside a drawing, such as `wps:cNvPr`.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_NonVisualDrawingProps">
 *   <xsd:sequence>
 *     <xsd:element name="hlinkClick" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="hlinkHover" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="id" type="ST_DrawingElementId" use="required"/>
 *   <xsd:attribute name="name" type="xsd:string" use="required"/>
 *   <xsd:attribute name="descr" type="xsd:string" use="optional" default=""/>
 *   <xsd:attribute name="hidden" type="xsd:boolean" use="optional" default="false"/>
 *   <xsd:attribute name="title" type="xsd:string" use="optional" default=""/>
 * </xsd:complexType>
 * ```
 */
export class NonVisualDrawingProperties extends XmlComponent {
    private readonly link?: DrawingLink;

    public constructor(name: string, { id, name: drawingName, description, title, link, decorative }: NonVisualDrawingPropertiesOptions) {
        super(name);

        this.link = link === undefined ? undefined : new DrawingLink(link);
        this.root.push(
            new NextAttributeComponent<{
                readonly id: number;
                readonly name: string;
                readonly description?: string;
                readonly title?: string;
            }>({
                id: { key: "id", value: id },
                name: { key: "name", value: drawingName },
                description: { key: "descr", value: description },
                title: { key: "title", value: title },
            }),
        );
        if (this.link) {
            this.root.push(this.link.createClick(false));
        }
        if (decorative) {
            this.root.push(createDecorativeExtensionList(false));
        }
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        this.link?.addRelationship(context);
        return super.prepForXml(context);
    }
}
