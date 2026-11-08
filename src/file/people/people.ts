/**
 * The people part (word/people.xml): who the authors of the comments and tracked changes are.
 *
 * Each person is named by the author name the comments and changes give, and may carry presence information: the
 * provider that knows the person (`None`, `AD`, `Windows Live`, ...) and the person's id there. Word uses it to tell
 * people with the same name apart.
 *
 * Reference: [MS-DOCX] 2.6.4 people, 2.6.5 person, 2.6.6 presenceInfo
 *
 * @module
 */
import { XmlAttributeComponent, XmlComponent } from "@file/xml-components";

/**
 * A person of the people part.
 *
 * @property author - The author name the comments and tracked changes give
 * @property providerId - The provider that knows the person, such as `None` or `AD`
 * @property userId - The person's id at that provider
 */
export type IPersonOptions = {
    readonly author: string;
    readonly providerId?: string;
    readonly userId?: string;
};

class PeopleAttributes extends XmlAttributeComponent<{
    readonly "xmlns:mc"?: string;
    readonly "xmlns:w15"?: string;
    readonly "mc:Ignorable"?: string;
}> {
    protected readonly xmlKeys = { "xmlns:mc": "xmlns:mc", "xmlns:w15": "xmlns:w15", "mc:Ignorable": "mc:Ignorable" };
}

class PersonAttributes extends XmlAttributeComponent<{ readonly author: string }> {
    protected readonly xmlKeys = { author: "w15:author" };
}

class PresenceInfoAttributes extends XmlAttributeComponent<{ readonly providerId: string; readonly userId: string }> {
    protected readonly xmlKeys = { providerId: "w15:providerId", userId: "w15:userId" };
}

class PresenceInfo extends XmlComponent {
    public constructor(providerId: string, userId: string) {
        super("w15:presenceInfo");
        this.root.push(new PresenceInfoAttributes({ providerId, userId }));
    }
}

class Person extends XmlComponent {
    public constructor({ author, providerId, userId }: IPersonOptions) {
        super("w15:person");
        this.root.push(new PersonAttributes({ author }));
        // both are required in a presence info
        if (providerId !== undefined && userId !== undefined) {
            this.root.push(new PresenceInfo(providerId, userId));
        }
    }
}

/**
 * The people part.
 *
 * ```typescript
 * new Document({
 *     people: [{ author: "Ann Author", providerId: "None", userId: "ann" }],
 *     sections: [],
 * });
 * ```
 */
export class People extends XmlComponent {
    public constructor(people: readonly IPersonOptions[]) {
        super("w15:people");
        this.root.push(
            new PeopleAttributes({
                "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
                "mc:Ignorable": "w15",
            }),
        );
        for (const person of people) {
            this.root.push(new Person(person));
        }
    }
}
