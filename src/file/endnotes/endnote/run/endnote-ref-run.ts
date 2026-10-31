import { EndnoteReference, type IRunOptions, Run } from "@file/paragraph";

export class EndnoteRefRun extends Run {
    public constructor(options: IRunOptions = {}) {
        super({
            style: "EndnoteReference",
            ...options,
        });

        this.root.push(new EndnoteReference());
    }
}
