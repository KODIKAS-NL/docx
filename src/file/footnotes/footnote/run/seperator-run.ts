/**
 * Footnote separator run module for WordprocessingML documents.
 *
 * This module provides the run containing the separator element.
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php
 *
 * @module
 */
import { type IRunOptions, Run } from "@file/paragraph";

import { Seperator } from "./seperator";

/**
 * Represents a footnote separator run in WordprocessingML.
 *
 * SeperatorRun creates a run containing the separator element that
 * defines the horizontal line between body text and footnotes.
 *
 * @internal
 */
export class SeperatorRun extends Run {
    public constructor(options: IRunOptions = {}) {
        super(options);

        this.root.push(new Seperator());
    }
}
