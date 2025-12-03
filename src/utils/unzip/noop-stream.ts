/**
 * Unzipper noop-stream module
 * Original source: https://github.com/ZJONSSON/node-unzipper
 * License: MIT
 * Copyright (c) 2012 - 2013 Near Infinity Corporation
 * Commits in this fork are (c) Ziggy Jonsson (ziggy.jonsson.nyc@gmail.com)
 */

import { Transform } from "stream";
import type { TransformCallback } from "stream";

export class NoopStream extends Transform {
  constructor() {
    super();
  }

  _transform(_chunk: any, _encoding: BufferEncoding, cb: TransformCallback): void {
    cb();
  }
}
