/**
 * Unzipper buffer-stream module
 * Original source: https://github.com/ZJONSSON/node-unzipper
 * License: MIT
 * Copyright (c) 2012 - 2013 Near Infinity Corporation
 * Commits in this fork are (c) Ziggy Jonsson (ziggy.jonsson.nyc@gmail.com)
 */

import { Transform } from "stream";
import type { TransformCallback, Readable } from "stream";

export function bufferStream(entry: Readable): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    const stream = new Transform({
      transform(d: Buffer, _encoding: BufferEncoding, cb: TransformCallback) {
        chunks.push(d);
        cb();
      }
    });

    stream.on("finish", () => {
      resolve(Buffer.concat(chunks));
    });
    stream.on("error", reject);

    entry.on("error", reject).pipe(stream);
  });
}
