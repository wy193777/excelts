/**
 * Unzipper pull-stream module
 * Original source: https://github.com/ZJONSSON/node-unzipper
 * License: MIT
 * Copyright (c) 2012 - 2013 Near Infinity Corporation
 * Commits in this fork are (c) Ziggy Jonsson (ziggy.jonsson.nyc@gmail.com)
 */

import { Duplex, PassThrough, Transform } from "stream";
import type { TransformCallback } from "stream";

const STR_FUNCTION = "function";

export class PullStream extends Duplex {
  buffer: Buffer;
  cb?: () => void;
  finished: boolean;
  match?: number;
  __emittedError?: Error;

  constructor() {
    super({ decodeStrings: false, objectMode: true });
    this.buffer = Buffer.from("");
    this.finished = false;

    this.on("finish", () => {
      this.finished = true;
      this.emit("chunk", false);
    });
  }

  _write(chunk: Buffer, _encoding: BufferEncoding, cb: () => void): void {
    this.buffer = Buffer.concat([this.buffer, chunk]);
    this.cb = cb;
    this.emit("chunk");
  }

  _read(): void {}

  /**
   * The `eof` parameter is interpreted as `file_length` if the type is number
   * otherwise (i.e. buffer) it is interpreted as a pattern signaling end of stream
   */
  stream(eof: number | Buffer, includeEof?: boolean): PassThrough {
    const p = new PassThrough();
    let done = false;

    const cb = (): void => {
      if (typeof this.cb === STR_FUNCTION) {
        const callback = this.cb;
        this.cb = undefined;
        callback();
      }
    };

    const pull = (): void => {
      let packet: Buffer | undefined;
      if (this.buffer && this.buffer.length) {
        if (typeof eof === "number") {
          packet = this.buffer.slice(0, eof);
          this.buffer = this.buffer.slice(eof);
          eof -= packet.length;
          done = done || !eof;
        } else {
          let match = this.buffer.indexOf(eof);
          if (match !== -1) {
            // store signature match byte offset to allow us to reference
            // this for zip64 offset
            this.match = match;
            if (includeEof) {
              match = match + eof.length;
            }
            packet = this.buffer.slice(0, match);
            this.buffer = this.buffer.slice(match);
            done = true;
          } else {
            const len = this.buffer.length - eof.length;
            if (len <= 0) {
              cb();
            } else {
              packet = this.buffer.slice(0, len);
              this.buffer = this.buffer.slice(len);
            }
          }
        }
        if (packet) {
          p.write(packet, () => {
            if (
              this.buffer.length === 0 ||
              (typeof eof !== "number" && eof.length && this.buffer.length <= eof.length)
            ) {
              cb();
            }
          });
        }
      }

      if (!done) {
        if (this.finished) {
          this.removeListener("chunk", pull);
          this.emit("error", new Error("FILE_ENDED"));
          return;
        }
      } else {
        this.removeListener("chunk", pull);
        p.end();
      }
    };

    this.on("chunk", pull);
    pull();
    return p;
  }

  pull(eof: number | Buffer, includeEof?: boolean): Promise<Buffer> {
    if (eof === 0) {
      return Promise.resolve(Buffer.from(""));
    }

    // If we already have the required data in buffer
    // we can resolve the request immediately
    if (typeof eof === "number" && this.buffer.length > eof) {
      const data = this.buffer.slice(0, eof);
      this.buffer = this.buffer.slice(eof);
      return Promise.resolve(data);
    }

    // Otherwise we stream until we have it
    let buffer = Buffer.from("");

    const concatStream = new Transform({
      transform(d: Buffer, _e: BufferEncoding, cb: TransformCallback) {
        buffer = Buffer.concat([buffer, d]);
        cb();
      }
    });

    let rejectHandler: (reason?: any) => void;
    let pullStreamRejectHandler: (e: Error) => void;

    return new Promise<Buffer>((resolve, reject) => {
      rejectHandler = reject;
      pullStreamRejectHandler = (e: Error) => {
        this.__emittedError = e;
        reject(e);
      };
      if (this.finished) {
        return reject(new Error("FILE_ENDED"));
      }
      this.once("error", pullStreamRejectHandler); // reject any errors from pullstream itself
      this.stream(eof, includeEof)
        .on("error", reject)
        .pipe(concatStream)
        .on("finish", () => {
          resolve(buffer);
        })
        .on("error", reject);
    }).finally(() => {
      this.removeListener("error", rejectHandler);
      this.removeListener("error", pullStreamRejectHandler);
    });
  }
}
