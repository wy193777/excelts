import { describe, it, expect } from "vitest";
import { NoopStream } from "../../../utils/unzip/noop-stream.js";
import { Readable } from "stream";

describe("noop-stream", () => {
  describe("NoopStream", () => {
    it("should consume all data without emitting anything", async () => {
      const noop = new NoopStream();
      const chunks: Buffer[] = [];

      noop.on("data", chunk => {
        chunks.push(chunk);
      });

      const input = Readable.from([Buffer.from("hello"), Buffer.from("world")]);

      await new Promise<void>((resolve, reject) => {
        input
          .pipe(noop)
          .on("finish", () => {
            resolve();
          })
          .on("error", reject);
      });

      expect(chunks.length).toBe(0);
    });

    it("should emit finish event", async () => {
      const noop = new NoopStream();
      let finished = false;

      noop.on("finish", () => {
        finished = true;
      });

      const input = Readable.from([Buffer.from("test")]);

      await new Promise<void>(resolve => {
        input.pipe(noop).on("finish", () => {
          resolve();
        });
      });

      expect(finished).toBe(true);
    });
  });
});
