import { describe, it, expect } from "vitest";
import { PullStream } from "../../../utils/unzip/pull-stream.js";
import { Readable } from "stream";

describe("pull-stream", () => {
  describe("PullStream", () => {
    it("should pull exact number of bytes", async () => {
      const pull = new PullStream();
      const input = Readable.from([Buffer.from("hello world")]);

      input.pipe(pull);

      const result = await pull.pull(5);
      expect(result.toString()).toBe("hello");
    });

    it("should pull remaining bytes after first pull", async () => {
      const pull = new PullStream();
      const input = Readable.from([Buffer.from("hello world")]);

      input.pipe(pull);

      const first = await pull.pull(6);
      expect(first.toString()).toBe("hello ");

      const second = await pull.pull(5);
      expect(second.toString()).toBe("world");
    });

    it("should handle chunked input", async () => {
      const pull = new PullStream();
      const input = Readable.from([Buffer.from("hel"), Buffer.from("lo "), Buffer.from("world")]);

      input.pipe(pull);

      const result = await pull.pull(11);
      expect(result.toString()).toBe("hello world");
    });

    it("should return empty buffer for pull(0)", async () => {
      const pull = new PullStream();
      const input = Readable.from([Buffer.from("hello")]);

      input.pipe(pull);

      const result = await pull.pull(0);
      expect(result.length).toBe(0);
    });

    it("should handle immediate data availability", async () => {
      const pull = new PullStream();
      pull.buffer = Buffer.from("immediate data");

      const result = await pull.pull(9);
      expect(result.toString()).toBe("immediate");
    });

    it("should emit error on FILE_ENDED when stream finishes before pull completes", async () => {
      const pull = new PullStream();
      const input = Readable.from([Buffer.from("short")]);

      input.pipe(pull);

      // Wait for data to be available then try to pull more than available
      await new Promise(resolve => setTimeout(resolve, 10));

      await expect(pull.pull(100)).rejects.toThrow("FILE_ENDED");
    });

    it("should stream data until eof number of bytes", async () => {
      const pull = new PullStream();
      const input = Readable.from([Buffer.from("hello world")]);

      input.pipe(pull);

      const chunks: Buffer[] = [];
      const stream = pull.stream(5);

      await new Promise<void>(resolve => {
        stream.on("data", chunk => chunks.push(chunk));
        stream.on("end", () => resolve());
      });

      expect(Buffer.concat(chunks).toString()).toBe("hello");
    });

    it("should stream data until eof pattern found", async () => {
      const pull = new PullStream();
      const input = Readable.from([Buffer.from("hello|world")]);

      input.pipe(pull);

      const chunks: Buffer[] = [];
      const eof = Buffer.from("|");
      const stream = pull.stream(eof);

      await new Promise<void>(resolve => {
        stream.on("data", chunk => chunks.push(chunk));
        stream.on("end", () => resolve());
      });

      expect(Buffer.concat(chunks).toString()).toBe("hello");
    });
  });
});
