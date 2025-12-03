import { describe, it, expect } from "vitest";
import { bufferStream } from "../../../utils/unzip/buffer-stream.js";
import { Readable } from "stream";

describe("buffer-stream", () => {
  describe("bufferStream", () => {
    it("should concatenate all chunks into a single buffer", async () => {
      const input = Readable.from([Buffer.from("hello"), Buffer.from(" "), Buffer.from("world")]);

      const result = await bufferStream(input);

      expect(result.toString()).toBe("hello world");
    });

    it("should handle empty stream", async () => {
      const input = Readable.from([]);

      const result = await bufferStream(input);

      expect(result.length).toBe(0);
    });

    it("should handle single chunk", async () => {
      const input = Readable.from([Buffer.from("single chunk")]);

      const result = await bufferStream(input);

      expect(result.toString()).toBe("single chunk");
    });

    it("should handle binary data", async () => {
      const binary = Buffer.from([0x00, 0x01, 0x02, 0xff, 0xfe, 0xfd]);
      const input = Readable.from([binary]);

      const result = await bufferStream(input);

      expect(result).toEqual(binary);
    });

    it("should reject on stream error", async () => {
      const input = new Readable({
        read() {
          this.destroy(new Error("Stream error"));
        }
      });

      await expect(bufferStream(input)).rejects.toThrow("Stream error");
    });
  });
});
