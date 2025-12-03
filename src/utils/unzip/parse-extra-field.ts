/**
 * Unzipper parse-extra-field module
 * Original source: https://github.com/ZJONSSON/node-unzipper
 * License: MIT
 * Copyright (c) 2012 - 2013 Near Infinity Corporation
 * Commits in this fork are (c) Ziggy Jonsson (ziggy.jonsson.nyc@gmail.com)
 */

import { parse } from "./parse-buffer.js";

export interface ZipVars {
  uncompressedSize: number;
  compressedSize: number;
  offsetToLocalFileHeader?: number;
}

export interface ExtraField {
  uncompressedSize?: number;
  compressedSize?: number;
  offsetToLocalFileHeader?: number;
}

export function parseExtraField(extraField: Buffer, vars: ZipVars): ExtraField {
  let extra: ExtraField | undefined;

  // Find the ZIP64 header, if present.
  while (!extra && extraField && extraField.length) {
    const candidateExtra = parse(extraField, [
      ["signature", 2],
      ["partSize", 2]
    ]);

    if (candidateExtra.signature === 0x0001) {
      // parse buffer based on data in ZIP64 central directory; order is important!
      const fieldsToExpect: [string, number][] = [];
      if (vars.uncompressedSize === 0xffffffff) {
        fieldsToExpect.push(["uncompressedSize", 8]);
      }
      if (vars.compressedSize === 0xffffffff) {
        fieldsToExpect.push(["compressedSize", 8]);
      }
      if (vars.offsetToLocalFileHeader === 0xffffffff) {
        fieldsToExpect.push(["offsetToLocalFileHeader", 8]);
      }

      // slice off the 4 bytes for signature and partSize
      extra = parse(extraField.slice(4), fieldsToExpect) as ExtraField;
    } else {
      // Advance the buffer to the next part.
      // The total size of this part is the 4 byte header + partsize.
      extraField = extraField.slice((candidateExtra.partSize || 0) + 4);
    }
  }

  extra = extra || {};

  if (vars.compressedSize === 0xffffffff) {
    vars.compressedSize = extra.compressedSize!;
  }

  if (vars.uncompressedSize === 0xffffffff) {
    vars.uncompressedSize = extra.uncompressedSize!;
  }

  if (vars.offsetToLocalFileHeader === 0xffffffff) {
    vars.offsetToLocalFileHeader = extra.offsetToLocalFileHeader;
  }

  return extra;
}
