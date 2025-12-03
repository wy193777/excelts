const textDecoder = typeof TextDecoder === "undefined" ? null : new TextDecoder("utf-8");

function bufferToString(chunk: Buffer<ArrayBuffer> | string): string {
  if (typeof chunk === "string") {
    return chunk;
  }
  if (textDecoder) {
    return textDecoder.decode(chunk);
  }
  return String.fromCharCode(...new Uint8Array(chunk));
}

export { bufferToString };
