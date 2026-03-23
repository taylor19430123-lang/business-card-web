function installTypedArrayToHexPolyfill() {
  if (typeof Uint8Array === "undefined" || typeof Uint8Array.prototype.toHex === "function") {
    return;
  }

  Object.defineProperty(Uint8Array.prototype, "toHex", {
    value() {
      return Array.from(this, (value) => value.toString(16).padStart(2, "0")).join("");
    },
    writable: true,
    configurable: true
  });
}

installTypedArrayToHexPolyfill();

export * from "/vendor/pdfjs/pdf.core.mjs";
