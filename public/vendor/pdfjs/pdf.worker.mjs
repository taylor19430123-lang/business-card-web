if (typeof Uint8Array !== "undefined" && typeof Uint8Array.prototype.toHex !== "function") {
  Object.defineProperty(Uint8Array.prototype, "toHex", {
    value() {
      return Array.from(this, (value) => value.toString(16).padStart(2, "0")).join("");
    },
    writable: true,
    configurable: true
  });
}

const installCollectionInsertPolyfills = (Prototype) => {
  if (!Prototype) {
    return;
  }

  if (typeof Prototype.getOrInsertComputed !== "function") {
    Object.defineProperty(Prototype, "getOrInsertComputed", {
      value(key, compute) {
        if (this.has(key)) {
          return this.get(key);
        }

        const value = compute(key);
        this.set(key, value);
        return value;
      },
      writable: true,
      configurable: true
    });
  }

  if (typeof Prototype.getOrInsert !== "function") {
    Object.defineProperty(Prototype, "getOrInsert", {
      value(key, defaultValue) {
        if (this.has(key)) {
          return this.get(key);
        }

        this.set(key, defaultValue);
        return defaultValue;
      },
      writable: true,
      configurable: true
    });
  }
};

installCollectionInsertPolyfills(globalThis.Map?.prototype);
installCollectionInsertPolyfills(globalThis.WeakMap?.prototype);

await import("/vendor/pdfjs/pdf.worker.core.mjs");
