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

function installCollectionInsertPolyfills() {
  const install = (Prototype) => {
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

  install(globalThis.Map?.prototype);
  install(globalThis.WeakMap?.prototype);
}

installTypedArrayToHexPolyfill();
installCollectionInsertPolyfills();

export * from "/vendor/pdfjs/pdf.core.mjs";
