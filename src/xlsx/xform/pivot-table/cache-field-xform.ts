import { BaseXform } from "../base-xform.js";
import { xmlDecode } from "../../../utils/utils.js";

/**
 * Parsed cache field model containing name and shared items (if any)
 */
interface CacheFieldModel {
  name: string;
  sharedItems: string[] | null;
  // Numeric field metadata
  containsNumber?: boolean;
  containsInteger?: boolean;
  minValue?: number;
  maxValue?: number;
}

/**
 * Xform for parsing individual <cacheField> elements within a pivot cache definition.
 *
 * Example XML:
 * ```xml
 * <cacheField name="Category" numFmtId="0">
 *   <sharedItems count="3">
 *     <s v="A" />
 *     <s v="B" />
 *     <s v="C" />
 *   </sharedItems>
 * </cacheField>
 *
 * <cacheField name="Value" numFmtId="0">
 *   <sharedItems containsSemiMixedTypes="0" containsString="0"
 *                containsNumber="1" containsInteger="1" minValue="5" maxValue="45" />
 * </cacheField>
 * ```
 */
class CacheFieldXform extends BaseXform {
  declare public model: CacheFieldModel | null;
  private inSharedItems: boolean;

  constructor() {
    super();
    this.model = null;
    this.inSharedItems = false;
  }

  get tag(): string {
    return "cacheField";
  }

  reset(): void {
    this.model = null;
    this.inSharedItems = false;
  }

  parseOpen(node: any): boolean {
    const { name, attributes } = node;

    switch (name) {
      case "cacheField":
        // Initialize the model with field name
        this.model = {
          name: xmlDecode(attributes.name || ""),
          sharedItems: null
        };
        break;

      case "sharedItems":
        this.inSharedItems = true;
        // Check if this is a numeric field (no string items)
        if (attributes.containsNumber === "1" || attributes.containsInteger === "1") {
          if (this.model) {
            this.model.containsNumber = attributes.containsNumber === "1";
            this.model.containsInteger = attributes.containsInteger === "1";
            if (attributes.minValue !== undefined) {
              this.model.minValue = parseFloat(attributes.minValue);
            }
            if (attributes.maxValue !== undefined) {
              this.model.maxValue = parseFloat(attributes.maxValue);
            }
            // Numeric fields have sharedItems = null
            this.model.sharedItems = null;
          }
        } else {
          // String field - initialize sharedItems array if count > 0
          if (this.model) {
            const count = parseInt(attributes.count || "0", 10);
            if (count > 0) {
              this.model.sharedItems = [];
            }
          }
        }
        break;

      case "s":
        // String value in sharedItems
        if (this.inSharedItems && this.model?.sharedItems !== null) {
          // Decode XML entities in the value
          const value = xmlDecode(attributes.v || "");
          this.model.sharedItems!.push(value);
        }
        break;

      case "n":
        // Numeric value in sharedItems (less common, but possible)
        if (this.inSharedItems && this.model?.sharedItems !== null) {
          const value = parseFloat(attributes.v || "0");
          this.model.sharedItems!.push(value as any);
        }
        break;

      case "b":
        // Boolean value in sharedItems
        if (this.inSharedItems && this.model?.sharedItems !== null) {
          const value = attributes.v === "1";
          this.model.sharedItems!.push(value as any);
        }
        break;

      case "e":
        // Error value in sharedItems
        if (this.inSharedItems && this.model?.sharedItems !== null) {
          const value = `#${attributes.v || "ERROR!"}`;
          this.model.sharedItems!.push(value);
        }
        break;

      case "m":
        // Missing/null value in sharedItems
        if (this.inSharedItems && this.model?.sharedItems !== null) {
          this.model.sharedItems!.push(null as any);
        }
        break;

      case "d":
        // Date value in sharedItems
        if (this.inSharedItems && this.model?.sharedItems !== null) {
          const value = new Date(attributes.v || "");
          this.model.sharedItems!.push(value as any);
        }
        break;
    }

    return true;
  }

  parseText(_text: string): void {
    // No text content in cacheField elements
  }

  parseClose(name: string): boolean {
    switch (name) {
      case "cacheField":
        // End of this cacheField element
        return false;

      case "sharedItems":
        this.inSharedItems = false;
        break;
    }

    return true;
  }
}

export { CacheFieldXform, type CacheFieldModel };
