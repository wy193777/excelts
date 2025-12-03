import { describe, it, expect } from "vitest";
import {
  decode_col,
  encode_col,
  decode_row,
  encode_row,
  decode_cell,
  encode_cell,
  decode_range,
  encode_range,
  json_to_sheet,
  sheet_add_json,
  sheet_to_json,
  sheet_to_csv,
  aoa_to_sheet,
  sheet_add_aoa,
  sheet_to_aoa,
  book_new,
  book_append_sheet,
  utils
} from "../../../utils/extra-utils.js";

describe("xlsx-compat utils", () => {
  // ===========================================================================
  // Cell Address Encoding/Decoding
  // ===========================================================================

  describe("decode_col", () => {
    it("should decode single letter columns", () => {
      expect(decode_col("A")).toBe(0);
      expect(decode_col("B")).toBe(1);
      expect(decode_col("Z")).toBe(25);
    });

    it("should decode double letter columns", () => {
      expect(decode_col("AA")).toBe(26);
      expect(decode_col("AB")).toBe(27);
      expect(decode_col("AZ")).toBe(51);
      expect(decode_col("BA")).toBe(52);
    });

    it("should handle lowercase letters", () => {
      expect(decode_col("a")).toBe(0);
      expect(decode_col("aa")).toBe(26);
    });
  });

  describe("encode_col", () => {
    it("should encode single letter columns", () => {
      expect(encode_col(0)).toBe("A");
      expect(encode_col(1)).toBe("B");
      expect(encode_col(25)).toBe("Z");
    });

    it("should encode double letter columns", () => {
      expect(encode_col(26)).toBe("AA");
      expect(encode_col(27)).toBe("AB");
      expect(encode_col(51)).toBe("AZ");
      expect(encode_col(52)).toBe("BA");
    });
  });

  describe("decode_row", () => {
    it("should decode row strings to 0-indexed numbers", () => {
      expect(decode_row("1")).toBe(0);
      expect(decode_row("10")).toBe(9);
      expect(decode_row("100")).toBe(99);
    });
  });

  describe("encode_row", () => {
    it("should encode 0-indexed numbers to row strings", () => {
      expect(encode_row(0)).toBe("1");
      expect(encode_row(9)).toBe("10");
      expect(encode_row(99)).toBe("100");
    });
  });

  describe("decode_cell", () => {
    it("should decode cell references to CellAddress", () => {
      expect(decode_cell("A1")).toEqual({ c: 0, r: 0 });
      expect(decode_cell("B2")).toEqual({ c: 1, r: 1 });
      expect(decode_cell("AA10")).toEqual({ c: 26, r: 9 });
    });

    it("should handle lowercase references", () => {
      expect(decode_cell("a1")).toEqual({ c: 0, r: 0 });
      expect(decode_cell("b2")).toEqual({ c: 1, r: 1 });
    });
  });

  describe("encode_cell", () => {
    it("should encode CellAddress to cell references", () => {
      expect(encode_cell({ c: 0, r: 0 })).toBe("A1");
      expect(encode_cell({ c: 1, r: 1 })).toBe("B2");
      expect(encode_cell({ c: 26, r: 9 })).toBe("AA10");
    });
  });

  describe("decode_cell and encode_cell roundtrip", () => {
    it("should roundtrip correctly", () => {
      const addresses = ["A1", "B2", "Z100", "AA1", "XFD1048576"];
      for (const addr of addresses) {
        expect(encode_cell(decode_cell(addr))).toBe(addr);
      }
    });
  });

  describe("decode_range", () => {
    it("should decode range strings", () => {
      expect(decode_range("A1:B2")).toEqual({
        s: { c: 0, r: 0 },
        e: { c: 1, r: 1 }
      });
    });

    it("should decode single cell as range", () => {
      expect(decode_range("A1")).toEqual({
        s: { c: 0, r: 0 },
        e: { c: 0, r: 0 }
      });
    });
  });

  describe("encode_range", () => {
    it("should encode Range object", () => {
      expect(encode_range({ s: { c: 0, r: 0 }, e: { c: 1, r: 1 } })).toBe("A1:B2");
    });

    it("should encode two CellAddress objects", () => {
      expect(encode_range({ c: 0, r: 0 }, { c: 1, r: 1 })).toBe("A1:B2");
    });

    it("should return single cell for same start and end", () => {
      expect(encode_range({ c: 0, r: 0 }, { c: 0, r: 0 })).toBe("A1");
    });
  });

  // ===========================================================================
  // Workbook/Worksheet Functions
  // ===========================================================================

  describe("book_new", () => {
    it("should create empty workbook", () => {
      const wb = book_new();
      expect(wb.worksheets).toHaveLength(0);
    });
  });

  describe("book_append_sheet", () => {
    it("should append existing worksheet to workbook", () => {
      const wb = book_new();
      const ws = json_to_sheet([{ name: "Alice", age: 30 }]);
      book_append_sheet(wb, ws, "Sheet1");
      expect(wb.worksheets).toHaveLength(1);
      expect(wb.worksheets[0].name).toBe("Sheet1");
      expect(wb.worksheets[0].getCell("A1").value).toBe("name");
    });

    it("should auto-generate name if not provided", () => {
      const wb = book_new();
      const ws = json_to_sheet([{ a: 1 }]);
      book_append_sheet(wb, ws);
      expect(wb.worksheets[0].name).toBeTruthy();
    });
  });

  // ===========================================================================
  // JSON/Sheet Conversion
  // ===========================================================================

  describe("json_to_sheet", () => {
    it("should convert JSON array to worksheet with headers", () => {
      const data = [
        { name: "Alice", age: 30 },
        { name: "Bob", age: 25 }
      ];
      const ws = json_to_sheet(data);

      expect(ws.getCell("A1").value).toBe("name");
      expect(ws.getCell("B1").value).toBe("age");
      expect(ws.getCell("A2").value).toBe("Alice");
      expect(ws.getCell("B2").value).toBe(30);
      expect(ws.getCell("A3").value).toBe("Bob");
      expect(ws.getCell("B3").value).toBe(25);
    });

    it("should respect header option for ordering", () => {
      const data = [{ name: "Alice", age: 30, city: "NYC" }];
      const ws = json_to_sheet(data, { header: ["age", "name"] });

      // header specifies order, but all keys are included
      expect(ws.getCell("A1").value).toBe("age");
      expect(ws.getCell("B1").value).toBe("name");
      expect(ws.getCell("C1").value).toBe("city");
      expect(ws.getCell("A2").value).toBe(30);
      expect(ws.getCell("B2").value).toBe("Alice");
      expect(ws.getCell("C2").value).toBe("NYC");
    });

    it("should skip header when skipHeader is true", () => {
      const data = [{ name: "Alice", age: 30 }];
      const ws = json_to_sheet(data, { skipHeader: true });

      expect(ws.getCell("A1").value).toBe("Alice");
      expect(ws.getCell("B1").value).toBe(30);
    });
  });

  describe("sheet_add_json", () => {
    it("should add JSON data to existing worksheet", () => {
      const ws = aoa_to_sheet([["Header1", "Header2"]]);
      sheet_add_json(ws, [{ a: 1, b: 2 }], { origin: "A2", skipHeader: true });

      expect(ws.getCell("A1").value).toBe("Header1");
      expect(ws.getCell("A2").value).toBe(1);
      expect(ws.getCell("B2").value).toBe(2);
    });

    it("should append to bottom with origin: -1", () => {
      const ws = aoa_to_sheet([
        ["a", "b"],
        [1, 2]
      ]);
      sheet_add_json(ws, [{ c: 3, d: 4 }], { origin: -1 });

      expect(ws.getCell("A3").value).toBe("c");
      expect(ws.getCell("B3").value).toBe("d");
      expect(ws.getCell("A4").value).toBe(3);
      expect(ws.getCell("B4").value).toBe(4);
    });
  });

  describe("sheet_to_json", () => {
    it("should convert worksheet to JSON array (default: first row as header)", () => {
      const ws = aoa_to_sheet([
        ["name", "age"],
        ["Alice", 30],
        ["Bob", 25]
      ]);

      const result = sheet_to_json(ws);

      expect(result).toHaveLength(2);
      expect(result[0]).toMatchObject({ name: "Alice", age: 30 });
      expect(result[1]).toMatchObject({ name: "Bob", age: 25 });
    });

    it("should return array of arrays with header: 1", () => {
      const ws = aoa_to_sheet([
        ["name", "age"],
        ["Alice", 30],
        ["Bob", 25]
      ]);

      const result = sheet_to_json(ws, { header: 1 });

      expect(result).toHaveLength(3);
      expect(result[0]).toEqual(["name", "age"]);
      expect(result[1]).toEqual(["Alice", 30]);
      expect(result[2]).toEqual(["Bob", 25]);
    });

    it("should use column letters as keys with header: 'A'", () => {
      const ws = aoa_to_sheet([
        ["name", "age"],
        ["Alice", 30]
      ]);

      const result = sheet_to_json(ws, { header: "A" });

      expect(result).toHaveLength(2);
      expect(result[0]).toMatchObject({ A: "name", B: "age" });
      expect(result[1]).toMatchObject({ A: "Alice", B: 30 });
    });

    it("should use custom keys with header: string[]", () => {
      const ws = aoa_to_sheet([
        ["Alice", 30],
        ["Bob", 25]
      ]);

      const result = sheet_to_json(ws, { header: ["person", "years"] });

      expect(result).toHaveLength(2);
      expect(result[0]).toMatchObject({ person: "Alice", years: 30 });
      expect(result[1]).toMatchObject({ person: "Bob", years: 25 });
    });

    it("should handle empty cells with defval", () => {
      const ws = aoa_to_sheet([
        ["col1", "col2"],
        ["value", null]
      ]);

      const result = sheet_to_json(ws, { defval: "" });

      expect(result[0]).toMatchObject({ col1: "value", col2: "" });
    });

    it("should skip blank rows by default for objects", () => {
      const ws = aoa_to_sheet([["name"], ["Alice"], [null], ["Bob"]]);

      const result = sheet_to_json(ws);

      expect(result).toHaveLength(2);
      expect(result[0]).toMatchObject({ name: "Alice" });
      expect(result[1]).toMatchObject({ name: "Bob" });
    });

    it("should include blank rows with blankrows: true for objects", () => {
      const ws = aoa_to_sheet([["name"], ["Alice"], [null], ["Bob"]]);

      const result = sheet_to_json(ws, { blankrows: true });

      expect(result).toHaveLength(3);
    });

    it("should include blank rows by default with header: 1", () => {
      const ws = aoa_to_sheet([["name"], ["Alice"], [null], ["Bob"]]);

      const result = sheet_to_json(ws, { header: 1 });

      expect(result).toHaveLength(4);
    });

    it("should disambiguate duplicate headers", () => {
      const ws = aoa_to_sheet([
        ["name", "name", "name"],
        ["Alice", "Bob", "Charlie"]
      ]);

      const result = sheet_to_json(ws);

      expect(result[0]).toHaveProperty("name", "Alice");
      expect(result[0]).toHaveProperty("name_1", "Bob");
      expect(result[0]).toHaveProperty("name_2", "Charlie");
    });
  });

  // ===========================================================================
  // AOA (Array of Arrays) Functions
  // ===========================================================================

  describe("aoa_to_sheet", () => {
    it("should convert array of arrays to worksheet", () => {
      const data = [
        ["Name", "Age"],
        ["Alice", 30],
        ["Bob", 25]
      ];
      const ws = aoa_to_sheet(data);

      expect(ws.getCell("A1").value).toBe("Name");
      expect(ws.getCell("B1").value).toBe("Age");
      expect(ws.getCell("A2").value).toBe("Alice");
      expect(ws.getCell("B2").value).toBe(30);
    });

    it("should handle origin option", () => {
      const data = [["a", "b"]];
      const ws = aoa_to_sheet(data, { origin: "C3" });

      expect(ws.getCell("C3").value).toBe("a");
      expect(ws.getCell("D3").value).toBe("b");
    });

    it("should handle different data types", () => {
      const date = new Date("2024-01-01");
      const data = [["string", 123, true, date, null]];
      const ws = aoa_to_sheet(data);

      expect(ws.getCell("A1").value).toBe("string");
      expect(ws.getCell("B1").value).toBe(123);
      expect(ws.getCell("C1").value).toBe(true);
      expect(ws.getCell("D1").value).toEqual(date);
    });
  });

  describe("sheet_add_aoa", () => {
    it("should add array of arrays to existing worksheet", () => {
      const ws = aoa_to_sheet([["Header"]]);
      sheet_add_aoa(ws, [["Data"]], { origin: "A2" });

      expect(ws.getCell("A1").value).toBe("Header");
      expect(ws.getCell("A2").value).toBe("Data");
    });

    it("should append with origin: -1", () => {
      const ws = aoa_to_sheet([["Row1"], ["Row2"]]);
      sheet_add_aoa(ws, [["Row3"]], { origin: -1 });

      expect(ws.getCell("A3").value).toBe("Row3");
    });
  });

  describe("sheet_to_aoa", () => {
    it("should convert worksheet to array of arrays", () => {
      const ws = aoa_to_sheet([
        ["Name", "Age"],
        ["Alice", 30]
      ]);

      const result = sheet_to_aoa(ws);

      expect(result[0]).toEqual(["Name", "Age"]);
      expect(result[1]).toEqual(["Alice", 30]);
    });
  });

  // ===========================================================================
  // CSV Functions
  // ===========================================================================

  describe("sheet_to_csv", () => {
    it("should convert worksheet to CSV string", () => {
      const ws = aoa_to_sheet([
        ["name", "age"],
        ["Alice", 30]
      ]);

      const csv = sheet_to_csv(ws);

      expect(csv).toBe("name,age\nAlice,30");
    });

    it("should use custom field separator", () => {
      const ws = aoa_to_sheet([["a", "b"]]);

      const csv = sheet_to_csv(ws, { FS: ";" });

      expect(csv).toBe("a;b");
    });

    it("should use custom record separator", () => {
      const ws = aoa_to_sheet([["a"], ["b"]]);

      const csv = sheet_to_csv(ws, { RS: "\r\n" });

      expect(csv).toBe("a\r\nb");
    });

    it("should quote values containing separator", () => {
      const ws = aoa_to_sheet([["hello,world"]]);

      const csv = sheet_to_csv(ws);

      expect(csv).toBe('"hello,world"');
    });

    it("should escape quotes", () => {
      const ws = aoa_to_sheet([['say "hello"']]);

      const csv = sheet_to_csv(ws);

      expect(csv).toBe('"say ""hello"""');
    });
  });

  // ===========================================================================
  // Utils Export
  // ===========================================================================

  describe("utils", () => {
    it("should export all functions", () => {
      expect(utils.decode_col).toBe(decode_col);
      expect(utils.encode_col).toBe(encode_col);
      expect(utils.decode_row).toBe(decode_row);
      expect(utils.encode_row).toBe(encode_row);
      expect(utils.decode_cell).toBe(decode_cell);
      expect(utils.encode_cell).toBe(encode_cell);
      expect(utils.decode_range).toBe(decode_range);
      expect(utils.encode_range).toBe(encode_range);
      expect(utils.json_to_sheet).toBe(json_to_sheet);
      expect(utils.sheet_add_json).toBe(sheet_add_json);
      expect(utils.sheet_to_json).toBe(sheet_to_json);
      expect(utils.sheet_to_csv).toBe(sheet_to_csv);
      expect(utils.aoa_to_sheet).toBe(aoa_to_sheet);
      expect(utils.sheet_add_aoa).toBe(sheet_add_aoa);
      expect(utils.sheet_to_aoa).toBe(sheet_to_aoa);
      expect(utils.book_new).toBe(book_new);
      expect(utils.book_append_sheet).toBe(book_append_sheet);
    });
  });
});
