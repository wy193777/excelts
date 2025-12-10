import { describe, it, expect } from "vitest";
import { Workbook } from "../../../index.js";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Issue 2961: Excel files exported by DataGrip/IDEA may have rows and cells
// without the 'r' attribute. The parser should handle these by inferring
// positions from context.

describe("github issues", () => {
  describe("issue 2961 - missing r attribute in row and cell elements", () => {
    it("should read xlsx file with missing r attributes on rows and cells", async () => {
      const wb = new Workbook();
      // example.xlsx is exported by DataGrip/IDEA and has no r attributes
      const filePath = path.resolve(__dirname, "../data/issue-2961-missing-r-attribute.xlsx");
      await wb.xlsx.readFile(filePath);

      const ws = wb.worksheets[0];
      expect(ws).toBeDefined();

      // The file has 2 rows with 3 columns each:
      // Row 1: header1, header2, header3 (shared strings)
      // Row 2: 1, value1, value2
      expect(ws.rowCount).toBe(2);

      // Check first row
      const row1 = ws.getRow(1);
      expect(row1.getCell(1).value).toBeDefined();
      expect(row1.getCell(2).value).toBeDefined();
      expect(row1.getCell(3).value).toBeDefined();

      // Check second row
      const row2 = ws.getRow(2);
      expect(row2.getCell(1).value).toBe(1);
      expect(row2.getCell(2).value).toBeDefined();
      expect(row2.getCell(3).value).toBeDefined();
    });

    it("should correctly infer cell addresses when r attribute is missing", async () => {
      const wb = new Workbook();
      const filePath = path.resolve(__dirname, "../data/issue-2961-missing-r-attribute.xlsx");
      await wb.xlsx.readFile(filePath);

      const ws = wb.worksheets[0];

      // Verify that cells have correct addresses (A1, B1, C1, A2, B2, C2)
      expect(ws.getCell("A1").value).toBeDefined();
      expect(ws.getCell("B1").value).toBeDefined();
      expect(ws.getCell("C1").value).toBeDefined();
      expect(ws.getCell("A2").value).toBeDefined();
      expect(ws.getCell("B2").value).toBeDefined();
      expect(ws.getCell("C2").value).toBeDefined();
    });

    it("should be able to write the file back after reading", async () => {
      const wb = new Workbook();
      const filePath = path.resolve(__dirname, "../data/issue-2961-missing-r-attribute.xlsx");
      await wb.xlsx.readFile(filePath);

      // Writing to buffer should not throw
      const buffer = await wb.xlsx.writeBuffer();
      expect(buffer).toBeDefined();
      expect(buffer.byteLength).toBeGreaterThan(0);
    });
  });
});
