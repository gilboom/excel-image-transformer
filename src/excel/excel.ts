import * as XLSX from "xlsx";
import { ExcelImageResolver } from "./image";

export class Excel {
  workbook: XLSX.WorkBook;
  imageResolver: ExcelImageResolver;

  load(buffer: Uint8Array | Blob | ArrayBuffer) {
    this.workbook = XLSX.read(buffer, { type: "buffer", bookFiles: true });
    console.debug("workbook %o", this.workbook);
    this.imageResolver = new ExcelImageResolver(this.workbook);
  }

  async transformImagesToStr(
    cb: (file: File, cell: { row: number, col: number }) => Promise<string>
  ) {
    const locationsMap = await this.imageResolver.resolveImageLocations();
    for (let sheetName in locationsMap) {
      const locations = locationsMap[sheetName];
      const data = this.getDataBySheetName(sheetName);
      for (const location of locations) {
        const str = await cb(location.file, location.from);
        data[location.from.row] = data[location.from.row] || [];
        data[location.from.row][location.from.col] = str;
      }
      this.setData(sheetName, data);
    }
  }

  getData() {
    this._assertHasLoaded();
    return this.workbook.SheetNames.map(name => {
      return {
        sheetName: name,
        data: this.getDataBySheetName(name),
      }
    });
  }

  getDataBySheetName(sheetName: string) {
    const sheet = this.workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json<Array<any>>(sheet, { header: 1 });
  }

  setData(sheetName: string, data: Array<Array<any>>) {
    this._assertSheetExist(sheetName);
    const sheet = this.workbook.Sheets[sheetName];
    XLSX.utils.sheet_add_aoa(sheet, data);
  }

  export() {
    const workbook = XLSX.utils.book_new();
    this.workbook.SheetNames.forEach(name => {
      const sheet = this.workbook.Sheets[name];
      XLSX.utils.book_append_sheet(workbook, sheet, name);
    });
    XLSX.writeFile(workbook, "Sheet.xlsx");
  }

  private _assertHasLoaded() {
    if (!this.workbook) throw new Error("还未加载 excel 文件");
  }

  private _assertSheetExist(sheetName: string) {
    if (!this.workbook.Sheets[sheetName]) throw new Error(`没有名为 ${sheetName} 的工作薄`);
  }

}

