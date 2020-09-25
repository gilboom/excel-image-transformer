import * as XLSX from "xlsx";
import { xml2js } from "xml-js"

export interface IImageMetadatasMap {
  [sheetName: string]: Array<IImageMetadata>;
}

export interface IImageAnchor {
  from: { col: number, row: number }
  to: { col: number, row: number }
}

export interface IImageMetadata extends IImageAnchor {
  file: File
}

export type RelationId = string;
export type SheetName = string;
export type SheetId = string;
export type Sheet = {
  id: SheetId;
  name: SheetName;
}

export class ExcelImageResolver {
  sheetMap: Map<SheetName, Sheet>;
  constructor(
    private workbook: XLSX.WorkBook
  ) { }

  async resolveImageLocations(): Promise<IImageMetadatasMap> {
    await this._resolveSheetMap();
    const map = await this._resolveAllSheetImages();
    console.debug("metadatasMap", map);
    return map;
  }

  private async _resolveSheetMap() {
    const wbXmlObj = await this._parseMetaXmlToObject("xl/workbook.xml");
    console.debug("wbXmlObj %o", wbXmlObj);

    this.sheetMap = new Map<SheetName, Sheet>(
      wbXmlObj.workbook.sheets.sheet.map(sheet => {
        return [sheet.name, { name: sheet.name, id: sheet.sheetId }];
      })
    );
    console.debug("sheetMap %o", this.sheetMap);
  }

  private async _resolveAllSheetImages(): Promise<IImageMetadatasMap> {
    const metadatasMap: IImageMetadatasMap = {};

    for (const sheetName of this.workbook.SheetNames) {
      const sheet = this.sheetMap.get(sheetName);
      if (!sheet) throw new Error("sheet not found");
      metadatasMap[sheetName] = await this._resolveSheetImages(sheet.id);
    }
    return metadatasMap;
  }

  private async _resolveDrawing(sheetId: string) {
    const sheetXmlObj = await this._parseMetaXmlToObject(`xl/worksheets/sheet${sheetId}.xml`);
    if (!sheetXmlObj.worksheet.drawing) return null;
    const drawingRelId = sheetXmlObj.worksheet.drawing['r:id'];

    const sheetRelsXmlObj = await this._parseMetaXmlToObject(`xl/worksheets/_rels/sheet${sheetId}.xml.rels`);
    if (!sheetRelsXmlObj) return null;
    const relationships: Array<any> = sheetRelsXmlObj.Relationships.Relationship;
    const relationshipOfDrawing = relationships.find(r => {
      return r.Id = drawingRelId;
    });
    if (!relationshipOfDrawing) return null;

    const drawingTarget = (relationshipOfDrawing.Target as string)
    const drawingXmlFilename = drawingTarget.substr(drawingTarget.lastIndexOf("/") + 1);
    const drawingXmlFilePath = drawingTarget.replace("..", "xl");
    const drawingXmlObj = await this._parseMetaXmlToObject(drawingXmlFilePath);

    const drawingRelXmlFilename = drawingXmlFilename + ".rels";
    const drawingRelXmlFilePath = `xl/drawings/_rels/${drawingRelXmlFilename}`;
    const drawingRelXmlObj = await this._parseMetaXmlToObject(drawingRelXmlFilePath);
    return { drawingXmlObj, drawingRelXmlObj };
  }

  private async _resolveSheetImages(sheetId: string): Promise<Array<IImageMetadata>> {
    const drawingObj = await this._resolveDrawing(sheetId);
    if (!drawingObj) return [];

    const { drawingXmlObj, drawingRelXmlObj } = drawingObj;
    console.debug("drawingXmlObj", drawingXmlObj);
    console.debug("drawingRelXmlObj", drawingRelXmlObj);

    const imageAnchors: Array<any> = drawingXmlObj["xdr:wsDr"]["xdr:twoCellAnchor"];
    const imageAnchorsMap = new Map<RelationId, IImageAnchor>(
      imageAnchors.map(anchor => {
        const id = anchor["xdr:pic"]["xdr:blipFill"]["a:blip"]["r:embed"];
        const from = {
          row: Number(anchor["xdr:from"]["xdr:row"]),
          col: Number(anchor["xdr:from"]["xdr:col"])
        };
        const to = {
          row: Number(anchor["xdr:to"]["xdr:row"]),
          col: Number(anchor["xdr:to"]["xdr:col"]),
        };
        return [id, { from, to }];
      })
    );
    console.debug("imageAnchorsMap", imageAnchorsMap);

    const relationships: Array<any> = drawingRelXmlObj.Relationships.Relationship;

    const imageMetadatas: Array<IImageMetadata> = relationships.map(r => {
      const id: RelationId = r.Id;
      const target: string = r.Target;
      const filename = target.substr(target.lastIndexOf("/") + 1);
      const path = `xl/media/${filename}`;
      const file = this._parseMetaImageToFile(path, filename);
      const imageAnchor = imageAnchorsMap.get(id);
      if (!imageAnchor) throw new Error("image anchor not found");
      const { from, to } = imageAnchor;
      return { file, from, to };
    })
    console.debug("imageMetadatas", imageMetadatas);
    return imageMetadatas;
  }

  private _getDataFromWbFile(path: string) {
    const file = this.workbook["files"][path];
    if (!file) return null;

    const _data: Uint8Array = file._data.getContent ? file._data.getContent() : file._data;
    return _data;
  }

  private async _parseMetaXmlToObject<T = any>(path: string) {
    const _data: Uint8Array | null = this._getDataFromWbFile(path);
    if (!_data) return null;

    const fileReader = new FileReader();
    return new Promise<T>((resolve, reject) => {
      fileReader.onload = (e) => {
        const xml = e.target.result as string;
        console.debug(`FileLoader 加载文件事件: `, e)
        try {
          resolve(this._xml2Obj<T>(xml));
        } catch (err) {
          reject(new Error(`解析 xml 失败 文件名: ${path} 文件内容: ${xml}`))
        }
      };
      fileReader.onerror = (e) => {
        reject(e.target.error);
      }
      fileReader.readAsText(new Blob([_data]));
    });
  }


  private _parseMetaImageToFile(path: string, filename: string) {
    const _data: Uint8Array | null = this._getDataFromWbFile(path);
    if (!_data) return null;

    const file = new File([_data], filename);
    return file;
  }

  private _xml2Obj<T>(xml: string): T {
    const attributesKey = "_attr";
    const textKey = "_text";
    const obj = xml2js(xml, { compact: true, alwaysArray: ["xdr:twoCellAnchor", "Relationship", "sheet"], alwaysChildren: false, attributesKey, textKey }) as T;
    const newObj: any = {};

    function _mergeAttrAndText(obj: any, newObj: any, newObjKey?: string, newObjParent?: any) {
      Object.keys(obj)
        .forEach(k => {
          const attrs = obj[k]
          if (k === attributesKey) {
            Object.assign(newObj, attrs);
            delete newObj[k];
          }
          else if (k === textKey) {
            if (newObjParent) newObjParent[newObjKey] = attrs;
          }
          else {
            Object.assign(newObj, { [k]: attrs });
            if (typeof attrs === "object") {
              _mergeAttrAndText(attrs, newObj[k], k, newObj);
            }
          }
        })
    }
    _mergeAttrAndText(obj, newObj);
    return newObj;
  }
}