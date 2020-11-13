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

  /**
   * 解析 excel 所有工作薄上的所有图片的位置和文件
   */
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

  /**
   * 解析所有 worksheet 上的图片
   */
  private async _resolveAllSheetImages(): Promise<IImageMetadatasMap> {
    const metadatasMap: IImageMetadatasMap = {};

    for (const sheetName of this.workbook.SheetNames) {
      const sheet = this.sheetMap.get(sheetName);
      if (!sheet) throw new Error("sheet not found");
      metadatasMap[sheetName] = await this._resolveSheetImages(sheet.id);
    }
    return metadatasMap;
  }

  /**
   * 解析 worksheet 对应的 drawing xml 和 drawing.rels xml
   * @param sheetId sheetId
   */
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

  /**
   * 解析 worksheet 上的图片
   * @param sheetId sheetId
   */
  private async _resolveSheetImages(sheetId: string): Promise<Array<IImageMetadata>> {
    const drawingObj = await this._resolveDrawing(sheetId);
    if (!drawingObj) return [];

    const { drawingXmlObj, drawingRelXmlObj } = drawingObj;
    console.debug("drawingXmlObj", drawingXmlObj);
    console.debug("drawingRelXmlObj", drawingRelXmlObj);

    const twoCellAnchors = drawingXmlObj["xdr:wsDr"]["xdr:twoCellAnchor"] as Array<any> || [];
    const oneCellAnchors = drawingXmlObj["xdr:wsDr"]["xdr:oneCellAnchor"] as Array<any> || [];
    const imageAnchors: Array<any> = oneCellAnchors.concat(twoCellAnchors);
    const imageAnchorsMap = new Map<RelationId, Array<IImageAnchor>>();
    imageAnchors.forEach(anchor => {
      // relationship id
      const id: RelationId = anchor["xdr:pic"]["xdr:blipFill"]["a:blip"]["r:embed"];
      const from = {
        row: Number(anchor["xdr:from"]["xdr:row"]),
        col: Number(anchor["xdr:from"]["xdr:col"])
      };
      const isOneCellAnchors = !anchor["xdr:to"];
      const to = {
        row: Number(anchor[`xdr:${isOneCellAnchors ? "from" : "to"}`]["xdr:row"]),
        col: Number(anchor[`xdr:${isOneCellAnchors ? "from" : "to"}`]["xdr:col"]),
      };
      // 同一张图片可能被多个单元格引用
      let entity = imageAnchorsMap.get(id);
      if (!entity) {
        entity = [];
      }
      entity.push({ from, to });
      imageAnchorsMap.set(id, entity);
    })
    console.debug("imageAnchorsMap", imageAnchorsMap);

    // 根据 relationship id 找到实际的图片文件
    const relationships: Array<any> = drawingRelXmlObj.Relationships.Relationship;
    const imageMetadatas: Array<IImageMetadata> = relationships.flatMap(r => {
      const id: RelationId = r.Id;
      const target: string = r.Target;
      const filename = target.substr(target.lastIndexOf("/") + 1);
      const path = `xl/media/${filename}`;
      const file = this._parseMetaImageToFile(path, filename);
      const imageAnchors = imageAnchorsMap.get(id);
      if (!imageAnchors) throw new Error("image anchor not found");
      return imageAnchors.map(anchor => {
        return { file, ...anchor }
      });
    });
    console.debug("imageMetadatas", imageMetadatas);
    return imageMetadatas;
  }

  /**
   * 从 workbook 对象的 files 属性获取文件 buffer
   * files 包含了 excel 文件解压后的各个文件的 buffer
   * @param path 文件路径
   */
  private _getDataFromWbFile(path: string) {
    const file = this.workbook["files"][path];
    if (!file) return null;

    // 有些文件的 _data 属性是 buffer , 有的 _data 属性是一个包含 getContent 方法的对象
    // 通过 getContent 方法可以获取文件的 buffer
    const _data: Uint8Array = file._data.getContent ?
      file._data.getContent() :
      file._data;

    return _data;
  }

  /**
   * 把 excel 文件解压后的 xml 文件转成 js 对象
   * @param path 文件路径
   */
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

  /**
   * 把 xml 字符串转成 js 对象
   * @param xml xml 字符串
   */
  private _xml2Obj<T>(xml: string): T {
    const attributesKey = "_attr";
    const textKey = "_text";
    const obj = xml2js(xml, {
      textKey,
      compact: true,
      attributesKey,
      alwaysChildren: false,
      alwaysArray: ["xdr:twoCellAnchor", "Relationship", "sheet"],
    }) as T;
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