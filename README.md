# excel-image-transformer

这是一个可以将 Excel 文件中的图片转换为字符串，并替换到图片放置的单元格中

## Example

在 Angular 中的例子

```typescript
import { Component } from '@angular/core';
import { Excel } from "excel-image-transformer";
import { UploadService } from "./upload.service"

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  constructor(
    private uploadService: UploadService
  ) { }

  // 监听文件变更
  onFileChange(e: any) {
    const target: DataTransfer = e.target;
    if (!target.files[0]) return;

    const fileReader = new FileReader();
    fileReader.onload = (e: any) => {
      const excel = new Excel();
      excel.load(e.target.result);
      excel.transformImagesToStr(async (file) => {
        // 通过 oss 上传图片然后得到图片的 url，并返回
        // 图片所在的单元格的内容就会被替换为 url
        const url = await this.uploadService.upload(file);
        return url;
      });

      const data = excel.getData();
      console.log("data", data);
    };
    fileReader.readAsArrayBuffer(target.files[0]);
  }
}

```