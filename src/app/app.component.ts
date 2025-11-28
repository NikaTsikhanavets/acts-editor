import { ChangeDetectionStrategy, Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { FileUploaderComponent } from '@components/file-uploader/file-uploader.component';
import { DriversListComponent } from '@components/drivers-list/drivers-list.component';
import { PdfViewerComponent } from '@components/pdf-viewer/pdf-viewer.component';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
  imports: [
    FileUploaderComponent,
    DriversListComponent,
    PdfViewerComponent
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  standalone: true
})
export class AppComponent {
  public excelDocument?: XLSX.WorkBook;
  public pdfFile?: File;
}
