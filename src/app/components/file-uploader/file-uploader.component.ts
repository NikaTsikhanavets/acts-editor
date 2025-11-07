import { ChangeDetectionStrategy, Component, EventEmitter, Output } from '@angular/core';
import * as XLSX from 'xlsx';
import { DocumentService } from '@services';
import { LoaderComponent } from '@components/loader/loader.component';
import { DropDirective } from '../../drop.directive';

@Component({
  selector: 'app-file-uploader',
  templateUrl: './file-uploader.component.html',
  styleUrls: ['./file-uploader.component.scss'],
  imports: [
    LoaderComponent,
    DropDirective
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  standalone: true
})
export class FileUploaderComponent {
  public excelDocument!: XLSX.WorkBook;
  public inProgress: boolean = false;
  @Output() uploadedFile: EventEmitter<XLSX.WorkBook> = new EventEmitter<XLSX.WorkBook>();

  constructor(private readonly documentService: DocumentService) {
  }

  public uploadFile(file: any): void {
    this.importFile(file?.target?.files[0]);
  }

  public importFile(file: File | undefined): void {
    if (!file) {
      return;
    }

    this.inProgress = true;
    this.documentService.importExcelFile(file).then((excelDocument: XLSX.WorkBook) => {
      this.updateDocument(excelDocument);
    });
  }

  public loadFileByUrl(): void {
    this.inProgress = true;

    this.documentService.getFile().subscribe((excelDocument) => {
      this.updateDocument(excelDocument);
    });
  }

  private updateDocument(excelDocument: XLSX.WorkBook): void {
    this.excelDocument = excelDocument;
    this.inProgress = false;
    this.uploadedFile.emit(this.excelDocument);
  }
}
