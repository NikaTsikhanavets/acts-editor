import { ChangeDetectionStrategy, Component, EventEmitter, Output } from '@angular/core';
import * as XLSX from 'xlsx';
import { DocumentService } from '@services';
import { LoaderComponent } from '@components/loader/loader.component';
import { DropDirective } from '../../drop.directive';
import jsPDF from 'jspdf';

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
  @Output() uploadedPdf: EventEmitter<File> = new EventEmitter<File>();

  constructor(private readonly documentService: DocumentService) {
  }

  public uploadFile(file: any): void {
    this.importFile(file?.target?.files[0]);
  }

  public async uploadPdfOrImageFile(file: any): Promise<void> {
    const uploadedFile = file?.target?.files[0];
    if (!uploadedFile) {
      return;
    }

    // Check if it's a PDF
    if (uploadedFile.type === 'application/pdf') {
      this.uploadedPdf.emit(uploadedFile);
      return;
    }

    // Check if it's an image
    if (uploadedFile.type.startsWith('image/')) {
      this.inProgress = true;
      try {
        const pdfFile = await this.convertImageToPdf(uploadedFile);
        this.inProgress = false;
        this.uploadedPdf.emit(pdfFile);
      } catch (error) {
        console.error('Error converting image:', error);
        alert('Ошибка при конвертации изображения');
        this.inProgress = false;
      }
      return;
    }

    // Unsupported file type
    alert('Пожалуйста, загрузите PDF файл или изображение');
  }

  private async convertImageToPdf(imageFile: File): Promise<File> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        const img = new Image();
        
        img.onload = () => {
          try {
            // Create PDF with image dimensions
            const imgWidth = img.width;
            const imgHeight = img.height;
            
            // Calculate PDF page size (A4 or custom based on image)
            const maxWidth = 210; // A4 width in mm
            const maxHeight = 297; // A4 height in mm
            
            let pdfWidth = maxWidth;
            let pdfHeight = (imgHeight * maxWidth) / imgWidth;
            
            // If image is too tall, use full height and adjust width
            if (pdfHeight > maxHeight) {
              pdfHeight = maxHeight;
              pdfWidth = (imgWidth * maxHeight) / imgHeight;
            }
            
            // Create PDF
            const pdf = new jsPDF({
              orientation: pdfWidth > pdfHeight ? 'landscape' : 'portrait',
              unit: 'mm',
              format: [pdfWidth, pdfHeight]
            });
            
            // Add image to PDF
            pdf.addImage(
              e.target?.result as string,
              imageFile.type.includes('png') ? 'PNG' : 'JPEG',
              0,
              0,
              pdfWidth,
              pdfHeight
            );
            
            // Convert to File
            const pdfBlob = pdf.output('blob');
            const pdfFile = new File(
              [pdfBlob],
              imageFile.name.replace(/\.(jpg|jpeg|png|gif|bmp|webp)$/i, '.pdf'),
              { type: 'application/pdf' }
            );
            
            resolve(pdfFile);
          } catch (error) {
            reject(error);
          }
        };
        
        img.onerror = () => {
          reject(new Error('Не удалось загрузить изображение'));
        };
        
        img.src = e.target?.result as string;
      };
      
      reader.onerror = () => {
        reject(new Error('Не удалось прочитать файл'));
      };
      
      reader.readAsDataURL(imageFile);
    });
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
