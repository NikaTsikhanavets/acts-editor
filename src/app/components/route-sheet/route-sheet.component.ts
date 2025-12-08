import { Component, EventEmitter, inject, Input, OnInit, Output, signal, WritableSignal } from '@angular/core';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { Driver, Executor } from '@models';
import { ExcelParserService } from '@services';
import { ErrorInfo } from '../../interfaces/error-info.interface';
import { ErrorsComponent } from '@components/errors/errors.component';

@Component({
  selector: 'app-route-sheet',
  imports: [CommonModule, ReactiveFormsModule, ErrorsComponent],
  templateUrl: './route-sheet.component.html',
  styleUrl: './route-sheet.component.scss',
  standalone: true
})
export class RouteSheetComponent implements OnInit {
  @Input() uploadedDocument!: XLSX.WorkBook | null;
  @Output() goBack: EventEmitter<void> = new EventEmitter<void>();

  drivers: Driver[] = [];
  selectedDrivers: Set<Driver> = new Set();
  isTouchedDriversList: WritableSignal<boolean> = signal(false);
  managers: string[] = [];
  executors: Executor[] = [];
  routeSheetForm!: FormGroup;
  parsingErrors: ErrorInfo[] = [];


  private parserService: ExcelParserService = inject(ExcelParserService);
  private fb: FormBuilder = inject(FormBuilder);
  private http: HttpClient = inject(HttpClient);

  public ngOnInit(): void {
    this.initForm();

    if (!this.uploadedDocument) {
      return;
    }

    const {drivers, managers, clients, executors, errors} = this.parserService.parseDocument(this.uploadedDocument);

    if (errors?.length) {
      this.parsingErrors = errors || [];
      return;
    }

    this.drivers = drivers;
    this.managers = managers;
    this.executors = Object.values(executors);
  }

  private initForm(): void {
    this.routeSheetForm = this.fb.group({
      executor: [null, Validators.required],
      monthOption: ['next', Validators.required]
    });
  }

  public toggleDriver(driver: Driver, event: Event): void {
    const checked = (event.target as HTMLInputElement).checked;
    if (checked) {
      this.selectedDrivers.add(driver);
    } else {
      this.selectedDrivers.delete(driver);
    }

    this.isTouchedDriversList.set(true);
  }

  public isDriverSelected(driver: Driver): boolean {
    return this.selectedDrivers.has(driver);
  }

  public onSubmit(): void {
    if (this.routeSheetForm.invalid) {
      return;
    }

    if (this.selectedDrivers.size === 0) {
      return;
    }

    const selectedDriversArray: Driver[] = Array.from(this.selectedDrivers);
    const selectedExecutor: Executor = this.routeSheetForm.value.executor;
    const monthOption: 'current' | 'next' = this.routeSheetForm.value.monthOption;

    if (selectedDriversArray.length === 1) {
      this.generateSingleRouteSheet(selectedDriversArray[0], selectedExecutor, monthOption);
    } else {
      this.generateRouteSheetsZip(selectedDriversArray, selectedExecutor, monthOption);
    }
  }

  private async generateSingleRouteSheet(driver: Driver, executor: Executor, monthOption: 'current' | 'next'): Promise<void> {
    this.loadTemplate(async (templateData) => {
      const buffer = await this.createRouteSheetBuffer(templateData, driver, executor, monthOption);
      this.downloadFile(buffer, `маршрутный_лист_${driver.shortName}.xlsx`);
    });
  }

  private async generateRouteSheetsZip(drivers: Driver[], executor: Executor, monthOption: 'current' | 'next'): Promise<void> {
    this.loadTemplate(async (templateData) => {
      const zip = new JSZip();

      for (const driver of drivers) {
        const buffer = await this.createRouteSheetBuffer(templateData, driver, executor, monthOption);
        zip.file(`маршрутный_лист_${driver.shortName}.xlsx`, buffer);
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      this.downloadFile(zipBlob, `маршрутные_листы_${this.getMonthInfo(monthOption).monthName}.zip`);
    });
  }

  private loadTemplate(callback: (data: ArrayBuffer) => Promise<void>): void {
    const templatePath = '/assets/doc-templates/route-sheet.xlsx';

    this.http.get(templatePath, { responseType: 'arraybuffer' }).subscribe({
      next: async (data: ArrayBuffer) => {
        try {
          await callback(data);
        } catch (error) {
          console.error('Error processing template:', error);
        }
      },
      error: (error) => {
        console.error('Error loading template:', error);
      }
    });
  }

  private async createRouteSheetBuffer(
    templateData: ArrayBuffer,
    driver: Driver,
    executor: Executor,
    monthOption: 'current' | 'next'
  ): Promise<ExcelJS.Buffer> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateData);

    workbook.eachSheet((worksheet) => {
      this.replaceInWorksheet(worksheet, driver, executor, monthOption);
    });

    return await workbook.xlsx.writeBuffer();
  }

  private downloadFile(data: Blob | ExcelJS.Buffer, filename: string): void {
    const blob = data instanceof Blob ? data : new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.click();
    window.URL.revokeObjectURL(url);
  }

  private replaceInWorksheet(worksheet: ExcelJS.Worksheet, driver: Driver, executor: Executor, monthOption: 'current' | 'next'): void {
    const { monthName, lastDay, year } = this.getMonthInfo(monthOption);
    const executorInfo = this.getExecutorInfo(executor);

    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (cell.value && typeof cell.value === 'string') {
          if (cell.value.includes('{{DRIVER_NAME}}')) {
            cell.value = cell.value.replace(/\{\{DRIVER_NAME\}\}/g, driver.fullName);
          }
          if (cell.value.includes('{{MONTH}}')) {
            cell.value = cell.value.replace(/\{\{MONTH\}\}/g, monthName);
          }
          if (cell.value.includes('{{LAST_DAY}}')) {
            cell.value = cell.value.replace(/\{\{LAST_DAY\}\}/g, lastDay.toString());
          }
          if (cell.value.includes('{{YEAR}}')) {
            cell.value = cell.value.replace(/\{\{YEAR\}\}/g, year);
          }
          if (cell.value.includes('{{EXECUTOR}}')) {
            cell.value = cell.value.replace(/\{\{EXECUTOR\}\}/g, executorInfo);
          }
          if (cell.value.includes('{{DRIVER_ID}}')) {
            cell.value = cell.value.replace(/\{\{DRIVER_ID\}\}/g, driver.id);
          }
          if (cell.value.includes('{{CAR_MAKE}}')) {
            cell.value = cell.value.replace(/\{\{CAR_MAKE\}\}/g, driver.carMake);
          }
          if (cell.value.includes('{{CAR_NUMBER}}')) {
            cell.value = cell.value.replace(/\{\{CAR_NUMBER\}\}/g, driver.carNumber);
          }
        }
      });
    });
  }

  private getMonthInfo(monthOption: 'current' | 'next'): { monthName: string; lastDay: number; year: string } {
    const months = [
      'Января',
      'Февраля',
      'Марта',
      'Апреля',
      'Мая',
      'Июня',
      'Июля',
      'Августа',
      'Сентября',
      'Октября',
      'Ноября',
      'Декабря'
    ];

    const now = new Date();
    let monthIndex = now.getMonth();
    let year = now.getFullYear();

    if (monthOption === 'next') {
      monthIndex += 1;
      if (monthIndex > 11) {
        monthIndex = 0;
        year += 1;
      }
    }

    const lastDayDate = new Date(year, monthIndex + 1, 0);

    return {
      monthName: months[monthIndex],
      lastDay: lastDayDate.getDate(),
      year: year.toString()
    };
  }

  private getExecutorInfo(executor: Executor): string {
    return `${executor.actualName}, ${executor.address} ОГРН ${executor.ogrn} ИНН ${executor.inn} ${executor.cpp}`;
  }

  public return(): void {
    this.goBack.emit();
  }
}
