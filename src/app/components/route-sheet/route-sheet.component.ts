import { Component, EventEmitter, inject, Input, OnInit, Output } from '@angular/core';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs';
import { Driver, Executor } from '@models';
import { ExcelParserService } from '@services';

@Component({
  selector: 'app-route-sheet',
  imports: [CommonModule, ReactiveFormsModule],
  templateUrl: './route-sheet.component.html',
  styleUrl: './route-sheet.component.scss',
  standalone: true
})
export class RouteSheetComponent implements OnInit {
  @Input() uploadedDocument!: XLSX.WorkBook | null;
  @Output() goBack: EventEmitter<void> = new EventEmitter<void>();

  drivers: Driver[] = [];
  managers: string[] = [];
  executors: Executor[] = [];
  routeSheetForm!: FormGroup;

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
      // this.parsingErrors = errors || [];
      return;
    }

    this.drivers = drivers;
    this.managers = managers;
    this.executors = Object.values(executors);
  }

  private initForm(): void {
    this.routeSheetForm = this.fb.group({
      driver: [null, Validators.required],
      executor: [null, Validators.required]
    });
  }

  public onSubmit(): void {
    if (this.routeSheetForm.invalid) {
      return;
    }

    const selectedDriver: Driver = this.routeSheetForm.value.driver;
    const selectedExecutor: Executor = this.routeSheetForm.value.executor;

    this.generateRouteSheet(selectedDriver, selectedExecutor);
  }

  private async generateRouteSheet(driver: Driver, executor: Executor): Promise<void> {
    const templatePath = '/assets/doc-templates/route-sheet.xlsx';

    this.http.get(templatePath, { responseType: 'arraybuffer' }).subscribe({
      next: async (data: ArrayBuffer) => {
        try {
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(data);

          // Replace placeholders in all sheets
          workbook.eachSheet((worksheet) => {
            this.replaceInWorksheet(worksheet, driver, executor);
          });

          // Generate and download the file with all formatting preserved
          const buffer = await workbook.xlsx.writeBuffer();
          const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
          const url = window.URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url;
          link.download = `маршрутный_лист_${driver.shortName}.xlsx`;
          link.click();
          window.URL.revokeObjectURL(url);
        } catch (error) {
          console.error('Error processing template:', error);
        }
      },
      error: (error) => {
        console.error('Error loading template:', error);
      }
    });
  }

  private replaceInWorksheet(worksheet: ExcelJS.Worksheet, driver: Driver, executor: Executor): void {
    const currentMonth = this.getCurrentMonthInRussian();
    const lastDay = this.getLastDayOfMonth();
    const currentYear = new Date().getFullYear().toString();
    const executorInfo = this.getExecutorInfo(executor);

    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (cell.value && typeof cell.value === 'string') {
          if (cell.value.includes('{{DRIVER_NAME}}')) {
            cell.value = cell.value.replace(/\{\{DRIVER_NAME\}\}/g, driver.fullName);
          }
          if (cell.value.includes('{{MONTH}}')) {
            cell.value = cell.value.replace(/\{\{MONTH\}\}/g, currentMonth);
          }
          if (cell.value.includes('{{LAST_DAY}}')) {
            cell.value = cell.value.replace(/\{\{LAST_DAY\}\}/g, lastDay.toString());
          }
          if (cell.value.includes('{{YEAR}}')) {
            cell.value = cell.value.replace(/\{\{YEAR\}\}/g, currentYear);
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

  private getCurrentMonthInRussian(): string {
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

    const currentMonthIndex = new Date().getMonth();
    return months[currentMonthIndex];
  }

  private getLastDayOfMonth(): number {
    const now = new Date();
    const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    return lastDay.getDate();
  }

  private getExecutorInfo(executor: Executor): string {
    return `${executor.actualName}, ${executor.address} ОГРН ${executor.ogrn} ИНН ${executor.inn} ${executor.cpp}`;
  }

  public return(): void {
    this.goBack.emit();
  }
}
