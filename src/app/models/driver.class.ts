import { Column, Page } from '@enums';
import { ErrorInfo } from '../interfaces/error-info.interface';

export class Driver {
  public firstName: string;
  public lastName: string;
  public id: string;
  public carMake: string;
  public carNumber: string;
  public patronymic: string;
  public car: string;
  public manager: string;
  public error: ErrorInfo | null;
  private readonly requiredFields: Column[] = [Column.FIRST_NAME, Column.LAST_NAME, Column.PATRONYMIC, Column.CAR, Column.MANAGER];

  constructor(driver: any) {
    this.error = this.getErrors(driver);
    this.firstName = driver[Column.FIRST_NAME]?.trim();
    this.lastName = driver[Column.LAST_NAME]?.trim();
    this.patronymic = driver[Column.PATRONYMIC]?.trim();
    this.id = driver[Column.ID]?.trim();
    this.car = driver[Column.CAR]?.trim();
    this.carMake = driver[Column.CAR_MAKE]?.trim();
    this.carNumber = driver[Column.CAR_NUMBER]?.trim();
    this.manager = driver[Column.MANAGER];
  }

  public get fullName(): string {
    return `${this.lastName} ${this.firstName} ${this.patronymic}`;
  }

  public get shortName(): string {
    const [nameFirstLetter]: string = this.firstName;
    const [patronymicFirstLetter]: string = this.patronymic;
    return `${this.lastName} ${nameFirstLetter}.${patronymicFirstLetter}.`;
  }

  public getErrors(docDriver: any): ErrorInfo | null {
    const emptyColumns: Column[] = this.requiredFields.filter((field: Column) => !docDriver[field]);
    return emptyColumns?.length ? {
      sheetName: Page.DRIVERS,
      rowId: docDriver.__rowNum__ + 1,
      emptyColumns
    } : null;
  }
}
