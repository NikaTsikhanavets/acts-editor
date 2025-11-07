import { ChangeDetectionStrategy, Component, EventEmitter, Input, Output } from '@angular/core';
import { ErrorInfo } from '../../interfaces/error-info.interface';

@Component({
    selector: 'app-errors',
    templateUrl: './errors.component.html',
    styleUrls: ['./errors.component.scss'],
    changeDetection: ChangeDetectionStrategy.OnPush,
    standalone: true
})
export class ErrorsComponent {
  @Input() public errors: ErrorInfo[] = [];

  @Output() public return: EventEmitter<void> = new EventEmitter<void>();
}
