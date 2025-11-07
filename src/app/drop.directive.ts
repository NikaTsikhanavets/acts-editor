import { Directive, EventEmitter, HostListener, Output } from '@angular/core';
import { DomSanitizer, SafeUrl } from '@angular/platform-browser';

export interface FileHandle {
  file: File,
  url: SafeUrl
}

@Directive({
    selector: '[appDrop]',
    standalone: true
})
export class DropDirective {
  @Output() files: EventEmitter<FileHandle> = new EventEmitter();

  constructor(private readonly sanitizer: DomSanitizer) { }

  @HostListener("dragover", ["$event"]) public onDragOver(evt: DragEvent) {
    evt.preventDefault();
    evt.stopPropagation();
  }

  @HostListener("dragleave", ["$event"]) public onDragLeave(evt: DragEvent) {
    evt.preventDefault();
    evt.stopPropagation();
  }

  @HostListener('drop', ['$event']) public onDrop(evt: DragEvent) {
    evt.preventDefault();
    evt.stopPropagation();

    let files: FileHandle[] = [];
    const filesCount: number = evt.dataTransfer?.files?.length || 0;
    for (let i = 0; i < filesCount; i++) {
      const file = evt.dataTransfer?.files[i];
      const url = this.sanitizer.bypassSecurityTrustUrl((window as any).URL.createObjectURL(file));

      if (file) {
        files.push({ file, url });
      }
    }
    if (files.length > 0) {
      this.files.emit(files[0]);
    }
  }
}
