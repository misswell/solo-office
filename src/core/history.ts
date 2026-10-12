import type { Spreadsheet, WorkbookSnapshot } from "./xlsx.js";

/**
 * Undo/redo over whole-workbook snapshots. An edit touches a few hundred cells
 * at most, so a snapshot is cheaper and far less error-prone than inverse
 * operations, and it keeps structural edits — which move lines on every sheet —
 * restorable as one step.
 */
export class History {
  private undo: { snap: WorkbookSnapshot; rev: number }[] = [];
  private redo: { snap: WorkbookSnapshot; rev: number }[] = [];

  constructor(private readonly limit = 80) {}

  /** Call before a mutation the user should be able to take back. */
  record(wb: Spreadsheet): void {
    this.undo.push({ snap: wb.snapshot(), rev: wb.revision });
    if (this.undo.length > this.limit) this.undo.shift();
    this.redo.length = 0;
  }

  canUndo(wb: Spreadsheet): boolean {
    const last = this.undo[this.undo.length - 1];
    return !!last && last.rev !== wb.revision;
  }

  canRedo(wb: Spreadsheet): boolean {
    const last = this.redo[this.redo.length - 1];
    return !!last && last.rev !== wb.revision;
  }

  undoStep(wb: Spreadsheet): boolean {
    const step = this.undo.pop();
    if (!step || step.rev === wb.revision) return false;
    this.redo.push({ snap: wb.snapshot(), rev: wb.revision });
    wb.restore(step.snap);
    return true;
  }

  redoStep(wb: Spreadsheet): boolean {
    const step = this.redo.pop();
    if (!step || step.rev === wb.revision) return false;
    this.undo.push({ snap: wb.snapshot(), rev: wb.revision });
    wb.restore(step.snap);
    return true;
  }
}
