import type {
  FormDefinition,
  SheetDefinition,
  ClusterDefinition,
  ImageDefinition,
  MergedCell,
  CellStyle,
  PageSettings,
  PrintArea,
} from "./formDefinition";

export interface DesignerState {
  form: FormDefinition;
  activeSheetId: string | null;
  dirty: boolean;
  selectedClusterId: string | null;
  zoom: number;
}

export type DesignerListener = (state: DesignerState) => void;

const DEFAULT_PAGE_SETTINGS: PageSettings = {
  paperSize: "Letter",
  orientation: "portrait",
  widthPt: 612,
  heightPt: 792,
  margins: { left: 70, top: 70, right: 70, bottom: 70 },
  centerHorizontally: false,
  centerVertically: false,
  zoom: 100,
  fitToPagesWide: 0,
  fitToPagesTall: 0,
};

function createEmptyForm(): FormDefinition {
  return {
    workbook: {
      title: "Untitled Form",
      author: "",
      created: new Date().toISOString(),
      modified: new Date().toISOString(),
      version: "1.0",
      description: "",
    },
    sheets: [],
    clusters: [],
    images: [],
    metadata: {},
  };
}

export function createDefaultSheet(index: number): SheetDefinition {
  return {
    id: `sheet_${index}`,
    name: `Sheet${index + 1}`,
    index,
    pageSettings: { ...DEFAULT_PAGE_SETTINGS },
    printArea: null,
    backgroundImage: null,
    backgroundWidth: 0,
    backgroundHeight: 0,
    thumbnail: null,
    rowHeights: {},
    columnWidths: {},
    mergedCells: [],
    freezePane: null,
    cellStyles: {},
    cellValues: {},
  };
}

export class DesignerController {
  private state: DesignerState;
  private listeners: Set<DesignerListener> = new Set();
  private undoStack: FormDefinition[] = [];
  private redoStack: FormDefinition[] = [];
  private batchDepth = 0;
  private pendingNotify = false;

  constructor() {
    this.state = {
      form: createEmptyForm(),
      activeSheetId: null,
      dirty: false,
      selectedClusterId: null,
      zoom: 1,
    };
  }

  getState(): DesignerState {
    return this.state;
  }

  subscribe(listener: DesignerListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  beginBatch(): void {
    this.batchDepth++;
  }

  endBatch(): void {
    if (this.batchDepth > 0) {
      this.batchDepth--;
    }
    if (this.batchDepth === 0 && this.pendingNotify) {
      this.pendingNotify = false;
      this.notify();
    }
  }

  private pushUndo(): void {
    this.undoStack.push(JSON.parse(JSON.stringify(this.state.form)));
    if (this.undoStack.length > 50) this.undoStack.shift();
    this.redoStack = [];
  }

  undo(): void {
    if (this.undoStack.length === 0) return;
    this.redoStack.push(JSON.parse(JSON.stringify(this.state.form)));
    this.state.form = this.undoStack.pop()!;
    this.state.dirty = true;
    this.notify();
  }

  redo(): void {
    if (this.redoStack.length === 0) return;
    this.undoStack.push(JSON.parse(JSON.stringify(this.state.form)));
    this.state.form = this.redoStack.pop()!;
    this.state.dirty = true;
    this.notify();
  }

  private markDirty(): void {
    this.state.dirty = true;
    this.state.form.workbook.modified = new Date().toISOString();
    if (this.batchDepth > 0) {
      this.pendingNotify = true;
    } else {
      this.notify();
    }
  }

  private notify(): void {
    const s = this.state;
    this.listeners.forEach((l) => l(s));
  }

  setActiveSheet(sheetId: string | null): void {
    if (this.state.activeSheetId !== sheetId) {
      this.state.activeSheetId = sheetId;
      this.notify();
    }
  }

  setZoom(zoom: number): void {
    this.state.zoom = Math.max(0.25, Math.min(3, zoom));
    this.notify();
  }

  selectCluster(clusterId: string | null): void {
    this.state.selectedClusterId = clusterId;
    this.notify();
  }

  loadForm(form: FormDefinition): void {
    this.pushUndo();
    this.state.form = form;
    this.state.dirty = false;
    this.state.activeSheetId = form.sheets[0]?.id ?? null;
    this.state.selectedClusterId = null;
    this.notify();
  }

  resetForm(): void {
    this.pushUndo();
    this.state.form = createEmptyForm();
    this.state.activeSheetId = null;
    this.state.dirty = false;
    this.state.selectedClusterId = null;
    this.notify();
  }

  private getSheet(sheetId: string): SheetDefinition {
    const sheet = this.state.form.sheets.find((s) => s.id === sheetId);
    if (!sheet) throw new Error(`Sheet not found: ${sheetId}`);
    return sheet;
  }

  addSheet(name?: string): SheetDefinition {
    this.pushUndo();
    const index = this.state.form.sheets.length;
    const sheet = createDefaultSheet(index);
    if (name) sheet.name = name;
    this.state.form.sheets.push(sheet);
    this.markDirty();
    return sheet;
  }

  setSheetName(sheetId: string, name: string): void {
    this.pushUndo();
    const sheet = this.getSheet(sheetId);
    sheet.name = name;
    this.markDirty();
  }

  setPrintArea(sheetId: string, printArea: PrintArea | null): void {
    this.pushUndo();
    const sheet = this.getSheet(sheetId);
    sheet.printArea = printArea;
    this.markDirty();
  }

  setPageSettings(sheetId: string, settings: Partial<PageSettings>): void {
    this.pushUndo();
    const sheet = this.getSheet(sheetId);
    Object.assign(sheet.pageSettings, settings);
    this.markDirty();
  }

  setRowHeight(sheetId: string, row: number, height: number): void {
    this.pushUndo();
    const sheet = this.getSheet(sheetId);
    if (height > 0) sheet.rowHeights[row] = height;
    else delete sheet.rowHeights[row];
    this.markDirty();
  }

  setColumnWidth(sheetId: string, col: number, width: number): void {
    this.pushUndo();
    const sheet = this.getSheet(sheetId);
    if (width > 0) sheet.columnWidths[col] = width;
    else delete sheet.columnWidths[col];
    this.markDirty();
  }

  addMergedCell(sheetId: string, merged: MergedCell): void {
    this.pushUndo();
    const sheet = this.getSheet(sheetId);
    const idx = sheet.mergedCells.findIndex((m) => m.address === merged.address);
    if (idx >= 0) sheet.mergedCells[idx] = merged;
    else sheet.mergedCells.push(merged);
    this.markDirty();
  }

  removeMergedCell(sheetId: string, address: string): void {
    this.pushUndo();
    const sheet = this.getSheet(sheetId);
    sheet.mergedCells = sheet.mergedCells.filter((m) => m.address !== address);
    this.markDirty();
  }

  setCellStyle(sheetId: string, cellAddress: string, style: CellStyle): void {
    this.pushUndo();
    const sheet = this.getSheet(sheetId);
    sheet.cellStyles[cellAddress] = style;
    this.markDirty();
  }

  setCellValue(sheetId: string, cellAddress: string, value: string): void {
    this.pushUndo();
    const sheet = this.getSheet(sheetId);
    sheet.cellValues[cellAddress] = value;
    this.markDirty();
  }

  setFreezePane(sheetId: string, pane: string | null): void {
    this.pushUndo();
    const sheet = this.getSheet(sheetId);
    sheet.freezePane = pane;
    this.markDirty();
  }

  addCluster(cluster: ClusterDefinition): void {
    this.pushUndo();
    const idx = this.state.form.clusters.findIndex(
      (c) => c.clusterId === cluster.clusterId
    );
    if (idx >= 0) this.state.form.clusters[idx] = cluster;
    else this.state.form.clusters.push(cluster);
    this.markDirty();
  }

  removeCluster(clusterId: string): void {
    this.pushUndo();
    this.state.form.clusters = this.state.form.clusters.filter(
      (c) => c.clusterId !== clusterId
    );
    if (this.state.selectedClusterId === clusterId) {
      this.state.selectedClusterId = null;
    }
    this.markDirty();
  }

  updateCluster(
    clusterId: string,
    updates: Partial<ClusterDefinition>
  ): void {
    this.pushUndo();
    const cluster = this.state.form.clusters.find(
      (c) => c.clusterId === clusterId
    );
    if (cluster) {
      Object.assign(cluster, updates);
      this.markDirty();
    }
  }

  getClustersForSheet(sheetId: string): ClusterDefinition[] {
    return this.state.form.clusters.filter((c) => c.sheetId === sheetId);
  }

  addImage(image: ImageDefinition): void {
    this.pushUndo();
    const idx = this.state.form.images.findIndex((i) => i.id === image.id);
    if (idx >= 0) this.state.form.images[idx] = image;
    else this.state.form.images.push(image);
    this.markDirty();
  }

  removeImage(imageId: string): void {
    this.pushUndo();
    this.state.form.images = this.state.form.images.filter(
      (i) => i.id !== imageId
    );
    this.markDirty();
  }

  getFormSnapshot(): FormDefinition {
    return JSON.parse(JSON.stringify(this.state.form));
  }

  save(): FormDefinition {
    this.state.dirty = false;
    this.notify();
    return this.getFormSnapshot();
  }
}
