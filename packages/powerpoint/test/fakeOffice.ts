// A strict, counting fake of the Office.js surface the PowerPoint ops use.
// Like Office.js: loaded values are readable only after a sync, writes and method calls queue until the next sync,
// and a failing statement fails its sync after the statements before it already applied.
// Every PowerPoint.run records how many syncs it made and how many of them changed the document.

export interface FakeShape {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  zOrderPosition: number;
  rotation: number;
  textFrame?: {
    autoSizeSetting: string;
    wordWrap: boolean;
    verticalAlignment: string;
    textRange: {
      text: string;
      font: Record<string, unknown>;
      paragraphFormat: { horizontalAlignment: string; bulletFormat: { visible: boolean } };
    };
  };
  placeholderFormat?: { type: string };
  fill: { type: string; foregroundColor: string; transparency: number };
  lineFormat: { visible: boolean; color: string; weight: number; dashStyle: string };
  table?: {
    rowCount: number;
    columnCount: number;
    values: string[][];
    rows: object[];
    columns: object[];
  };
  group?: { shapes: FakeShape[] };
}

export interface FakeSlide {
  id: string;
  layout: { id: string; name: string };
  shapes: FakeShape[];
}

export interface FakeDeck {
  slides: FakeSlide[];
  layouts: { id: string; name: string }[];
  selectedSlideIds: string[];
}

export interface RunStats {
  syncs: number;
  /** Syncs whose batch changed the document. Selection changes don't count. */
  writeSyncs: number;
  /** Syncs whose batch changed the selection. */
  selectionSyncs: number;
}

export class FakeOfficeError extends Error {
  debugInfo = {};
  constructor(
    readonly code: string,
    message: string,
  ) {
    super(message);
  }
}

const TEXT_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder"]);
let nextId = 1000;

export function fakeShape(
  id: string,
  {
    type = "TextBox",
    text = "",
    placeholder,
    ...rest
  }: Partial<FakeShape> & { text?: string; placeholder?: string } = {},
): FakeShape {
  const shape: FakeShape = {
    id,
    name: `Shape ${id}`,
    type: placeholder ? "Placeholder" : type,
    left: 10,
    top: 10,
    width: 100,
    height: 50,
    zOrderPosition: 0,
    rotation: 0,
    fill: { type: "NoFill", foregroundColor: "", transparency: 0 },
    lineFormat: { visible: false, color: "#000000", weight: 1, dashStyle: "Solid" },
    ...rest,
  };
  if (TEXT_TYPES.has(shape.type))
    shape.textFrame = {
      autoSizeSetting: "AutoSizeNone",
      wordWrap: true,
      verticalAlignment: "Top",
      textRange: {
        text,
        font: {
          name: "Aptos",
          size: 18,
          color: "#000000",
          bold: false,
          italic: false,
          underline: "None",
        },
        paragraphFormat: { horizontalAlignment: "Left", bulletFormat: { visible: false } },
      },
    };
  if (placeholder) shape.placeholderFormat = { type: placeholder };
  return shape;
}

export function fakeSlide(
  id: string,
  shapes: FakeShape[],
  layout = "Title and Content",
): FakeSlide {
  return { id, layout: { id: `layout-${layout}`, name: layout }, shapes };
}

type StepKind = "read" | "write" | "select";
type Method = (state: Node, ...args: any[]) => unknown;

interface Node {
  kind: string;
  /** The model object, or null for a null object. Throws when the object doesn't exist. */
  resolve: () => any;
  /** The array that holds the model, for delete, moves and z-order. */
  owner?: () => any[];
  nullable?: boolean;
  /** The object a navigation property was read from. */
  parent?: Node;
  snapshot: Record<string, unknown>;
  children: Map<string, any>;
}

interface LoadTree {
  [prop: string]: LoadTree;
}

const STATE = Symbol("node state");

/** Navigation properties, which need no load, mapped to the kind of object they lead to. */
const NAV: Record<string, string> = {
  slides: "slides",
  shapes: "shapes",
  layout: "data",
  slideMaster: "master",
  slideMasters: "slideMasters",
  themeColorScheme: "themeColorScheme",
  pageSetup: "data",
  layouts: "layouts",
  textFrame: "textFrame",
  textRange: "textRange",
  font: "data",
  paragraphFormat: "data",
  bulletFormat: "data",
  fill: "fill",
  lineFormat: "data",
  placeholderFormat: "data",
  group: "data",
  rows: "rows",
  columns: "columns",
};
const ITEM_KIND: Record<string, string> = {
  slideMasters: "master",
  slides: "slide",
  shapes: "shape",
  layouts: "data",
  rows: "data",
  columns: "data",
};

function parseLoad(props: string | string[]): LoadTree {
  const tree: LoadTree = {};
  for (const path of Array.isArray(props) ? props : props.split(",")) {
    let level = tree;
    for (const part of path.trim().split("/")) level = level[part] ??= {};
  }
  return tree;
}

/** Installs Office, OfficeExtension and PowerPoint globals backed by `deck`. */
export function installFakeOffice(deck: FakeDeck, apiVersion = "1.10") {
  const runs: RunStats[] = [];
  const master = {
    id: "master-1",
    name: "Office Theme",
    layouts: deck.layouts,
    shapes: [],
    themeColorScheme: {},
  };
  const presentation = {
    get slides() {
      return deck.slides;
    },
    slideMasters: [master],
    pageSetup: { slideWidth: 960, slideHeight: 540 },
  };
  /** Exported slides by their base64 token: a copy of the slide as it was when exported. */
  const exported = new Map<string, FakeSlide>();
  /** When set, every sync after one that changed the document fails, like a read-back that loses the connection. */
  const faults = { failReadsAfterWrite: false };
  const supportedMinor = Number(apiVersion.split(".")[1]);

  function createContext() {
    const stats: RunStats = { syncs: 0, writeSyncs: 0, selectionSyncs: 0 };
    runs.push(stats);
    const queue: { kind: StepKind; run: () => void }[] = [];
    const enqueue = (kind: StepKind, run: () => void) => void queue.push({ kind, run });
    const write = (run: () => void) => enqueue("write", run);

    const context: any = {
      async sync() {
        stats.syncs += 1;
        const steps = queue.splice(0);
        if (faults.failReadsAfterWrite && stats.writeSyncs > 0)
          throw new FakeOfficeError("GeneralException", "The connection dropped.");
        if (steps.some((step) => step.kind === "write")) stats.writeSyncs += 1;
        if (steps.some((step) => step.kind === "select")) stats.selectionSyncs += 1;
        for (const step of steps) step.run();
      },
    };

    function node(spec: Omit<Node, "snapshot" | "children">): any {
      const state: Node = { ...spec, snapshot: {}, children: new Map() };
      const proxy: any = new Proxy(state, {
        get(_, prop) {
          if (prop === STATE) return state;
          if (typeof prop !== "string" || prop === "then") return undefined;
          if (prop === "context") return context;
          if (prop === "load")
            return (props: string | string[] = []) => {
              enqueue("read", () => loadInto(state, parseLoad(props)));
              return proxy;
            };
          if (prop === "toJSON") return () => state.snapshot;
          const method =
            methods[state.kind]?.[prop] ??
            (state.kind in ITEM_KIND ? collectionMethods[prop] : undefined);
          if (method) return (...args: unknown[]) => method(state, ...args);
          if (prop in NAV) return child(state, prop);
          if (!(prop in state.snapshot))
            throw new Error(
              `PropertyNotLoaded: ${state.kind}.${prop} was read before load and sync`,
            );
          return state.snapshot[prop];
        },
        set(_, prop, value) {
          write(() => {
            const target = state.resolve();
            if (target !== null) target[prop as string] = value;
          });
          return true;
        },
      });
      return proxy;
    }

    function child(parent: Node, prop: string): any {
      let found = parent.children.get(prop);
      if (!found) {
        found = node({
          kind: NAV[prop]!,
          resolve: () => {
            const model = parent.resolve();
            if (model === null) return null;
            if (prop === "slideMaster") return master;
            if (model[prop] === undefined)
              throw new FakeOfficeError("GeneralException", `This ${parent.kind} has no ${prop}`);
            return model[prop];
          },
          parent,
          ...(parent.nullable && { nullable: true }),
        });
        parent.children.set(prop, found);
      }
      return found;
    }

    function loadInto(state: Node, tree: LoadTree): void {
      let model: any;
      try {
        model = state.resolve();
      } catch (error) {
        if (!state.nullable) throw error;
        model = null;
      }
      state.snapshot.isNullObject = model === null;
      if (model === null) return;
      const props = Object.keys(tree).length > 0 ? tree : scalarProps(model);
      for (const [prop, sub] of Object.entries(props)) {
        if (prop === "items") {
          const items = (model as any[]).map((item) => itemNode(state, item));
          state.snapshot.items = items;
          for (const item of items) loadInto(item[STATE], sub);
        } else if (prop in NAV) {
          loadInto(child(state, prop)[STATE], sub);
        } else if (prop === "hasText") {
          state.snapshot.hasText = model.textRange.text.length > 0;
        } else {
          if (!(prop in model))
            throw new FakeOfficeError("InvalidArgument", `A ${state.kind} has no property ${prop}`);
          state.snapshot[prop] = model[prop];
        }
      }
    }

    const scalarProps = (model: object): LoadTree =>
      Object.fromEntries(
        Object.keys(model)
          .filter((key) => !(key in NAV))
          .map((key) => [key, {}]),
      );

    function itemNode(collection: Node, model: any): any {
      const owner = () => collection.resolve() as any[];
      return node({
        kind: ITEM_KIND[collection.kind]!,
        resolve: () => {
          if (!owner().includes(model))
            throw new FakeOfficeError("ItemNotFound", "The item was deleted.");
          return model;
        },
        owner,
      });
    }

    /** Like Office.js, a looked-up object binds to the item it first resolved to, even if that item moves later. */
    function itemLookup(collection: Node, find: (items: any[]) => any, nullable: boolean): any {
      const owner = () => collection.resolve() as any[];
      let bound: any;
      return node({
        kind: ITEM_KIND[collection.kind]!,
        resolve: () => {
          bound ??= find(owner());
          if (!bound || !owner().includes(bound))
            throw new FakeOfficeError("ItemNotFound", "The requested item doesn't exist.");
          return bound;
        },
        owner,
        ...(nullable && { nullable: true }),
      });
    }

    const collectionMethods: Record<string, Method> = {
      getItem: (state, id: string) =>
        itemLookup(state, (items) => items.find((item) => item.id === id), false),
      getItemOrNullObject: (state, id: string) =>
        itemLookup(state, (items) => items.find((item) => item.id === id), true),
      getItemAt: (state, index: number) => itemLookup(state, (items) => items[index], false),
    };

    /** A node for an object a queued statement creates; it resolves once that statement has run. */
    function created(collection: Node, make: () => any): any {
      let model: any;
      const owner = () => collection.resolve() as any[];
      write(() => {
        model = make();
        owner().push(model);
      });
      return node({
        kind: ITEM_KIND[collection.kind]!,
        resolve: () => {
          if (!model)
            throw new FakeOfficeError(
              "GeneralException",
              "Used before its creating statement ran.",
            );
          return model;
        },
        owner,
      });
    }

    function clientResult(compute: () => unknown) {
      let value: unknown;
      let ready = false;
      enqueue("read", () => {
        value = compute();
        ready = true;
      });
      return {
        get value() {
          if (!ready) throw new Error("PropertyNotLoaded: ClientResult.value was read before sync");
          return value;
        },
      };
    }

    function move(state: Node, place: (items: any[], model: any) => void): void {
      write(() => {
        const model = state.resolve();
        const owner = state.owner!();
        owner.splice(owner.indexOf(model), 1);
        place(owner, model);
      });
    }

    const addShape = (state: Node, init: Partial<FakeShape> & { text?: string }) =>
      created(state, () => fakeShape(`${nextId++}`, init));

    const methods: Record<string, Record<string, Method>> = {
      presentation: {
        insertSlidesFromBase64: (_, base64: string, options: { targetSlideId?: string } = {}) =>
          write(() => {
            const source = exported.get(base64);
            if (!source) throw new FakeOfficeError("InvalidArgument", "Not an exported slide.");
            const shapes = structuredClone(source.shapes);
            const at = options.targetSlideId
              ? deck.slides.findIndex((slide) => slide.id === options.targetSlideId) + 1
              : 0;
            deck.slides.splice(at, 0, fakeSlide(`${nextId++}#`, shapes, source?.layout.name));
          }),
        setSelectedSlides: (_, ids: string[]) =>
          enqueue("select", () => {
            deck.selectedSlideIds = [...ids];
          }),
        getSelectedSlides: () =>
          node({
            kind: "slides",
            resolve: () => deck.slides.filter((slide) => deck.selectedSlideIds.includes(slide.id)),
          }),
        getSelectedShapes: () => node({ kind: "shapes", resolve: () => [] }),
        getSelectedTextRangeOrNullObject: () =>
          node({ kind: "textRange", resolve: () => null, nullable: true }),
      },
      slides: {
        add: (state, options: { layoutId?: string } = {}) =>
          created(state, () => {
            const layout =
              deck.layouts.find((candidate) => candidate.id === options.layoutId) ??
              deck.layouts[0]!;
            return fakeSlide(
              `${nextId++}#`,
              [fakeShape(`${nextId++}`, { placeholder: "Title" })],
              layout.name,
            );
          }),
      },
      slide: {
        exportAsBase64: (state) =>
          clientResult(() => {
            const token = `pptx:${nextId++}`;
            exported.set(token, structuredClone(state.resolve()));
            return token;
          }),
        applyLayout: (state, layout: any) =>
          write(() => {
            state.resolve().layout = { ...layout[STATE].resolve() };
          }),
        delete: (state) =>
          write(() => {
            const { id } = state.resolve();
            if (deck.selectedSlideIds.includes(id))
              throw new FakeOfficeError(
                "GeneralException",
                "Sorry, we ran into a problem (deleted the slide on screen).",
              );
            state.owner!().splice(state.owner!().indexOf(state.resolve()), 1);
          }),
        moveTo: (state, index: number) =>
          move(state, (items, model) => items.splice(index, 0, model)),
      },
      shapes: {
        addTextBox: (state, text: string, bounds: object) =>
          addShape(state, { type: "TextBox", text, ...bounds }),
        addGeometricShape: (state, _geometry: string, bounds: object) =>
          addShape(state, { type: "GeometricShape", ...bounds }),
        addLine: (state, _connector: string, bounds: object) =>
          addShape(state, { type: "Line", ...bounds }),
        addTable: (
          state,
          rows: number,
          columns: number,
          { values, left, top, width, height }: any = {},
        ) =>
          addShape(state, {
            type: "Table",
            left,
            top,
            width,
            height,
            table: {
              rowCount: rows,
              columnCount: columns,
              values: Array.from({ length: rows }, (_r, r) =>
                Array.from({ length: columns }, (_c, c) => values?.[r]?.[c] ?? ""),
              ),
              rows: Array.from({ length: rows }, () => ({})),
              columns: Array.from({ length: columns }, () => ({})),
            },
          }),
      },
      shape: {
        delete: (state) => move(state, () => {}),
        setZOrder: (state, position: string) =>
          move(state, (items, model) =>
            position === "BringToFront" ? items.push(model) : items.unshift(model),
          ),
        setHyperlink: () => write(() => {}),
        getTable: (state) =>
          node({
            kind: "table",
            resolve: () => {
              const table = state.resolve().table;
              if (!table) throw new FakeOfficeError("InvalidArgument", "The shape isn't a table.");
              return table;
            },
          }),
        getTextFrameOrNullObject: (state) =>
          node({
            kind: "textFrame",
            resolve: () => state.resolve().textFrame ?? null,
            nullable: true,
          }),
      },
      themeColorScheme: {
        getThemeColor: () => clientResult(() => "#0F766E"),
      },
      textRange: {
        getSubstring: () => node({ kind: "textRange", resolve: () => ({ text: "", font: {} }) }),
        setHyperlink: () => write(() => {}),
      },
      fill: {
        setSolidColor: (state, color: string) =>
          write(() => Object.assign(state.resolve(), { type: "Solid", foregroundColor: color })),
        clear: (state) => write(() => Object.assign(state.resolve(), { type: "NoFill" })),
        setImage: (state) => write(() => Object.assign(state.resolve(), { type: "Picture" })),
      },
      table: {
        getCellOrNullObject: (state, row: number, column: number) =>
          node({
            kind: "data",
            resolve: () => {
              const values: string[][] = state.resolve().values;
              const cell: Record<string, unknown> = { font: {}, fill: {} };
              return new Proxy(cell, {
                set(target, prop, value) {
                  if (prop === "text") values[row]![column] = value;
                  else target[prop as string] = value;
                  return true;
                },
              });
            },
          }),
        mergeCells: () => write(() => {}),
      },
      rows: {
        add: (state, index: number, count: number) =>
          write(() => {
            const table = state.parent!.resolve();
            table.rowCount += count;
            table.rows.push(...Array.from({ length: count }, () => ({})));
            table.values.splice(
              index,
              0,
              ...Array.from({ length: count }, () => Array(table.columnCount).fill("")),
            );
          }),
        deleteRows: (state, rows: unknown[]) =>
          write(() => {
            const table = state.parent!.resolve();
            table.rowCount -= rows.length;
            table.rows.splice(0, rows.length);
            table.values.splice(0, rows.length);
          }),
      },
      columns: {
        add: (state, index: number, count: number) =>
          write(() => {
            const table = state.parent!.resolve();
            table.columnCount += count;
            table.columns.push(...Array.from({ length: count }, () => ({})));
            for (const row of table.values) row.splice(index, 0, ...Array(count).fill(""));
          }),
        deleteColumns: (state, columns: unknown[]) =>
          write(() => {
            const table = state.parent!.resolve();
            table.columnCount -= columns.length;
            table.columns.splice(0, columns.length);
            for (const row of table.values) row.splice(0, columns.length);
          }),
      },
    };

    context.presentation = node({ kind: "presentation", resolve: () => presentation });
    return { context, queue };
  }

  const globals = globalThis as any;
  globals.Office = {
    context: {
      requirements: {
        isSetSupported: (_set: string, version: string) =>
          Number(version.split(".")[1]) <= supportedMinor,
      },
    },
  };
  globals.OfficeExtension = { Error: FakeOfficeError };
  globals.PowerPoint = {
    async run<T>(batch: (context: unknown) => Promise<T>): Promise<T> {
      const { context, queue } = createContext();
      const result = await batch(context);
      // Office.js syncs whatever the batch left queued when its promise resolves.
      if (queue.length > 0) await context.sync();
      return result;
    },
  };

  return {
    faults,
    /** Totals over every PowerPoint.run since the last reset. */
    totals(): RunStats & { runs: number } {
      return runs.reduce<RunStats & { runs: number }>(
        (sum, run) => ({
          runs: sum.runs + 1,
          syncs: sum.syncs + run.syncs,
          writeSyncs: sum.writeSyncs + run.writeSyncs,
          selectionSyncs: sum.selectionSyncs + run.selectionSyncs,
        }),
        { runs: 0, syncs: 0, writeSyncs: 0, selectionSyncs: 0 },
      );
    },
    reset(): void {
      runs.length = 0;
    },
  };
}
