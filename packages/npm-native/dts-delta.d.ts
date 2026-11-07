// Members declared only by @libraz/formulon-native. tools/codegen/gen_native_dts.py
// splices the members of `Workbook` below into the shared `Workbook` interface of
// src/wasm/formulon.d.ts when it generates index.d.ts.
//
// `FormulonOptions` is shared as declared there: the native package ignores `wasm` like
// the other option keys, so it needs no delta of its own.

export interface Workbook {
  /** Estimated heap bytes the native workbook occupies.
   *
   *  The addon reports this figure to V8 as external memory when a
   *  workbook is created, loaded, recalculated or disposed, so a runtime
   *  that acts on the hint can weigh workbooks that are large in native
   *  memory but pointer-sized on the JS heap. Calling this method also
   *  refreshes that report, which is worth doing after a long run of
   *  cell writes. The value covers the cell store, shared strings,
   *  passthrough parts and workbook metadata; it is an estimate, not an
   *  allocation ledger. */
  getMemoryUsage(): number;
}
