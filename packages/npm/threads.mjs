// Public ESM surface for @libraz/formulon/threads: the pthread build, whose
// `recalcParallel` runs on Web Workers. It needs SharedArrayBuffer, so a
// browser page must be cross-origin isolated (COOP/COEP). The pthread
// workers load formulon_threads_core.js directly, not this shim (see
// scripts/stage.mjs).

import createFormulon from './formulon_threads_core.js';

export default createFormulon;
export * from './common.js';
