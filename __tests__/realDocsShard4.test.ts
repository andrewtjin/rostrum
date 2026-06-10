// REAL-DOCUMENT hide/reverse round-trips — shard 4 of 4.
//
// Jest parallelizes per test FILE; one file holding every sample's CPU-bound round-trip
// was the whole suite's wall (see runHideReverseShard in realDocs.ts for the full why).
// The `realDocs` filename prefix is load-bearing: `--testPathIgnorePatterns realDocs`
// must keep excluding the entire family. realDocs.test.ts guards the (shard, of) wiring.

import { runHideReverseShard } from "./realDocs";

runHideReverseShard(3, 4);
