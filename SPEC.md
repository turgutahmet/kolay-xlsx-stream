# KXSI — Kolay XLSX Stream Index

**An open sidecar format for random access, zone-map statistics, and
prefix integrity inside ordinary XLSX files.**

A KXSI-bearing workbook is a 100% valid XLSX file that Excel,
LibreOffice, PhpSpreadsheet, OpenSpout and every other OOXML consumer
opens normally. A KXSI-aware reader additionally gets O(1) row counts,
O(sync period) random row access, Parquet-style block pruning for
numeric predicates, and per-block integrity pins — all from one small
binary part, over local files or HTTP/S3 range requests alike.

The key words MUST, MUST NOT, REQUIRED, SHOULD, SHOULD NOT, MAY, and
OPTIONAL in this document are to be interpreted as described in
[RFC 2119](https://www.rfc-editor.org/rfc/rfc2119).

---

## 1. Status and document versioning

| | |
|---|---|
| Document version | **1.1.0** |
| Format version described | KXSI binary version **2** (the `version` header byte) |
| Reference implementation | `kolay/xlsx-stream` ≥ 3.2.0 (PHP) — writer + reader |
| Conformance suite | `tests/SpecVectors/` in the reference repository (§8) |

This document is versioned by semver **independently of the format's
`version` byte**. Editorial fixes bump the patch; new registered TLV
sections or clarified normative language bump the minor; a change that
alters the meaning of already-specified bytes bumps the major *and* the
format's `version` byte. By design (§4), the minor is expected to move
while the format byte stays at 2 indefinitely.

## 2. Container placement

### 2.1 The part

The index is stored as a single OPC package part:

```
xl/_kxs/index.bin
```

- The part MUST be a ZIP entry in the workbook archive. It MAY be
  deflated or stored; readers locate and read it through the ZIP
  central directory like any other entry.
- The part MUST be declared in `[Content_Types].xml` with an explicit
  Override:

  ```xml
  <Override PartName="/xl/_kxs/index.bin"
            ContentType="application/octet-stream"/>
  ```

  This is REQUIRED, not cosmetic. The `bin` extension has no Default
  extension mapping in the package, and OPC forbids parts without a
  content type — empirically, **MS Excel triggers repair mode on open
  when the Override is missing**, even though every other part is
  valid. Repair mode then strips the part. `application/octet-stream`
  declares the part as opaque binary so validating consumers carry it
  through without inspecting it.
- The part MUST NOT be referenced from any relationships (`.rels`)
  chain. It is deliberately an *unreferenced* part: OOXML consumers
  only traverse parts reachable from relationships, so the sidecar is
  invisible to them, while OPC validation still passes because the
  content type is declared.

### 2.2 Survival and staleness

An unreferenced part survives archive-level copying (S3 transfer,
`zip`/`unzip`, checksum-preserving pipelines) byte-for-byte. It does
**not** reliably survive an application-level resave: **Excel drops the
part when a user edits and re-saves the workbook**, and other OOXML
editors may either drop it or — worse — carry it through while
rewriting the sheet parts it describes.

That second case is the dangerous one: the sidecar's offsets would then
point into a deflate stream that no longer exists. The format defends
against it with **CRC pinning**:

- The writer records each sheet's whole-entry CRC-32 (`sheet_crc32`,
  §3.2) — the same value it writes into the ZIP data descriptor and
  central directory for that entry.
- Before using the index, a reader MUST compare each sheet's
  `sheet_crc32` against the CRC-32 recorded in the **live** ZIP central
  directory for the same entry name. On any mismatch (or missing
  entry), the reader MUST treat the entire index as stale and MUST NOT
  use it for random access; it SHOULD fall back silently to a
  sequential scan. A rewritten sheet always gets a fresh deflate stream
  and therefore (except with 2⁻³² probability) a fresh CRC, so this
  check catches every realistic resave.
- Additionally, a reader MUST verify that every sync point's
  `comp_offset` is strictly less than the entry's `compressed_size` as
  recorded in the live central directory, and reject the index
  otherwise (§7).

## 3. Normative binary layout

All multi-byte integers are **little-endian**. There is no padding or
alignment anywhere in the format; every structure is packed. Entry
paths are UTF-8. "CRC-32" throughout means the CRC-32 used by ZIP
(polynomial 0x04C11DB7 reflected, i.e. `crc32b` / PHP `crc32()` /
zlib `crc32()`).

A payload is: **header**, then **core body** (one record per sheet),
then zero or more **TLV sections** (§4), ending exactly at the end of
the part.

### 3.1 Header (16 bytes)

| Offset | Size | Field | Description |
|---|---|---|---|
| 0 | 4 | `magic` | ASCII `KXSI` |
| 4 | 1 | `version` | `2`. Readers MUST reject any other value. |
| 5 | 1 | `flags` | Must-understand bitmask, §5. Writers MUST emit `0`. |
| 6 | 2 | `sheet_count` | uint16 — number of core-body records |
| 8 | 4 | `sync_period` | uint32 — approximate rows between sync points; MUST be > 0 |
| 12 | 4 | `payload_crc32` | uint32 — CRC-32 of **every byte after this header** (core body + all TLV sections) |

`payload_crc32` covers the full remainder of the part, including TLV
sections a given reader does not understand. Readers MUST verify it
before interpreting the body.

### 3.2 Core body — one record per sheet

Records appear in **workbook order** (the order of `<sheet>` elements
in `xl/workbook.xml`). This order is normative: TLV sections that carry
per-sheet payloads (§4) rely on it for alignment. A sheet entry path
MUST NOT appear more than once.

| Size | Field | Description |
|---|---|---|
| 2 | `entry_path_len` | uint16, MUST be ≥ 1 (reference decoder caps at 256, §7) |
| N | `entry_path` | UTF-8 ZIP entry name, e.g. `xl/worksheets/sheet1.xml` |
| 4 | `total_rows` | uint32 — the sheet's total row count **including** the header row |
| 4 | `sheet_crc32` | uint32 — CRC-32 of the sheet entry's complete uncompressed bytes; MUST equal the value in the ZIP central directory at write time (§2.2) |
| 4 | `sync_count` | uint32 — number of sync points that follow |
| 24 × K | `sync_points` | K = `sync_count`, each 24 bytes: |

Each sync point:

| Size | Field |
|---|---|
| 8 | `row` — uint64 |
| 8 | `comp_offset` — uint64 |
| 8 | `uncomp_offset` — uint64 |

### 3.3 Sync-point semantics

A sync point is a **resume marker inside the sheet's deflate stream**.

- `comp_offset` and `uncomp_offset` are relative to the start of the
  sheet entry's **data stream** — the first byte after the ZIP Local
  File Header (and its name/extra fields), *not* an absolute archive
  offset. Readers resolve them to absolute file positions at runtime
  via the live central directory entry for `entry_path`. This is what
  keeps the sidecar valid when the archive is copied, embedded, or has
  entries re-ordered ahead of it.
- **`row` is the first row that a fresh raw-deflate decompression
  started at `comp_offset` will yield.** Equivalently: the uncompressed
  bytes `[0, uncomp_offset)` end exactly between two complete `<row>`
  elements, and the byte at `uncomp_offset` begins the XML of row
  `row`. Rows are 1-based and include the header row (the header, when
  present, is row 1).
- Requirements on the **writer** at every sync point:
  1. The deflate stream MUST be **byte-aligned** at `comp_offset` (the
     writer emits a deflate sync flush — an empty stored block,
     `00 00 FF FF` — ending the prefix; `comp_offset` points *after*
     those marker bytes).
  2. The compression window MUST be **reset**: no compressed byte at or
     after `comp_offset` may back-reference data before
     `uncomp_offset`. (zlib's `Z_FULL_FLUSH` provides both properties;
     `Z_SYNC_FLUSH` provides only alignment and is NOT sufficient.)
  3. The boundary MUST fall **between two complete `<row>` elements** —
     never inside a row's XML.

  Together these guarantee a reader can seek to `comp_offset`,
  initialize a fresh raw-inflate context with no history and no
  priming, and immediately tokenize whole rows.
- Sync points MUST be **strictly increasing** in `row`, in
  `comp_offset`, and in `uncomp_offset`.
- `row` MUST NOT exceed `total_rows + 1`. The value `total_rows + 1` is
  legal: a flush triggered exactly at the end of the sheet's data
  produces a sync point that points past the last row (only the sheet
  footer follows it). Readers MUST accept such a trailing sync point.
- `sync_count` MAY be 0 (a sheet too small to ever cross
  `sync_period`); such a sheet still gets a core record — and a record
  in every per-sheet TLV section.

## 4. TLV extension sections

Everything after the last core-body record, up to the end of the
payload, is a sequence of tagged sections:

| Size | Field |
|---|---|
| 4 | `tag` — four ASCII characters; registered tags are uppercase `A–Z` |
| 4 | `length` — uint32, byte length of the payload that follows |
| N | `payload` |

Framing rules — these are the format's compatibility contract:

- Readers MUST **skip sections whose tag they do not recognize** by
  advancing `length` bytes. Unknown tags are not an error.
- A section's `length` MUST NOT overrun the payload; readers MUST
  reject the sidecar if it does.
- Sections MAY appear in **any order**. Readers MUST NOT require a
  particular order or the presence of any section.
- A writer MUST NOT emit the same tag more than once. Readers
  encountering a duplicate tag MAY use the last occurrence or reject
  the sidecar. (The reference decoder uses the last occurrence.)
- Writers MUST NOT emit trailing bytes that do not form a complete TLV
  section. (The reference decoder ignores a trailing fragment shorter
  than the 8-byte tag+length frame; readers MAY reject it instead.)
- **The `version` byte does not change when sections are added.** This
  is by construction, not accident: the version-2 decoder shipped
  before any TLV section existed parses exactly `sheet_count` core
  records and ignores everything after them, while `payload_crc32`
  always covered the full body — so a section-bearing sidecar is
  *invisible* to older readers rather than fatal, and they keep their
  random access. New sections ship without stranding a single deployed
  reader. Bump the version byte **only** for a change that alters the
  meaning of bytes already specified here. (The Parquet/ORC lesson: the
  container format outlives any single feature decision.)

Sections whose payload repeats per sheet MUST order those per-sheet
records in **core-body order** (§3.2) and MUST include a record for
*every* sheet, even when it is empty for that sheet — alignment is
positional, there are no per-record entry names.

### 4.1 Registered: `STAT` — per-block column statistics (zone maps)

Per-block min/max/sum/count statistics for chosen columns, plus a
per-sheet sortedness verdict. Enables block pruning for numeric range
predicates, sidecar-only column aggregates, and near-O(1) point lookup
on sorted columns.

**Block model.** A sheet's sync points divide it into
`sync_count + 1` *blocks*: block `k` (0-based) spans the rows between
sync points `k−1` and `k`; block 0 starts at row 1 (the top of the
sheet); block `sync_count` is the tail after the last sync point
(possibly empty). Consequently, for every column of every sheet:

> `block_count` MUST equal `sync_count + 1`.

Readers MUST reject a `STAT` section that violates this.

**Payload**, repeated per sheet in core-body order:

| Size | Field |
|---|---|
| 2 | `tracked_column_count` — uint16 (0 when the sheet has no tracked columns) |

then per tracked column:

| Size | Field |
|---|---|
| 2 | `column` — uint16, **1-based**; MUST be ≥ 1 |
| 1 | `flags` — bit0: values sorted ascending across the whole sheet; bit1: sorted descending. Both bits describe **numeric values in data rows only** — non-numeric cells and the header row are invisible to the ordering check. Other bits reserved, writers MUST emit 0. |
| 4 | `block_count` — uint32, MUST be `sync_count + 1` |

then per block (32 bytes each):

| Size | Field |
|---|---|
| 8 | `min` — float64 (IEEE 754, little-endian); meaningless when `count` == 0 |
| 8 | `max` — float64; meaningless when `count` == 0 |
| 8 | `sum` — float64 |
| 4 | `count` — uint32, numeric values folded into min/max/sum |
| 4 | `other` — uint32, nulls + non-numeric values in the block |

**Numeric interpretation.** A cell is "numeric" for statistics purposes
if the writer would render it as a number: ints and floats as-is,
numeric strings by their numeric value, booleans as 0/1, dates/times as
their Excel serial number. Everything else counts into `other`.

Two normative laws govern every `STAT` producer:

1. **Stats widen, never narrow.** When in doubt whether a value is
   numeric, a writer MUST fold it into min/max/sum rather than omit it.
   An over-wide `[min, max]` only makes pruning less selective — a
   value the stats missed could cause a reader to skip a block that
   actually matches, which is silent data loss at query time.
2. **Block statistics MUST account for every row the scan path can
   yield, including preamble-emitted rows.** In particular the header
   row (row 1) is written via the sheet preamble, not the row-writing
   path — but a full scan can still match it (a numeric-looking header
   passes numeric predicates). Its cells MUST be folded into block 0's
   min/max/sum/count/other. The sortedness flags deliberately exclude
   it: they describe the *data* ordering.

**Sizing.** 1 sheet × 4M rows at `sync_period` 10 000: core ≈ 9.7 KB;
each `STAT` column adds ≈ 12.8 KB.

### 4.2 Registered: `SCRC` — running CRC-32 at each sync point

Per-sheet running CRC-32 of the sheet's **uncompressed** bytes,
captured at every sync point. This is the shared integrity primitive
for resumable exports (resume the writer's live CRC from a pinned
prefix), appendable files (prove the prefix untouched before cutting at
a sync point), truncation detection, and future block-granular signing
(`SIGN`).

**Payload**, repeated per sheet in core-body order:

| Size | Field |
|---|---|
| 4 | `count` — uint32, MUST equal the sheet's `sync_count` |
| 4 × K | `running_crc` — uint32 each, K = `count` |

Semantics — aligned 1:1 with the sheet's sync points:

> `running_crc[k]` MUST equal the CRC-32 of the uncompressed sheet
> bytes `[0, sync_points[k].uncomp_offset)` — i.e. exactly the prefix
> that precedes sync point `k`'s row.

Readers MUST reject an `SCRC` section whose `count` differs from the
sheet's `sync_count` (the mirror of `STAT`'s `block_count` invariant).
A sheet with zero sync points contributes a `count = 0` record.

Note the relationship to the core `sheet_crc32`: it is the same running
CRC carried to the end of the entry. A verifier holding the first
`uncomp_offset[k]` uncompressed bytes can prove that prefix untampered
without inflating past it; a writer resuming at sync point `k` seeds
its CRC state from `running_crc[k]` and, on finish, must land exactly
on `sheet_crc32` — a free end-to-end self-check.

The reference writer emits `SCRC` whenever the index itself is enabled:
the cost is 4 bytes per sync point plus one CRC-state copy per
`sync_period` rows (measured ≈ 70 ns, ~0.4 % of the CRC update it rides
along with).

### 4.3 Registered: `TDIG` — per-column t-digest sketches (approximate quantiles)

One merging t-digest (Dunning) per tracked column per sheet, over the
column's numeric values across the **whole sheet** — not per block.
Unlike `STAT`, this is not a pruning structure: it exists so a reader
can answer `quantile` / `median` for a multi-GB file from the sidecar
alone, with zero row reads.

`TDIG` and `CHLL` (§4.4) share one generic **section frame**, repeated
per sheet in core-body order:

| Size | Field |
|---|---|
| 2 | `tracked_column_count` — uint16 (0 when the sheet has no tracked columns) |

then per tracked column, in ascending column order:

| Size | Field |
|---|---|
| 2 | `column` — uint16, **1-based**; MUST be ≥ 1 |
| 4 | `payload_len` — uint32, byte length of the sketch payload |
| N | `payload` — the sketch's own serialized form, below |

**T-digest payload** (all multi-byte fields little-endian):

| Size | Field |
|---|---|
| 2 | `format_version` — uint16, `1`. Readers MUST reject other values. |
| 2 | `compression` — uint16, the digest's δ (reference writer: 100); MUST be ≥ 1 |
| 4 | `centroid_count` — uint32 |
| 16 × C | `centroids` — per centroid: float64 `mean`, float64 `weight`. Means MUST be non-decreasing; weights MUST be positive, finite, and sum to `count`. |
| 8 | `min` — float64, exact minimum; meaningless when `count` == 0 |
| 8 | `max` — float64, exact maximum; meaningless when `count` == 0 |
| 8 | `count` — uint64, total values absorbed |

The payload length MUST equal `8 + 16 × centroid_count + 24` exactly.

**Value population.** A cell feeds the digest iff it is numeric under
the `STAT` interpretation (§4.1) **and** its float64 value is finite —
a non-finite value would poison every centroid mean, so writers MUST
skip NaN/±Inf here even though `STAT` folds them. Non-numeric cells are
invisible to the digest (they remain visible to `CHLL`).

**Header exclusion — the deliberate asymmetry with `STAT`.** The header
row (row 1, preamble-emitted) MUST NOT be folded into `TDIG` or `CHLL`
sketches. `STAT` folds the header because zone maps gate which rows a
query can see: a header the stats missed could cause a matching row to
be silently skipped, so pruning soundness forces over-inclusion. Sketch
sections gate nothing — they are estimates of the data distribution,
and every value folded in moves the estimate. Folding a header could
only *bias* quantiles and inflate distinct counts by a phantom value;
there is no soundness to buy. Estimation bias vs pruning soundness:
opposite failure modes, opposite rules.

**Quantile semantics.** `q = 0` and `q = 1` MUST return the exact
`min`/`max`. Interior quantiles interpolate linearly between centroid
means positioned at their cumulative-weight midpoints, anchored at
`(0, min)` and `(count, max)` — estimates therefore never leave
`[min, max]`. Note that interior estimation uses only arithmetic (no
libm), so decoding a committed sketch is bit-reproducible across
platforms even though *producing* one involves `asin`/`sin` in the
cluster-size bound and may differ in the last ulp across libm
implementations. Conformance for sketches is defined reader-side
(reproduce the vector goldens from the committed bytes), not
writer-side byte reproduction.

**Mergeability.** T-digests merge associatively: re-clustering the
union of two digests' centroids yields (statistically) the digest of
the concatenated streams. Combined with per-sheet records and the
future `WSTA`/append path, this is what lets segments or partitions of
a dataset carry their own sketches and be stitched without rereading
row data. Reference sizing: δ=100 holds ~60–200 centroids ≈ 1–4 KB per
(sheet × column).

### 4.4 Registered: `CHLL` — per-column HyperLogLog sketches (approximate distinct counts)

One HyperLogLog per tracked column per sheet, over the **canonical
string** of every non-empty cell value — text columns are first-class
here (that is half the point; a distinct count over IDs, emails, or
category labels needs no numeric interpretation). Same section frame as
`TDIG` (§4.3).

**HyperLogLog payload:**

| Size | Field |
|---|---|
| 2 | `format_version` — uint16, `1`. Readers MUST reject other values. |
| 1 | `p` — uint8, register-count exponent; m = 2^p registers (reference writer: p = 11, m = 2048, standard error ≈ ±2.3 %). Readers SHOULD bound p to a sane range (reference: 4–18). |
| m | `registers` — one byte per register, register j at payload offset 3 + j. Every register MUST be ≤ 64 − p + 1 (the maximum rank a 64-bit hash can produce). |

**Hash.** xxh64 of the canonical string; the top `p` bits of the
big-endian 64-bit digest select the register, the remaining 64 − p bits
feed the rank (position of the first 1-bit, 1-based; all-zero remainder
ranks 64 − p + 1).

**Canonicalization rule.** Distinctness is over canonical string forms,
not typed identity:

| Cell value | Canonical string |
|---|---|
| null, `''` (empty string) | **excluded** — an empty cell is the absence of a value |
| string | as-is, byte-exact (before XML escaping) |
| int / float | the decimal rendering the writer puts in the cell's `<v>` (PHP `(string) $value`) |
| bool | `'1'` / `'0'` |
| date/time | decimal rendering of its Excel serial number |
| other objects | their string conversion (the inlineStr rendering) |

Consequences worth stating: int `7` and string `'7'` collapse to one
value (both render `7`), while string `'1.50'` stays distinct from
float `1.5` (strings hash as-is even though the number cell renders
`1.5`). This mirrors "what a reader of the file would see" more closely
than PHP-level type identity would, and it is cheap to reproduce in any
implementation.

**Estimation.** The classic HLL estimator with the small-range (linear
counting) correction below `E ≤ 2.5m` — that is what keeps tiny
cardinalities near-exact. The classic large-range correction is
deliberately **omitted**: these sketches count values within one sheet,
bounded by its `total_rows` (uint32), which is astronomically below the
~2^57 regime where 64-bit hash-collision bias appears. Omitting it
keeps the estimator branch-free and exactly reproducible.

**Mergeability.** Two HLLs of equal `p` merge by register-wise max,
and the result is *exactly* the sketch a single pass over the union
stream would have produced — a stronger guarantee than the t-digest's
statistical mergeability, and the same stitch property for distributed
segments. The header row is excluded, as in `TDIG` (§4.3).

The reference writer emits both sections whenever
`withColumnSketches()` is enabled; either section is individually
optional for other writers (a quantile-only or distinct-only producer
is conforming).

### 4.5 Reserved tags

The following tags are reserved; their payloads are deliberately
unspecified in this document version. Writers MUST NOT emit them with
other meanings; readers encountering them before they are specified
simply skip them, as with any unknown tag.

| Tag | Intent |
|---|---|
| `WSTA` | Writer-state snapshot embedded in the file — makes a workbook appendable / an export resumable by carrying everything a writer needs to continue at the last sync point. |
| `RETR` | Retrofit seek data for files not born indexed: bit-aligned deflate offsets plus per-point inflate windows (zran-style). |
| `BLOM` | Per-block Bloom filters for membership pruning on non-range predicates. |
| `DICT` | Shared dictionary / string-table section supporting future shared-strings-aware indexing. |
| `SIGN` | Detached signature (Ed25519) over the sidecar, enabling block-granular tamper attribution together with `SCRC`. |
| `MRKL` | Merkle tree over per-block content hashes — verified partial reads (a ranged read proven against a signed root without fetching the rest of the file). |
| `PBMP` | Posting bitmaps for low-cardinality columns (value → blocks-containing bitmap; exact pruning, no false positives). |
| `TRGM` | Per-block trigram filters over opted-in text columns — substring-search pruning. |
| `VHIS` | Version-history chain for appendable files (per-append tail snapshot enabling in-file time travel). |
| `CSTR` | Columnar shadow stripes: raw numeric arrays for selected columns, block-aligned with `STAT`. |

Regardless of any section's presence, the primary worksheet entry stream
MUST remain decodable by a dictionary-unaware sequential inflater.
(Empirically verified: a deflate segment produced with a preset
dictionary after a full flush is NOT decodable by a plain inflater —
Excel-class consumers would reject the workbook. `DICT` payloads
therefore apply only to sidecar-internal data, never to the sheet
stream itself.)

## 5. The `flags` byte

Header byte 5 is a **must-understand bitmask**. All eight bits are
reserved in this document version.

- Writers MUST emit `0`.
- A reader that observes any bit it does not implement MUST reject the
  entire sidecar (and SHOULD fall back to non-indexed reading of the
  workbook).

This is the deliberate counterpart to TLV skipping: an unknown TLV
section is *ignorable by construction*, so it costs no version bump; a
flag bit says "the bytes you *do* parse don't mean what you think" —
for example a future alternate checksum or offset base. The flags byte
is the escape hatch that lets even such changes ship without burning
the version byte, at the price of locking out older readers only for
files that actually use them.

## 6. Normative reader invariants

A conforming reader MUST enforce all of the following before trusting
the index. Each check exists in the reference decoder; the listed
behavior on failure is "reject the sidecar" — a rejected sidecar makes
the workbook no worse than a plain XLSX (sequential reading MUST remain
available).

Structural (on the payload alone):

1. Payload is at least 16 bytes; `magic` == `KXSI`.
2. `version` == 2.
3. `flags` == 0 (§5).
4. CRC-32 of all bytes after the header equals `payload_crc32`.
5. `sync_period` > 0.
6. `sheet_count` within the implementation's sanity bound (§7).
7. Every length read stays inside the payload — path, sync-point
   array, TLV header, TLV payload, and every structure inside a TLV
   section. No read past the end, ever.
8. `entry_path_len` ≥ 1 and within the implementation bound (§7).
9. No duplicate `entry_path`.
10. Per sheet, sync points strictly increase in `row`, `comp_offset`,
    and `uncomp_offset`.
11. Per sheet, every `row` ≤ `total_rows + 1`.
12. `STAT`: per column, `column` ≥ 1 and
    `block_count` == `sync_count + 1`.
13. `SCRC`: per sheet, `count` == `sync_count`.
14. `TDIG`/`CHLL`: per column, `column` ≥ 1, every `payload_len` stays
    inside the section, and — before a sketch is *used* — its payload
    satisfies the internal invariants of §4.3/§4.4 (exact length,
    positive finite weights, ascending means, weight/count agreement;
    register values within the rank bound). The reference decoder
    validates framing at decode time and payload internals lazily at
    first access.

Container cross-checks (against the live ZIP central directory —
REQUIRED whenever the index will be used for seeking):

15. For every sheet record, the live central directory contains
    `entry_path` and its CRC-32 equals `sheet_crc32`; otherwise the
    whole index is stale (§2.2).
16. Every `comp_offset` is strictly less than the entry's live
    `compressed_size`; otherwise the index would seek the inflater into
    adjacent archive bytes.

Failing checks 15–16 SHOULD degrade silently to non-indexed reading
(the file itself is fine — only the sidecar is unusable); failing
checks 1–14 indicates a corrupt or crafted sidecar and MAY additionally
be surfaced as an error.

## 7. Security considerations

A KXSI sidecar is **untrusted input**, even inside a workbook you
wrote: archives get edited, proxied, and crafted. `payload_crc32`
detects accidental corruption only — it is trivially recomputable by an
attacker and MUST NOT be treated as a security mechanism (that is
`SIGN`'s future job).

The threat model is a sidecar with a *valid* CRC and hostile content.
Decoders defend in depth:

- **Allocation guards.** Every count field is validated against the
  remaining payload length *before* allocating (a `sync_count` of
  4 billion in a 200-byte body must die on the truncation check, not in
  the allocator). The reference decoder additionally caps `sheet_count`
  at 1024 (`MAX_SHEETS`) and `entry_path_len` at 256 (`MAX_PATH_LEN`);
  conforming decoders SHOULD impose comparable sanity bounds — real
  workbooks are nowhere near them.
- **Seek containment.** Offsets from the sidecar MUST never be used to
  read outside the sheet entry's data stream: cross-check against the
  live central directory (§6, checks 15–16) before the first seek.
  Without check 16, a crafted `comp_offset` walks the inflater into
  arbitrary archive bytes.
- **64-bit integers.** `row`/`comp_offset`/`uncomp_offset` are uint64.
  Implementations without native 64-bit integers MUST reject the
  sidecar loudly rather than silently truncate offsets into valid-
  looking garbage positions (the reference decoder refuses to run on
  32-bit PHP). Values above 2⁶³−1 do not occur in practice (ZIP32
  archives cap far lower) and MAY be rejected.
- **Monotonicity as a safety property.** Checks 10–11 are not
  pedantry: seek resolution walks the sync-point list assuming order;
  non-monotonic input would resolve a target row to the wrong offset
  and yield attacker-chosen rows while looking healthy.
- **Fail toward the slow path.** The worst acceptable outcome of any
  sidecar attack is a sequential scan. Readers MUST NOT let sidecar
  content change *which* bytes of the sheet are treated as row data
  beyond choosing a validated resume offset.

## 8. Test vectors — the conformance suite

The directory `tests/SpecVectors/` of the reference repository is the
normative conformance suite for this document. **Other implementations
(a TypeScript reader is planned) MUST validate against these files** —
matching them, not the prose above, is the practical definition of
interoperability; on any disagreement between prose and vectors, report
a spec bug.

Each vector is a triple:

| File | Meaning |
|---|---|
| `<name>.xlsx` | A real workbook produced by the reference writer |
| `<name>.expected.json` | Golden decode of its sidecar: `sync_period`, and per sheet `total_rows`, `sheet_crc32`, `sync_points`, `sync_point_crcs` (SCRC), `column_stats` (STAT), `column_sketches` (TDIG/CHLL — derived quantile and distinct-count estimates, pinning the estimators, not just the bytes) |
| `<name>.sidecar.hex` | Hexdump of the raw `xl/_kxs/index.bin` payload — pins the exact bytes |

The vectors:

| Vector | Exercises |
|---|---|
| `vector-01-plain-indexed` | Header + core body + `SCRC` only; 1 sheet, 250 data rows, `sync_period` 100 |
| `vector-02-stats` | `STAT` with mixed numeric / non-numeric cells (`count` vs `other`), unsorted |
| `vector-03-multisheet` | 2 sheets with different sync cadences — per-sheet record alignment in `STAT` and `SCRC` |
| `vector-04-sorted` | Sortedness flags: one ascending, one descending, one unsorted numeric column |
| `vector-05-sketches` | `TDIG` + `CHLL` on one mixed numeric column and one text column, no `STAT` (sections are orthogonal); golden pins quantile and distinct estimates computed from the committed sketch bytes |

A conforming **reader** must, for each vector: decode
`<name>.sidecar.hex` (or extract the part from the `.xlsx`) and
reproduce `<name>.expected.json` exactly; and verify the `SCRC`
semantics by inflating the sheet entry and checking
`crc32(bytes[0, uncomp_offset_k)) == sync_point_crcs[k]` for every k.

A conforming **writer** is checked the other way around: files it
produces must decode through a conforming reader with all §6 invariants
holding — and, for the reference writer itself, reproduce these exact
vector bytes.

The reference repository enforces all of this in
`tests/SpecVectors/SpecVectorsTest.php`, which **only reads the
committed fixtures** — nothing is regenerated at test time, so any
accidental format drift fails CI. Regeneration
(`php tests/SpecVectors/generate.php`) is a deliberate act that MUST
accompany a reviewed change to this document.
