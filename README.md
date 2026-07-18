# olecfsdk

Typed Rust SDK for Microsoft Office compound binary file formats.

The initial workspace contains:

- `crates/olecfsdk`: runtime, CFB container, and Office binary format models.
- `crates/olecfsdk-derive`: derive macros for symmetric binary read/write code.

The project is developed round-trip first: unknown bounded payloads are preserved while corpus-driven work replaces them with static Rust types.

## Typed file roots

`doc::DocFile`, `ppt::PptFile`, and `xls::XlsFile` open a compound file into
owned Rust structure trees and rebuild their managed streams from those trees.
They deliberately preserve physical and semantic hierarchy: DOC text pieces
retain CP/FC and encoding, PPT remains a recursive record/container tree, and
XLS retains BIFF records plus nested BOF/EOF substreams (including files that
contain both `/Workbook` and `/Book`). The SDK does not substitute lossy
plain-text, slide-summary, or cell-string projections for these trees.

The parse-time CFB is an immutable preservation snapshot. Typed edits update
the Rust tree; `to_compound_file`, `to_bytes`, and `save` rebuild every managed
stream from that tree while carrying unrelated CFB entries forward. Default
open/save operations are strict. Producer deviations must be opened through a
`*_compatible` entry point, inspected through its structured diagnostics, and
saved with `SaveOptions::preserving_compatibility()` only when retaining those
explicit compatibility nodes is intentional.

Runnable examples perform a semantic edit and strict reopen of the result:

```sh
cargo run -p olecfsdk --example edit_doc -- input.doc output.doc
cargo run -p olecfsdk --example edit_xls -- input.xls output.xls
cargo run -p olecfsdk --example edit_ppt -- input.ppt output.ppt
```

PPT's ordinary save preserves and relocates its existing physical incremental
history. Call the separate `PptHistoryStrategy` APIs only when an append or a
normalized live-state rebuild is explicitly required; history policy is not a
parse or compatibility option.

## CFB baseline

- CFB v3/v4 containers open into an owned logical storage/stream model and are
  rebuilt deterministically.
- Stream bytes, storage metadata, CLSIDs, state bits, and timestamps participate
  in logical round-trip comparison.
- Bounded readers, writers, allocation limits, and `SdkObject`/`SdkEnum`/`SdkBitfield` derives
  provide the base for typed DOC/XLS/PPT records.
- CFB reading and deterministic writing use SDK-owned static header,
  DIFAT/FAT, MiniFAT, directory, and regular/mini-stream types. Strict reopen,
  name ordering, allocation validation, stream/storage editing, and corpus
  assertions are all provided by `olecfsdk`; neither this workspace nor the
  external test suite depends on the sibling `rust-cfb` crate.

The external corpus workspace contains 1,533 generated legacy Office CFB tests:

```sh
cd ../ooxmlsdk-test-suite
cargo test -p olecfsdk-roundtrip-tests --test apache_poi_cfb_roundtrip -- --ignored --quiet
cargo test -p olecfsdk-roundtrip-tests --test libreoffice_cfb_roundtrip -- --ignored --quiet
```
