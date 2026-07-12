# olecfsdk

Typed Rust SDK for Microsoft Office compound binary file formats.

The initial workspace contains:

- `crates/olecfsdk`: runtime, CFB container, and Office binary format models.
- `crates/olecfsdk-derive`: derive macros for symmetric binary read/write code.

The project is developed round-trip first: unknown bounded payloads are preserved while corpus-driven work replaces them with static Rust types.

## Phase 1 baseline

- CFB v3/v4 containers open into an owned logical storage/stream model and are
  rebuilt deterministically.
- Stream bytes, storage metadata, CLSIDs, state bits, and timestamps participate
  in logical round-trip comparison.
- Bounded readers, writers, allocation limits, and `SdkObject`/`SdkEnum` derives
  provide the base for typed DOC/XLS/PPT records.
- CFB reading and deterministic writing use SDK-owned static header,
  DIFAT/FAT, MiniFAT, directory, and regular/mini-stream types. The sibling
  `rust-cfb` crate is dev-only and provides strict differential validation.

The external corpus workspace contains 1533 generated legacy Office tests:

```sh
cd ../ooxmlsdk-test-suite
cargo test -p olecfsdk-roundtrip-tests --test apache_poi_cfb_roundtrip -- --ignored --quiet
cargo test -p olecfsdk-roundtrip-tests --test libreoffice_cfb_roundtrip -- --ignored --quiet
```
