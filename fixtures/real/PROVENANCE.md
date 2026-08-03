# Real-world fixtures

Workbooks from the Apache POI project's test data, committed here so the
benchmark has fixed inputs. A round-trip result is only comparable across runs
if the files it ran against cannot change.

- Source: https://github.com/apache/poi/tree/0d6d4872c491b1f230f51c6878e57407c60ae697/test-data/spreadsheet
- Commit: `0d6d4872c491b1f230f51c6878e57407c60ae697`
- License: Apache-2.0. `LICENSE` and `NOTICE` are copied from `legal/` in that
  same commit and apply to these files.

The files are unmodified. POI collected them from a decade of bug reports, so
they were produced by real spreadsheet applications rather than by any
JavaScript library — which is the entire reason they are useful here. The
synthetic fixtures in `fixtures/quirks` pass almost entirely; these do not.

## Selection

`npm run fixtures:real` pulls every workbook under 600 KB from the pinned commit,
minus files named as a deliberately-broken input (a fuzzer minimisation, a
decompression bomb, a corrupt package) — those are error-path material, not
fidelity corpus, and one must not be fed to the parser at all. Each remaining
candidate is round-tripped and edited through the same measurement the harness
runs, and kept only if both passes are clean, so a file that would turn the
harness red never lands here by accident.

A handful are held out because they expose a real gap rather than a clean round
trip, and the script names them when it runs:

- **Strict OOXML** (`SimpleStrict`, `sample.strict`, `SampleSS.strict`, `57914`) —
  the ISO Strict variant uses a different relationship namespace, which the
  reader does not yet resolve. A genuine format to support; these join the corpus
  once it does.
- **Encrypted workbooks** (`58616`, `protected_passtika`) — an OLE2-wrapped,
  password-protected package. The reader rejects it with a located error telling
  the caller to decrypt first, which is correct, not a round trip.
- **No style table** (`47889`, `56278`, `59021`) — round-trip clean, but a file
  carrying no `styles.xml` cannot yet take a number format on edit.
- `49609` (no root `_rels/.rels`) and `sample-beta` (a 2007-beta shared-strings
  layout) are malformed or pre-release and are left out.
