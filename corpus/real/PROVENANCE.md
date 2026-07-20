# Real-world corpus

Sixty `.xlsx` files from the Apache POI project's test data, committed here so
the benchmark has fixed inputs. A round-trip result is only comparable across
runs if the files it ran against cannot change.

- Source: https://github.com/apache/poi/tree/0d6d4872c491b1f230f51c6878e57407c60ae697/test-data/spreadsheet
- Commit: `0d6d4872c491b1f230f51c6878e57407c60ae697`
- License: Apache-2.0. `LICENSE` and `NOTICE` are copied from `legal/` in that
  same commit and apply to these files.

The files are unmodified. POI collected them from a decade of bug reports, so
they were produced by real spreadsheet applications rather than by any
JavaScript library — which is the entire reason they are useful here. The
synthetic corpus in `corpus/quirks` passes almost entirely; these do not.

Run `npm run corpus:real` to pull more from the same pinned commit.
