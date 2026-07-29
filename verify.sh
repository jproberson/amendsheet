#!/usr/bin/env bash
# Run before every commit. Fails on the first problem so the output stays readable.
set -uo pipefail

failures=0

step() {
  printf '\n%s\n' "----- $1 -----"
}

fail() {
  printf 'FAIL: %s\n' "$1"
  failures=$((failures + 1))
}

# forbid <pattern> <explanation> [scope] [--no-tests] [ignore-pattern]
# Reports a failure if the pattern matches. Comment lines are ignored so that
# documenting a banned construct does not trip the check.
forbid() {
  local pattern="$1" explanation="$2" scope="${3:-src}" tests="${4:-}" ignore="${5:-}"
  local hits
  if [ "$tests" = "--no-tests" ]; then
    hits=$(grep -rnE "$pattern" "$scope" --include='*.ts' --exclude='*.test.ts' || true)
  else
    hits=$(grep -rnE "$pattern" "$scope" --include='*.ts' || true)
  fi
  hits=$(printf '%s' "$hits" | grep -vE '^\S+:[0-9]+: *(\*|//)' || true)
  if [ -n "$ignore" ]; then
    hits=$(printf '%s' "$hits" | grep -vE "$ignore" || true)
  fi
  if [ -n "$hits" ]; then
    fail "$explanation"
    printf '%s\n' "$hits" | sed 's/^/    /'
  fi
}

step "format"
if ! npx biome format --write src; then
  fail "formatter errored"
fi

step "lint"
npx biome lint --error-on-warnings src || fail "lint found problems"

step "typecheck"
npx tsc --noEmit || fail "typecheck failed"

step "house rules"
# `import { X as Y }` is an alias, not an assertion, so import lines are exempt.
forbid '\bas [A-Z][A-Za-z]*\b|\bas unknown\b' 'type assertions are banned; validate at the boundary instead' src '' '^\S+:[0-9]+: *import '
forbid ':\s*any\b|<any>' 'any is banned'
forbid '\w!\.' 'non-null assertions are banned'
forbid '^export default|export default ' 'default exports are banned'
forbid '\brequire\(' 'this package is ESM only'
if [ -d src/lib ]; then
  forbid 'console\.' 'no console in library code (scripts and harness may)' src/lib --no-tests
  forbid "from 'node:|require\('node:" 'no Node-only APIs in the library core' src/lib --no-tests
fi

step "docs"
# A doubled word from an edit — "the file is is at fault". Only words that are
# never validly repeated, so there is nothing to whitelist. CHANGELOG.md is a
# working note, not gated.
doc_dupes=$(grep -rniE '\b(is is|the the|a a|an an|and and|or or|of of|to to|in in|on on|for for|with with|as as)\b' \
  --include='*.md' --include='llms.txt' --exclude-dir=node_modules --exclude=CHANGELOG.md . 2>/dev/null || true)
if [ -n "$doc_dupes" ]; then
  fail "a word is doubled in the docs"
  printf '%s\n' "$doc_dupes" | sed 's/^/    /'
fi

# Every ts example in README.md and llms.txt is typechecked against the real
# exports, so a renamed or removed part of the API breaks the docs rather than
# drifting from them.
examples_output=$(node scripts/doc-examples.mjs 2>&1)
printf '%s\n' "$examples_output" | head -1
printf '%s' "$examples_output" | grep -qE '^(PASSED|SKIPPED)' || {
  fail "a doc example is out of step with the exports"
  printf '%s\n' "$examples_output" | tail -n +2 | sed 's/^/    /'
}

step "tests and coverage"
# Thresholds ratchet up, never down, with one exception recorded here.
#
# Lines sit at 99 rather than 100 because Node measures coverage through tsx's
# transpiled output and attributes a blank source line as unexecuted. Confirmed
# by lcov: xml.ts line 28 reports DA:28,0 and contains only a newline. Which
# blank line is blamed moves as files change, so 100 is not reachable. A real
# gap costs far more than half a percent, so the gate still bites.
#
# Branches are a FIXED BACKSTOP at 97, not a ratchet — because this metric is too
# noisy to ratchet. The count is measured through tsx's transpiled output and
# misattributes branch hits: lcov shows arms that ARE exercised reported with zero
# hits (patch.ts readSheetView's self-closing arm, hyperlinks.ts's append, the
# alignment arm of resolveStyle tested six times). The proof it is noise, not gaps:
# adding page.ts — a small module with zero untested lines — moved the aggregate
# 98.02 -> 97.99. Fully-tested code cannot lower real coverage, so the wobble is the
# measurement, and ratcheting a number that drifts down from noise would eventually
# make the gate unmeetable on its own. 97 gives headroom for that drift while still
# catching a gross regression. The rigor lives in functions=100, lines=99, and the
# 0-lossy harness, which are reliable and responsive; branch is a floor, nothing more.
if find src -name '*.test.ts' -print -quit | grep -q .; then
  node --import tsx --test \
    --experimental-test-coverage \
    --test-coverage-include='src/lib/**' \
    --test-coverage-exclude='**/*.test.ts' \
    --test-coverage-functions=100 \
    --test-coverage-lines=99 \
    --test-coverage-branches=97 \
    --test-reporter=spec \
    'src/**/*.test.ts' || fail "tests or coverage thresholds failed"
else
  printf 'NO TESTS FOUND. Rule 10 says behaviour needs a test that failed first.\n'
  failures=$((failures + 1))
fi

step "package"
npm run --silent build > /dev/null || fail "build failed"
npx publint --strict > /dev/null 2>&1 || { fail "publint found packaging problems"; npx publint --strict 2>&1 | tail -8; }
npx attw --pack . --format table-flipped > /dev/null 2>&1 || { fail "type resolution is broken for some consumers"; npx attw --pack . 2>&1 | tail -12; }

step "package integrity"
# Excel offers to repair a file with a dangling relationship, a part with no
# content type, or a table whose shape is inconsistent. No Excel here checks
# that, so this does, over every fixture edited two ways. Only a problem our
# edit introduces, not one the source file already had, fails.
if [ -d fixtures ] && find fixtures -name '*.xls[xm]' -print -quit | grep -q .; then
  validate_output=$(node --import tsx scripts/validate-opc.mjs 2>&1)
  printf '%s\n' "$validate_output" | tail -1
  printf '%s' "$validate_output" | grep -q '^FAIL' && { fail "package integrity"; printf '%s\n' "$validate_output" | grep '^FAIL\|^     ' | head -20; }
else
  printf 'skipped: no fixtures present\n'
fi

step "browser"
# We claim browser support, and the grep above only proves no Node API is named.
# This runs the built bundle in a real browser end to end. It skips cleanly on a
# machine with no Chrome, so it never fails the build for being unable to run.
browser_output=$(node scripts/browser-smoke.mjs 2>&1)
printf '%s\n' "$browser_output"
printf '%s' "$browser_output" | grep -q '^SKIPPED' || {
  printf '%s' "$browser_output" | grep -q '^PASSED' || fail "browser smoke test failed"
}

step "dates in other timezones"
# date.ts converts through calendar components to survive daylight saving, and
# that only means anything if it is run somewhere the clock actually moves.
for zone in America/Santiago Asia/Beirut Pacific/Auckland; do
  TZ="$zone" node --import tsx --test 'src/lib/date.test.ts' > /dev/null 2>&1 \
    || fail "date handling is wrong in $zone"
done

step "harness regression"
harness_problems=$(mktemp)
trap 'rm -f "$harness_problems"' EXIT
if [ -d fixtures ] && find fixtures -name '*.xlsx' -print -quit | grep -q .; then
  harness_output=$(npm run --silent harness)
  printf '%s\n' "$harness_output" | grep -E 'ROUND-TRIP|PARTS REWRITTEN|SUMMARY'

  # Only this library's blocks gate the build, and both of them do: reading the
  # comparison adapter's numbers instead is why this step went blind once.
  # Each block yields "rewritten lossy failed"; all three must be zero.
  ours=$(printf '%s\n' "$harness_output" |
    awk '/^ROUND-TRIP FIDELITY:/ {mine = ($0 ~ /amendsheet/)}
         mine && /^PARTS REWRITTEN/ {rewritten = $3}
         mine && /^SUMMARY/ {print rewritten, $5, $8}')

  if [ -z "$ours" ]; then
    fail "harness printed no summary for amendsheet"
  else
    printf '%s\n' "$ours" | while read -r rewritten lost unread; do
      [ "$lost" = "0" ] || printf 'FAIL: harness: %s file(s) lost something\n' "$lost"
      [ "$unread" = "0" ] || printf 'FAIL: harness: %s file(s) could not be processed\n' "$unread"
      [ "$rewritten" = "0" ] || printf 'FAIL: harness: %s file(s) had a part rewritten\n' "$rewritten"
    done > "$harness_problems"
    if [ -s "$harness_problems" ]; then
      cat "$harness_problems"
      fail "harness regression"
    fi
  fi
else
  printf 'skipped: no fixtures present (run npm run fixtures && npm run fixtures:real)\n'
fi

printf '\n=====================================\n'
if [ "$failures" -eq 0 ]; then
  printf 'VERIFY PASSED\n'
else
  printf 'VERIFY FAILED (%d problem areas)\n' "$failures"
fi
printf '=====================================\n'
exit "$failures"
