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

step "tests and coverage"
# Thresholds ratchet up, never down, with one exception recorded here.
#
# Lines sit at 99 rather than 100 because Node measures coverage through tsx's
# transpiled output and attributes a blank source line as unexecuted. Confirmed
# by lcov: xml.ts line 28 reports DA:28,0 and contains only a newline. Which
# blank line is blamed moves as files change, so 100 is not reachable. A real
# gap costs far more than half a percent, so the gate still bites.
#
# Branches sit below 100 because a few are unreachable by construction rather
# than untested; raise this as those get removed or covered.
if find src -name '*.test.ts' -print -quit | grep -q .; then
  node --import tsx --test \
    --experimental-test-coverage \
    --test-coverage-include='src/lib/**' \
    --test-coverage-exclude='**/*.test.ts' \
    --test-coverage-functions=100 \
    --test-coverage-lines=99 \
    --test-coverage-branches=98 \
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

step "harness regression"
if [ -d fixtures ] && find fixtures -name '*.xlsx' -print -quit | grep -q .; then
  npm run --silent harness | tail -n 15
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
