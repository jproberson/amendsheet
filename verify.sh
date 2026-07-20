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

# Reports a failure if the pattern matches anywhere under src/.
forbid() {
  local pattern="$1" explanation="$2"
  local hits
  hits=$(grep -rnE "$pattern" src --include='*.ts' | grep -vE '^\S+:[0-9]+: *(\*|//)' || true)
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
npx biome lint src || fail "lint found problems"

step "typecheck"
npx tsc --noEmit || fail "typecheck failed"

step "house rules"
forbid '\bas [A-Z][A-Za-z]*\b|\bas unknown\b' 'type assertions are banned; validate at the boundary instead'
forbid ':\s*any\b|<any>' 'any is banned'
forbid '\w!\.' 'non-null assertions are banned'
forbid '^export default|export default ' 'default exports are banned'
forbid '\brequire\(' 'this package is ESM only'
if [ -d src/lib ]; then
  forbid 'console\.' 'no console in library code (scripts and harness may)'
  forbid "from 'node:|require\('node:" 'no Node-only APIs in the library core'
fi

step "tests"
if compgen -G 'src/**/*.test.ts' > /dev/null 2>&1 || find src -name '*.test.ts' -print -quit | grep -q .; then
  node --import tsx --test 'src/**/*.test.ts' || fail "tests failed"
else
  printf 'NO TESTS FOUND. Rule 10 says behaviour needs a test that failed first.\n'
  failures=$((failures + 1))
fi

step "harness regression"
if [ -d corpus ] && find corpus -name '*.xlsx' -print -quit | grep -q .; then
  npm run --silent harness | tail -n 15
else
  printf 'skipped: no corpus present (run npm run corpus && npm run corpus:real)\n'
fi

printf '\n=====================================\n'
if [ "$failures" -eq 0 ]; then
  printf 'VERIFY PASSED\n'
else
  printf 'VERIFY FAILED (%d problem areas)\n' "$failures"
fi
printf '=====================================\n'
exit "$failures"
