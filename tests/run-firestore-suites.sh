#!/usr/bin/env bash
# Run from backend/functions-node under firebase emulators:exec.
set -eu
[[ "${FIRESTORE_EMULATOR_HOST:-}" == "127.0.0.1:8187" ]]
documents='http://127.0.0.1:8187/emulator/v1/projects/demo-geapa-functions-local/databases/(default)/documents'

status=0
for suite in test/enrollment-lifecycle.test.cjs test/*-emulator.test.cjs test/import-recognition-executor.test.cjs; do
  # Some fixtures use the default database and enumerate whole collections.
  curl --noproxy '*' --fail --silent --show-error --output /dev/null --request DELETE "$documents"
  node --test "$suite" || status=1
done
exit "$status"
