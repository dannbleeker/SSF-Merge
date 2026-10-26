#!/bin/bash
# Brings a Claude Code cloud container up to what this repository asks for.
#
# The container's image can lag the repository: on 2026-10-26 it shipped Node
# 22.22.0 against an `engines` floor of 22.22.2, so `npm ci` refused, and
# Chromium 1243 while the locked playwright wanted 1248, so the pane sweep
# could not launch its own browser. Both were fixed by hand and both would come
# back with the next session. This hook fixes them on every start instead.
#
# No version is written here. The Node major comes from `.nvmrc` and the floor
# from `package.json`'s `engines`, so a bump to either needs no edit to this
# file. Each step is skipped when it is already satisfied.
set -euo pipefail

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "$CLAUDE_PROJECT_DIR"

# npm's own copy of semver, which every Node install carries.
SEMVER="$(npm root -g)/npm/node_modules/semver"
RANGE="$(node -p 'require("./package.json").engines.node')"

satisfied() {
  node -e 'process.exit(require(process.argv[1]).satisfies(process.version, process.argv[2]) ? 0 : 1)' "$SEMVER" "$RANGE"
}

if satisfied; then
  echo "session-start: node $(node -v) meets $RANGE"
else
  MAJOR="$(tr -d ' v\n' < .nvmrc)"
  V="$(curl -sSfL https://nodejs.org/dist/index.json |
    node -e 'let s="";process.stdin.on("data",d=>s+=d).on("end",()=>{const m=process.argv[1];console.log(JSON.parse(s).find(r=>r.version.startsWith(`v${m}.`)).version)})' "$MAJOR")"
  echo "session-start: node $(node -v) does not meet $RANGE; installing $V"
  # Unpacked over the install already on PATH, so every global package in it
  # (the Claude CLI among them) stays where it was.
  PREFIX="$(dirname "$(dirname "$(command -v node)")")"
  TMP="$(mktemp -d)"
  TARBALL="node-$V-linux-x64.tar.xz"
  curl -sSfL -o "$TMP/$TARBALL" "https://nodejs.org/dist/$V/$TARBALL"
  (cd "$TMP" && curl -sSfL "https://nodejs.org/dist/$V/SHASUMS256.txt" | grep " $TARBALL\$" | sha256sum -c -)
  tar -xJf "$TMP/$TARBALL" -C "$PREFIX" --strip-components=1
  rm -rf "$TMP"
  hash -r
  satisfied || { echo "session-start: $V still does not meet $RANGE" >&2; exit 1; }
fi

# `ci`, not `install`. npm 10's `install` rewrites the lockfile, stripping every
# `libc` field (docs/DEPENDENCY-ALERTS.md, 2026-10-05), so each session would
# start with a modified `package-lock.json` for a `git add -A` to sweep up.
# `ci` never writes the lockfile and installs exactly what CI installs.
npm ci --no-audit --no-fund

# Fetches the Chromium the LOCKED playwright expects, and does nothing when it
# is already there. scripts/launch-chromium.mjs falls back to an older build
# when this is missing; with it present, no fallback is needed.
npx playwright install chromium
