#!/usr/bin/env bash
# Releases the version in mix.exs: tags it, waits for the release workflow
# to build the precompiled NIFs, commits their checksums and publishes to Hex.
#
# Run from a clean, up-to-date main. Safe to re-run after a failure: steps
# that are already done (tag pushed, checksums committed) are skipped.
set -euo pipefail

nif_module="Spreadsheet.Calamine"
checksum_file="checksum-Elixir.Spreadsheet.Calamine.exs"
commit_message="release: update checksums"

cd "$(dirname "$0")/.."

version=$(sed -n 's/^  @version "\(.*\)"$/\1/p' mix.exs)
tag="v$version"
[ -n "$version" ] || { echo "could not read @version from mix.exs" >&2; exit 1; }

echo "==> Releasing $tag"

[ "$(git branch --show-current)" = "main" ] || { echo "not on main" >&2; exit 1; }
[ -z "$(git status --porcelain)" ] || { echo "working tree is not clean" >&2; exit 1; }
git pull --ff-only origin main

if git ls-remote --exit-code --tags origin "refs/tags/$tag" >/dev/null; then
  echo "==> $tag already pushed"
else
  git tag "$tag"
  git push origin "$tag"
fi

echo "==> Waiting for the release workflow on $tag"
run_id=""
for _ in $(seq 1 30); do
  run_id=$(gh run list --workflow release.yml --branch "$tag" --limit 1 --json databaseId -q '.[0].databaseId')
  [ -n "$run_id" ] && break
  sleep 5
done
[ -n "$run_id" ] || { echo "no release workflow run found for $tag" >&2; exit 1; }
gh run watch "$run_id" --exit-status

echo "==> Downloading checksums"
mix rustler_precompiled.download "$nif_module" --all --print

if grep -q -- '-v[0-9]' "$checksum_file" && grep -- '-v[0-9]' "$checksum_file" | grep -qv -- "-$tag-"; then
  echo "$checksum_file still lists other versions" >&2
  exit 1
fi
echo "==> $(grep -c -- "-$tag-" "$checksum_file") artifacts in $checksum_file"

if [ -n "$(git status --porcelain -- "$checksum_file")" ]; then
  git commit -m "$commit_message" -- "$checksum_file"
  git push origin main
fi

mix hex.publish
