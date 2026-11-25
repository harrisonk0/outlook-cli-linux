#!/usr/bin/env bash
set -euo pipefail

release=v0.1.1
repository=https://github.com/harrisonk0/outlook-cli-linux.git
directory=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
if [[ ! -f "$directory/outlook.cpp" && ! -f "$directory/outlook" ]]; then
  directory="${XDG_DATA_HOME:-$HOME/.local/share}/outlook-cli"
  mkdir -p "$(dirname -- "$directory")"
  if [[ ! -d "$directory/.git" ]]; then
    git clone --branch "$release" --depth 1 "$repository" "$directory"
  fi
  exec bash "$directory/setup.sh" "$@"
fi

usage() {
  echo 'Usage: setup.sh [--install] [--prefix PATH] [--runner PATH]'
}
fail() { echo "$*" >&2; exit 1; }
prefix=${WINEPREFIX:-}
runner=${OUTLOOK_WINE:-}
install_office=false
while (($#)); do
  case "$1" in
    --help) usage; exit 0 ;;
    --install) install_office=true; shift ;;
    --prefix|--runner)
      (($# >= 2)) || fail "Missing value for $1"
      if [[ $1 == --prefix ]]; then prefix=$2; else runner=$2; fi
      shift 2 ;;
    *) usage >&2; exit 2 ;;
  esac
done

data=${XDG_DATA_HOME:-$HOME/.local/share}
roots=("$data/bottles" "$HOME/.var/app/com.usebottles.bottles/data/bottles")
shopt -s nullglob
find_outlook() {
  local candidate
  for candidate in "$1"/drive_c/Program\ Files*/Microsoft\ Office/{root/,}Office*/[Oo][Uu][Tt][Ll][Oo][Oo][Kk].[Ee][Xx][Ee]; do
    [[ -f "$candidate" ]] && { printf '%s\n' "$candidate"; return 0; }
  done
  return 1
}
discover() {
  local root candidate
  choices=()
  for candidate in "$HOME/.wine" "${roots[@]/%//bottles}"; do
    if [[ "$candidate" == "$HOME/.wine" ]]; then
      if find_outlook "$candidate" >/dev/null; then choices+=("$candidate"); fi
    else
      for root in "$candidate"/*; do
        if find_outlook "$root" >/dev/null; then choices+=("$root"); fi
      done
    fi
  done
}

open_bottles() {
  if command -v bottles >/dev/null; then
    bottles </dev/null >/dev/null 2>&1 &
  elif command -v flatpak >/dev/null; then
    if ! flatpak info com.usebottles.bottles >/dev/null 2>&1; then
      flatpak remote-add --user --if-not-exists flathub https://flathub.org/repo/flathub.flatpakrepo
      flatpak install --user flathub com.usebottles.bottles
    fi
    flatpak run com.usebottles.bottles </dev/null >/dev/null 2>&1 &
  else
    fail 'Install Flatpak or Bottles first. See https://usebottles.com/download/'
  fi
}
discover
if [[ -z "$prefix" ]] && [[ "$install_office" == true || ${#choices[@]} == 0 ]]; then
  [[ -t 0 && -n "${DISPLAY:-}${WAYLAND_DISPLAY:-}" ]] ||
    fail 'No Outlook bottle selected. On a desktop run setup.sh --install, or pass --prefix and --runner.'
  open_bottles
  printf '%s\n' 'In Bottles, enable pre-release runners and select a compatible Soda runner.' \
    'Create a Software bottle and install Microsoft 365 from Programs.' \
    'Open Classic Outlook once and finish licensing and account setup.' \
    'Office setup notes: https://github.com/bottlesdevs/programs/issues/500'
  read -r -p 'Press Enter when Outlook is installed. ' _
  discover
fi
if [[ -z "$prefix" ]]; then
  case ${#choices[@]} in
    0) fail 'No Classic Outlook installation found. Rerun setup after installing it, or pass --prefix.' ;;
    1) prefix=${choices[0]} ;;
    *) [[ -t 0 ]] || fail 'Multiple Outlook bottles found. Pass --prefix.'
       PS3='Select your Outlook bottle: '
       select prefix in "${choices[@]}"; do [[ -n "$prefix" ]] && break; done ;;
  esac
fi
prefix=$(cd -- "$prefix" && pwd)
office=${OUTLOOK_EXE:-}
[[ -n "$office" ]] || office=$(find_outlook "$prefix") || fail 'Classic Outlook was not found in that prefix.'
[[ -f "$office" ]] || fail 'OUTLOOK_EXE does not point to a file.'

if [[ -z "$runner" && -f "$prefix/bottle.yml" ]]; then
  runner_name=$(sed -nE "s/^Runner:[[:space:]]*['\"]?([[:alnum:]_.-]+)['\"]?[[:space:]]*$/\1/p" "$prefix/bottle.yml" | head -n 1)
  for root in "${roots[@]}" "$(dirname -- "$(dirname -- "$prefix")")"; do
    for name in wine wine64; do
      candidate="$root/runners/$runner_name/bin/$name"
      if [[ -n "$runner_name" && -x "$candidate" ]]; then runner=$candidate; break 2; fi
    done
  done
  [[ -n "$runner" ]] || fail 'The bottle runner was not found. Pass --runner with its Wine executable or wrapper.'
fi
runner=$(command -v -- "${runner:-wine}") || fail 'Wine was not found. Pass --runner with the runner path.'
runner=$(cd -- "$(dirname -- "$runner")" && printf '%s/%s\n' "$PWD" "$(basename -- "$runner")")
arch=${OUTLOOK_ARCH:-}
if [[ -z "$arch" ]]; then
  case "$(file -b -- "$office")" in
    PE32+*x86-64*) arch=x64 ;;
    PE32*Intel\ 80386*) arch=x86 ;;
    *) fail 'Could not determine Outlook architecture. Set OUTLOOK_ARCH=x64 or x86.' ;;
  esac
fi
case "$arch" in x64|x86) ;; *) fail 'OUTLOOK_ARCH must be x64 or x86.' ;; esac

if [[ ! -x "$directory/outlook-$arch.exe" ]]; then
  archive="outlook-cli-linux-$release"
  downloads="https://raw.githubusercontent.com/harrisonk0/outlook-cli-linux/downloads/$release"
  staging=$(mktemp -d)
  trap 'rm -rf -- "$staging"' EXIT
  for asset in "$archive.tar.gz" SHA256SUMS; do
    curl -fLSs --retry 2 "$downloads/$asset" -o "$staging/$asset"
  done
  (cd "$staging" && sha256sum -c SHA256SUMS && tar -xzf "$archive.tar.gz")
  install -m 755 "$staging/$archive"/outlook-{x64,x86}.exe "$directory/"
fi

printf 'Checking Outlook with %s (%s)\n' "$runner" "$arch"
if ! WINEPREFIX="$prefix" OUTLOOK_WINE="$runner" OUTLOOK_ARCH="$arch" "$directory/outlook" version; then
  fail 'Outlook did not respond. Open it once in Bottles and finish setup, or pass a runner wrapper that works outside Bottles.'
fi
config=${XDG_CONFIG_HOME:-$HOME/.config}/outlook-cli
mkdir -p "$config"
temporary=$(mktemp "$config/bottle.XXXXXX")
printf '%s\0' "$prefix" "$runner" "$arch" "$office" > "$temporary"
mv -f -- "$temporary" "$config/bottle"
printf 'Saved Outlook setup. Run: %q mapi-stores\n' "$directory/outlook"
