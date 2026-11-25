# Outlook CLI for Wine

A Linux command-line bridge to Classic Outlook. The Windows executable invokes
Outlook's COM object model through OLE Automation and calls Outlook's Extended
MAPI provider directly. It uses the existing Outlook installation, license,
profile, and data stores in the selected Wine prefix.

## Install

1. Install Classic Outlook for Windows in a Wine prefix or a Bottles Software
   bottle using a compatible Soda runner. Microsoft supplies Outlook separately.
   Use a licensed installation and create a profile. See the
   [Bottles Office setup notes](https://github.com/bottlesdevs/programs/issues/500).
2. Download `outlook-cli-linux-v0.1.0.tar.gz` and `SHA256SUMS` from
   [GitHub Releases](https://github.com/harrisonk0/outlook-cli/releases).
   On Linux, run these commands in the download directory:

   ```sh
   sha256sum -c SHA256SUMS
   tar -xzf outlook-cli-linux-v0.1.0.tar.gz
   cd outlook-cli-linux-v0.1.0
   ```

3. Configure the launcher as shown below. Linux needs Bash and GNU Coreutils.
   Install Xvfb and xauth if you will run without a desktop display. For an
   offline test profile, start Classic Outlook once with `/PIM OutlookCliTest`.

## Connect to a bottle

Use the bottle's prefix and matching runner. Set `OUTLOOK_ARCH=x86` for 32-bit
Outlook. The CLI defaults to x64. MAPI requires the same architecture as Outlook.

```sh
export WINEPREFIX='/path/to/bottle'
export OUTLOOK_WINE='/path/to/runner/bin/wine'
export OUTLOOK_ARCH=x64
./outlook com-check
./outlook version
./outlook mapi-check
./outlook mapi-stores
```

`OUTLOOK_WINE` accepts an executable path. Use a wrapper at that path if your
runner requires extra environment setup. For a standard Click-to-Run install,
the CLI looks for `OLMAPI32.DLL` under Office16 in Program Files. For another
installation path, set the absolute Windows path, for example
`OUTLOOK_MAPI_DLL='C:\Program Files\Microsoft Office\Office16\OLMAPI32.DLL'`.

If no display is set, the launcher uses `xvfb-run`. For repeated operations
against one Outlook session on a headless machine, keep one virtual display
alive for the entire session, for example with `xvfb-run -a bash` and then invoke
the CLI from that shell. Set `OUTLOOK_TIMEOUT` to a GNU timeout duration to
override the default of 60 seconds. A timeout during a write leaves its outcome
unknown. Check the relevant Outlook folder before retrying.

## Commands

| Command | Behavior |
| --- | --- |
| `batch COMMAND [-- COMMAND …]` | Runs multiple COM commands in order through one Outlook connection. |
| `com-check` | Exercises Wine OLE Automation using Scripting.Dictionary. |
| `com STEP [...]` | Runs a sequence of Outlook COM calls and property reads. |
| `version` | Activates Outlook.Application and prints its version. |
| `folders` | Prints the root folders exposed by Outlook's COM session. |
| `inbox` | Prints up to 50 default Inbox subjects, newest first, using COM. |
| `sync [PROFILE]` | Logs on to the Outlook profile and starts asynchronous Send/Receive. |
| `draft TO SUBJECT BODY` | Saves a mail item to Drafts and prints its entry ID. |
| `send TO SUBJECT BODY` | Resolves recipients and submits an Outlook mail item using COM. Outlook can ask for confirmation. |
| `mapi-check` | Loads and initializes Outlook's Extended MAPI provider. |
| `mapi-stores [PROFILE]` | Opens a MAPI session without requesting logon UI and lists its stores. |
| `mapi-inbox [PROFILE]` | Prints up to 50 Inbox entries, newest first, using Extended MAPI. |
| `mapi-drafts [PROFILE]` | Prints up to 50 Drafts entries, most recently modified first. |
| `mapi-draft SUBJECT BODY [PROFILE]` | Saves an unsent message to Drafts using Extended MAPI and prints its entry ID in hexadecimal. |
| `mapi-read ENTRYID [PROFILE]` | Prints the plain-text body for an entry ID from the default store. |
| `mapi-send TO SUBJECT BODY [PROFILE]` | Submits a message with one SMTP recipient to Outlook via Extended MAPI. |

For example, count the items in the default Calendar folder:

```sh
./outlook com get:Session call:GetDefaultFolder i:9 get:Items get:Count
```

Use `get:NAME` for a property, `call:NAME` for a method, or `put:NAME`
to set a property. Follow each step with its arguments in Outlook's
documented order. The supported argument types are `s:TEXT`, `i:INTEGER`,
`true`, `false`, and `missing` for an omitted optional argument. An intermediate
step can return an object to use for the next step. Property writes and methods
with no return value keep the current object. The final step can print a scalar.

Use a batch for several operations in one Outlook COM connection:

```sh
./outlook batch version -- folders -- inbox -- draft '' 'Offline test' 'A local draft'
```

The batch stops on the first error. Earlier operations can already have
taken effect. An Inbox or Drafts listing from Extended MAPI prints the
hexadecimal entry ID followed by a tab and the subject. Pass that ID to
`mapi-read` with the same profile. Output text is UTF-8. Quote multiword arguments.

To send without Outlook Object Model Guard dialogs, submit through Extended
MAPI, then start Send/Receive on the same configured profile:

```sh
./outlook mapi-send 'recipient@example.com' 'Subject' 'Plain-text body' MyProfile
./outlook sync MyProfile
```

`mapi-send` reports submission to Outlook. The transport delivers the message
during synchronization. Keep the virtual display open while synchronization
runs. If Outlook is already running with another profile, close it before
switching profiles.

## Verification status

Both Windows architectures built with LLVM-MinGW. The x64 OLE Automation
check passed under Soda 11.0-38 experimental and Wine 11.22 WoW64. Using a
local PST on Soda and Classic Outlook build 16.0.20625.20142, the COM batch
returned the Outlook version, listed folders, queried the Inbox and Calendar,
and saved a draft. A separate Soda prefix passed the COM and MAPI probes
with Classic Outlook build 16.0.20326.20144. The native Extended MAPI provider opened
the same PST, found the COM draft, and saved and read back its own Unicode
draft with the original body. On the newer build, Extended MAPI submission
and COM Send/Receive delivered a message to a loopback SMTP server. The
server received the expected sender, recipient, subject, and body. Outlook
also retrieved a seeded message from a loopback POP3 server. The CLI found
it in the Inbox through Extended MAPI and matched the original body.

Standalone repeated COM activation has been intermittent with this Outlook
and Soda combination. In a dedicated test bottle, two batches passed when
each ended with `-- com call:Quit` and the next started four seconds later.
`Quit` closes Outlook in that bottle. The delivery test used a local POP/SMTP
profile. Exchange and Microsoft 365 mailbox sign-in and 32-bit Outlook
execution were outside the verified setup.

- [Bottles Microsoft 365 installer](https://github.com/bottlesdevs/programs/blob/main/Software/microsoft365enterprise.yml)
- [Bottles Office compatibility and Outlook status](https://github.com/bottlesdevs/programs/issues/500)
- [Microsoft Office Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
- [Microsoft Outlook API selection guide](https://learn.microsoft.com/en-us/office/client-developer/outlook/selecting-an-api-or-technology-for-developing-solutions-for-outlook)
- [Microsoft MAPI provider loading reference](https://github.com/microsoft/MAPIStubLibrary)
- [Microsoft Extended MAPI store examples](https://github.com/microsoft/mfcmapi)
- [Microsoft guide to creating a local Outlook profile](https://support.microsoft.com/en-us/outlook/use-outlook-without-an-email-account)
- [Microsoft Extended MAPI message submission reference](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/imessage-submitmessage)

## Microsoft 365 and Exchange sign-in

Run `./outlook login` from a visible desktop using the same prefix and runner
as the CLI. Close Outlook before opening the profile picker. Create a dedicated
profile, add the mailbox through Outlook's automatic account setup, and complete
Microsoft sign-in and MFA in the Windows or browser prompt. Outlook handles
credentials and token storage. The CLI does not ask for a password or accept a
token on the command line. For a nonstandard install, set `OUTLOOK_EXE` to the
Linux path of Classic Outlook.

With that profile open and connected, run:

```sh
./outlook exchange-status 'you@example.com' 'Work'
./outlook mapi-stores 'Work'
./outlook mapi-inbox 'Work'
```

`exchange-status` checks Outlook's reported connection for the matching
default Exchange mailbox in the active profile. It exits unsuccessfully if the
profile or mailbox does not match, or if Outlook is offline or disconnected.
If the organization blocks the browser sign-in or requires a compliant device,
the CLI follows those policies. A real Exchange mailbox is required to verify
authentication end to end.

- [Microsoft's Classic Outlook profile setup](https://support.microsoft.com/en-us/outlook/create-an-outlook-profile)
- [Outlook Exchange connection modes](https://learn.microsoft.com/en-us/office/vba/api/outlook.olexchangeconnectionmode)
- [Soda Office compatibility and sign-in issues](https://github.com/bottlesdevs/programs/issues/500)

## Build from source

Install GNU Make and the x86 and x64 MinGW-w64 C++ compilers, then run `make`.
For example, on Debian or Ubuntu:

```sh
sudo apt install make g++-mingw-w64-i686 g++-mingw-w64-x86-64
make
```

LLVM-MinGW is also supported:

```sh
make CXX64=x86_64-w64-mingw32-clang++ CXX32=i686-w64-mingw32-clang++
```

Run `make dist VERSION=0.1.0` to build the release archive and checksum.
The archive contains the launcher and both CLI executables. The Windows
executables must run inside the Wine prefix that contains Outlook.
