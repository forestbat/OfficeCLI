<!--
  @aionui/officecli-sdk — AionUI-scoped distribution of the officecli Node SDK.
  Published ONLY for @aionui/officecli-sdk (selected in publish-sdk.yml). The
  other names (@officecli/sdk, @officecli/officecli-sdk, officecli-sdk) use
  README.md with their name substituted. Fill in AionUI branding below —
  placeholders are marked TODO(aionui).
-->

# @aionui/officecli-sdk

A thin **async** Node.js client over [officecli](https://github.com/iOfficeAI/OfficeCLI)'s
resident pipe — the **AionUI**-scoped distribution. Same SDK as `@officecli/sdk`.

```bash
npm install @aionui/officecli-sdk
```

Installing the SDK pulls `@officecli/officecli`, which bundles an auto-updating
native binary — so the CLI comes with it. If the binary is ever missing, the SDK
provisions it on first use (downloads the bundled signed binary, or falls back to
the official installer).

## Usage

```js
const oc = require('@aionui/officecli-sdk');

const doc = await oc.create('report.xlsx', ['--force']);
try {
  await doc.send({ command: 'set', path: '/Sheet1/A1', props: { text: 'Hello' } });
  const a1 = await doc.send({ command: 'get', path: '/Sheet1/A1' }); // → envelope object
  console.log(a1);
} finally {
  await doc.close(); // flushes to disk
}
```

## API

- `await create(path, args?, options?)` → `Document` — make a new file.
- `await open(path, options?)` → `Document` — open an existing file.
- `Document.send(item, asJson = true, timeoutMs?)` — forward one command.
- `Document.batch(items, { force = true, stopOnError = false, timeoutMs? })`.
- `Document.alive(timeoutMs?)` / `Document.close()`.

`options`: `{ binary?, timeoutMs?, autoInstall? }`. Transport/process failures
throw `OfficeCliError`; business outcomes live in the returned envelope's
`success` field.

- Source, issues and full docs: <https://github.com/iOfficeAI/OfficeCLI>
- AionUI: <!-- TODO(aionui): add AionUI homepage / product link here -->

Licensed under Apache-2.0.
