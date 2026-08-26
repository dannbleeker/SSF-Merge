# Security

## Reporting

Report a vulnerability privately through GitHub's
[security advisories](https://github.com/dannbleeker/SSF-Merge/security/advisories/new)
rather than in a public issue.

## What this add-in touches

Worth stating plainly, because merge data is usually personal: names,
addresses, salaries, patient identifiers.

- **The merge runs on the user's own machine.** The engine is client-side
  TypeScript. Rows pasted or read from a local file are never uploaded.
- **The only network destination is Microsoft**, and only when the user chooses
  a feature that needs it: reading an Excel table through Microsoft Graph, or
  saving output to OneDrive. Sign-in uses nested app authentication, so there is
  no middle tier holding a token.
- **The add-in has no backend.** There is no server to breach and no log to
  leak.

## Handling untrusted input

A .pptx and a pasted table both arrive from outside. Two habits the code keeps:

- Values are looked up on objects created with a null prototype, so a column
  called `__proto__` or `constructor` is data rather than a hazard.
- Every value written into XML is escaped at the point of writing, not by the
  caller.
