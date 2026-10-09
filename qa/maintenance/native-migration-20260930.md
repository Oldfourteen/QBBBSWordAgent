# Native PowerShell migration verification

## Scope

Windows 10/11; original four-gate workflow unchanged. agent.ps1 is now the native entry. Legacy Python source retained for maintenance, not required to run native commands. No user reports, outlines, source JSON or templates deleted. No registry, permanent environment variable, ACL or Office installation changes performed.

## Verified

- Windows PowerShell 5.1 native doctor and existing task precheck succeed without Python invocation.
- 16 isolated tests: gate ordering/confirmation, no content overwrite, placeholders, precheck state, numbering, stale field rejection, mock end-to-end gates, cache reuse, artifact tampering and path containment.
- One real Word multipage integration test using a COPY of the existing three-page school template and 20260930 report content. 19 TOC entries matched footer pages; fonts and TOC numbering/format checked in saved DOCX. Actual body 9 pages, total 13. This does NOT satisfy the real task's 12-page target and did not change its state.
- Real standalone Word capability probe succeeded: opened, repaginated, updated fields, saved and exited.
- Portable ZIP extracted into a Chinese/spaced directory. doctor and 16 isolated tests passed with child PATH restricted to Windows directories (no Python executable on PATH).

## Bugs fixed during live testing

- Word HWND was incorrectly read from Application. It is now read from the opened document's ActiveWindow. Process ownership is checked before quitting an application; existing user sessions are not terminated.
- PS5.1 Start-Process handle lifetime could lose ExitCode after completion. Native handle retained during bounded wait.
- Saved DOCX may normalize explicit formatting into styles. Validator resolves style inheritance rather than treating normalized XML as a format failure.
- PS5.1 atomic Replace backup argument and ZIP assembly loading compatibility fixed.

## Remaining limitations

- WPS registration not detected here. Adapter supports compatible COM ProgIDs but has not passed a WPS machine test. Detection is not a universal version/license compatibility promise.
- Automatic choice prefers registered Word, otherwise WPS; a failed live probe stops for diagnosis rather than launching further applications into a potentially hung session.
- Current native templates must be three-page single-section DOCX. Multi-section templates are rejected, not silently rewritten. Native publication is DOCX; legacy DOC conversion and explicit appendices are not implemented.
- Real desktop tests used the client's approved execution mechanism after unsigned-script execution was refused. This proves automation with authorization, not unrestricted operation in every sandbox/client.
- The reported historical temporary-work-file dialog was not reproduced in authorized tests; no claim of a universal Office environment repair is made.
- Initial HWND failure left an unowned/unknown-handle background Word instance; no global Word termination was attempted. Later test instances were tracked and exited normally.

## Distribution

dist/report-agent-native-20260930-082129-44561f.zip

SHA256: 001A57E8EAB7BD13AD8C3131B93762B5EC106146AD3EE50A1851EED2A20A6230

ZIP contains code, policies, generic outline, command docs and sanitized constraint nodes. It excludes private reports, tasks, school template DOCX, Python runtime and Office.

The fixed workflow hash is unchanged: 00A6800B7CB42ABDCFB7B32C04FF3555BD6FBEC20619418BFF4CFC4B3E6E0D44.
