# ViTO Office 365 Mail Banner Automation

This is a separate Outlook event-based add-in project. It does not use or modify
the ViTO signature pilot's manifest, AppId, JavaScript bundle, NAA registration,
Graph permissions, or GitHub Pages repository.

## Campaign

- Banner: `https://www.vito.com.tr/images/gastechbangkok2026banner.png`
- Display width: 600 px
- Automatic expiry: 18 September 2026, 00:00 (UTC+03:00)
- Event: `OnNewMessageCompose` (new, reply, reply all, forward)

## Signature-preservation strategy

- Classic Outlook, new Outlook, and Outlook on the web: read the complete HTML
  body, insert the banner immediately before the reply/forward boundary, and set
  the complete body back without changing the existing signature HTML.
- Outlook mobile: read the current reply body with `BodyMode.HostConfig`, then
  replace the current signature with the same signature HTML plus the banner.
- The campaign marker prevents a duplicate banner in the current compose area.
  A banner inside quoted history does not suppress the current message's banner.

## Required pilot tests before tenant deployment

Do not assign this add-in tenant-wide until all tests pass for a pilot mailbox:

1. Classic Outlook: new, reply, reply all, forward.
2. New Outlook: new, reply, reply all, forward.
3. Outlook on the web: new, reply, reply all, forward.
4. Outlook for iOS and Android: new, full-screen reply, reply all, forward.
5. Confirm the user's existing text, logo, links, spacing, and font are unchanged.
6. Confirm the banner is below the current signature and above quoted history.
7. Reply to a message that already contains an older banner and confirm only one
   new banner appears in the current compose area.

## Hosting target

Create a separate GitHub repository named:

`vito-mail-banner-automation-20260828`

Publish the repository root through GitHub Pages. The manifest already targets:

`https://infovitogit.github.io/vito-mail-banner-automation-20260828/`

## Local validation

Run:

```powershell
node --check .\banner-runtime.js
node .\tests\banner-runtime.test.js
```

The expected test output is:

`All ViTO banner runtime tests passed.`

## Deployment sequence

1. Confirm the public banner URL returns `Content-Type: image/png`.
2. Publish this project to its separate GitHub Pages repository.
3. Confirm `runtime.html` and `banner-runtime.js` return HTTP 200.
4. Validate `manifest.xml` with the Microsoft Office Add-in validator.
5. Assign `manifest.xml` only to the pilot user/group.
6. Run the complete client matrix above.
7. Expand the assignment only after approval.

## Rollback

Remove or disable the `ViTO Mail Banner Automation` deployment in Microsoft 365
Integrated Apps. The existing personal signatures and the separate ViTO signature
pilot remain independent.

