# Quick Templates – Outlook add-in

Inserts email templates **where your cursor is**, with fill-in placeholders, a subject line, categories and search.

- Works in classic Outlook for Microsoft 365 (and Outlook on the web)
- Templates are stored in your own mailbox (the add-in's roaming settings). This repo only holds the code.
- No sign-in, no server, no tracking

## Files

| File | Purpose |
|---|---|
| `manifest.xml` | The file you install in Outlook |
| `taskpane.html` / `.css` / `.js` | The add-in panel |
| `commands.html` | Required by Outlook for ribbon buttons |
| `index.html` | Simple landing page (the manifest's support link) |
| `icon-*.png` | Icons |
| `lz-string.min.js` | Compression library (MIT licence, see `lz-string.LICENSE.txt`) |

All URLs in `manifest.xml` point to `https://georgefarrellx.github.io/outlook-templates/`.
If you host the files somewhere else, find-and-replace that address in `manifest.xml`.

## Placeholders

Write `{Anything}` in a template's text or subject. When you insert the template, you're asked to fill each placeholder in.

These placeholders fill themselves in, and you can still change the value before inserting:

| Placeholder | Filled with |
|---|---|
| `{FirstName}` | First name of the person on the To line. With two people it becomes "John and Jane". |
| `{FullName}` | Display name of the first person on the To line |
| `{Today}` | Today's date, e.g. 15 September 2026 |
| `{MyName}` / `{MyFirstName}` | Your name from Outlook |

If you leave a placeholder blank, the add-in asks you to confirm first. A blank `{FirstName}` in "Hi {FirstName}," gives "Hi,".

## Subject when inserting

Choose an option above the template list before clicking a template. You can also change it on the fill-in screen:

- **Replace subject** uses the template's subject instead of the email's current subject.
- **Keep current subject** leaves the subject unchanged, including on an email with no subject. Placeholders used only in the template's subject do not need filling in.
- **Add to current subject** keeps the current subject and adds ` - ` followed by the filled-in template subject. For example, `RE: Accounts` becomes `RE: Accounts - Documents needed`. If the email has no subject, it uses the template subject without a separator.

A template with no subject always leaves the current subject alone. **Undo subject** restores the previous subject after replacing or adding.

Choose your usual option under **Settings → Subject line**. That default returns after each insertion and when you switch emails. Your previous replace/keep default carries over automatically.

## Limits

- Outlook gives add-ins about 32 KB of mailbox storage. Templates are compressed, and the meter in Settings shows how much you've used.
- Pictures and attachments aren't stored.
- Settings → Backup gives you a copy of every template. Keep one somewhere safe.
- If two compose windows have the panel open, only edit templates in one of them. Reopen the panel in the other window to see the changes.
