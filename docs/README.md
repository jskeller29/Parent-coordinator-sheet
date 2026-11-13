# The User Guide

| File | What it is |
|---|---|
| `USER_GUIDE.md` | The guide's text. **The only file to edit.** |
| `build_guide.js` | Makes the three files below from `USER_GUIDE.md`. |
| `PC_Tracker_User_Guide.pdf` | The guide as a PDF, with page numbers and clickable contents. |
| `PC_Tracker_User_Guide.docx` | The guide as a Word file. Upload it to Drive and open it with Google Docs. |
| `guide.html` | The page the PDF is printed from. Made by the script; don't edit it. |

## Rebuilding after a change

```bash
node docs/build_guide.js
```

It needs the `docx` package (`npm install docx`) for the Word file, Playwright for the PDF, and
Python with `pypdf` (`pip install pypdf`) to make the PDF's jump links work in simple viewers.
A missing one is skipped and the others are still made.

The look is the "District" style of the tracker's own screens (`Styles.html`): navy #183e69,
Source Sans 3 text, Source Serif 4 headings, thin grey lines. Writing style: short plain
sentences, emoji only where they are part of a menu or tab name in the sheet, the same words
as the sheet's own screens.

What `USER_GUIDE.md` may use: `#` title, `##` part, `###` section, `####` sub-section,
paragraphs, `- ` and `1. ` lists (a continued line is indented), tables, and these boxes:

```
> [!NOTE]        a tip (grey)
> [!IMPORTANT]   read this (navy)
> [!WARNING]     can lose data or share something private (red)
```

`<!-- comments -->` never appear in the guide.

## Putting it in Google Drive

The tracker's 📖 Open User Guide link is the **Version Tracker's "Document" column (F)**:
the newest dated row with a link there. Paste the new guide's link in the row of the release
that brings it (leave it blank in a release that keeps the same guide). When the tracker has
no Document link (or a sheet can't read it), the newest file in the guide folder is used
(`GUIDE_FOLDER_ID` in the 📌 YOUR LINKS box at the top of `Menus.gs`), so keep the newest
guide there too.

- **Google Doc (recommended):** upload the .docx to the guide folder, right-click it → Open
  with → Google Docs, then delete the .docx so the Google Doc is the only file there. Share it
  "Anyone with the link → Viewer". Each heading has its own link: see **The headings the
  tracker links to** below.
- **PDF:** upload the PDF instead. Its contents and section buttons jump inside the PDF, but a
  link from the sheet always opens it at page 1. Drive's and GitHub's previews may ignore the
  jump links; downloading it or opening it in a PDF reader works.

Keep old versions in a subfolder (for example "Old guides"): the backup only looks at files
directly in the guide folder.

## The headings the tracker links to

Four buttons in the tracker open the guide at one heading. Each heading has the same words
as its button, and the code keeps the list in `GUIDE_LINK_HEADINGS_` (Menus.gs, just under
the 📌 YOUR LINKS box). A test fails if `USER_GUIDE.md` loses one of these headings, so
rename a heading in both places.

| Heading in the guide | Where it is | Link in 📌 YOUR LINKS | The button that opens it |
|---|---|---|---|
| How to download your data | Part 1, Step 5: Load your data | `downloadDataDoc` | Welcome pop-up "How to download your data"; Getting Started step 1 "Guide: how to download your data" |
| What Build Sheets does | Part 2, Build and Full Sync | `buildDoc` | Getting Started step 2 "Guide: what Build Sheets does" |
| What each setting does | Part 2, Settings | `settingsDoc` | Getting Started step 3 "Guide: what each setting does" |
| How to log a parent contact | Part 2, The Contact Log | `contactLogDoc` | Getting Started step 4 "Guide: how to log a parent contact" |

`guideDoc` (Getting Started Help, "Getting started guide") opens the whole guide, so it has
no heading. The two videos (`setupVideo`, `downloadDataVideo`) are YouTube links.

The Google Doc in Drive does not change when this file does. Its Contents page is plain
text with links, made from the Word file it was imported from, so headings typed into the
Doc later don't appear there by themselves. Two ways to get these headings into it:

**A. A new Doc from the new Word file** (also brings every other change made to the guide
since the Doc was made):

1. Upload the new `PC_Tracker_User_Guide.docx` to the guide folder, right-click it → Open
   with → Google Docs, then delete the .docx. Share the Doc "Anyone with the link → Viewer".
2. Put its link in the Version Tracker's Document column (F) in the newest release's row.
3. Copy each heading's link (step 3 below) into 📌 YOUR LINKS, and the Doc's own link into
   `guideDoc`. Update the two links in its **Start here** box.
4. In the old Doc, type "This guide has moved:" and the new link at the very top (old links
   still lead somewhere), or move the old Doc to the "Old guides" subfolder.

**B. The same Doc, by hand** (keeps its link):

1. Type each heading at the place in the table: rename "Where to get each export" to "How
   to download your data", and add the others as new lines (Build and Full Sync also gets
   "What Full Sync does" above its second paragraph).
2. Make each one a Heading 3: click in the line, then pick Heading 3 in the style box
   (where it says Normal text). The outline on the left lists it now; the Contents page only
   if you type a line there too.
3. Get its link: click the heading in the outline on the left; the address in the browser
   now ends in `#heading=h.` and a code. Copy the whole address. (A bookmark works too:
   Insert → Bookmark, then click the bookmark → Copy link.)
4. Paste it into that link's line in 📌 YOUR LINKS (Extensions → Apps Script → Menus),
   save, then paste Menus into the master template and every sheet.

🔗 My Links (🐞 Debug) shows each link with the heading it should open.

## Updating the guide that is already in Drive

Edit the Google Doc **in place**; don't upload a new file. Its link stays the same, so the
Version Tracker's F column, the 👋 Welcome tabs and the document links in 📌 YOUR LINKS
(`#bookmark=` links into the Doc) keep working.

- Small change: type it in the Doc, or copy that part from the new .docx and paste it.
- Avoid replacing the whole text with a paste: the bookmarks the YOUR LINKS document
  links point to would be gone, and those buttons would open the Doc at the top.
- A brand-new file means a new link: put it in the tracker's F column and update the
  document links in YOUR LINKS (then paste Menus into the master template).

At every release, update the two links in the **Start here** box on the first page (right-click
the link → Edit link): Make your copy = the Blank copy's /copy link (tracker column D), See an
example tracker = the Example copy's link (column E). 📦 Make Release Copies shows both.

