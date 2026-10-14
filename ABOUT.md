# v6-b: v6 with photos on the toolkit cards

v6-b is a copy of v6 with one change: the toolkit cards have photos. v6 itself was not edited. Open `index.html` in Chrome to see it.

## What's new in v6-b

The empty middle of each toolkit card now holds a photo that matches the card, recoloured to the card's own colour so the deck still looks like one set:

| Card | Photo | Licence |
|------|-------|---------|
| 1 Broadcast management | TV master control room (Food Network) | CC BY 4.0 |
| 2 ERP and SAP | Tokyo Stock Exchange market centre and ticker | CC BY 2.0 |
| 3 Playout and MAM | Sony D1 broadcast video deck | CC BY-SA 3.0 |
| 4 Infrastructure | Data centre aisle of server racks | CC BY-SA 4.0 |
| 5 Product and design | Hand-drawn website wireframe | CC BY 3.0 |
| 6 Running the business | Meeting room | CC0 |
| 7 Credentials | Graduation cap | Public domain |

- **Credits:** all photos come from Wikimedia Commons and are free to use. Some licences require crediting the photographer, so there is a "Toolkit photo credits" link at the very bottom of the page. The full details are in `toolkit/credits.json`.
- **Layout:** each photo takes whatever space the card's text leaves, so it also fits on phones.
- **Smoothness:** the photos are coloured ahead of time rather than live in the browser, so the deck scrolls as smoothly as v6's.

---

## Everything else is the same as v6

v6 combines the parts of v1-b and v2 you liked into one page. Neither v1-b nor v2 was changed.

| # | Section | Comes from | Notes |
|---|---------|-----------|-------|
| – | Top bar: KM, ON AIR + live IST clock, red underline, light/dark switch | v1-b | Menu now lists How I work, Journey, Reach, Track record, Toolkit and Contact |
| 1 | Landing: portrait with TV name caption, headline, buttons, colour bars | v1-b | **New:** an "ON AIR WITH" news ticker of the 15 broadcasters you've taken live, with a pause button |
| 2 | Scattering letters on scroll | v2 | Now says "From first meeting to on air." so it doesn't repeat the landing headline |
| 3 | "Channel 1, Tuesday" schedule that builds up step by step | v1-b | Heading is now "Eight steps, every go-live." |
| 4 | Journey: "First, / I learned the systems." | v2 | Small progress dots added at the bottom |
| 5 | 12 go-lives globe with the country list | v2 | Same globe |
| 6 | As-run log, case studies, 18 solutions | v1-b | The solutions are shown as tags |
| 7 | Toolkit card deck | v2 | **New 7th card:** your credentials (CSPO, GNIIT, B.Com) |
| 8 | Your quote | v1-b | Set in the serif italic font from the journey section |
| 9 | "What's next?" orb that fills the screen, then the gradient contact panel | v2 | Phone number still appears only when clicked |

**Removed (you didn't like them):** v2's gradient landing with the 3D "KM", and v1-b's plain "Toolkit and credentials" grid. The card deck replaces the grid.

**Checked:**
- desktop, tablet and phone, in dark and light mode
- reduced motion (everything shows without animation)
- keyboard and phone-menu use
- no sideways scrolling at any width
- no errors

**Smoothness:** averages 60 frames a second while scrolling, with a typical 12 slow frames per test against 5 for v1-b and 2–3 for v2. v6 runs more effects at once than either parent.

**Accent colour:** v2's brighter orange-red (#FF3B2F), used on the buttons, ON AIR tags, "on air" letters, menu underline, name caption bar, globe markers and first toolkit card. Very small red labels (like "Step 1") use a slightly deeper red in light mode so they stay readable.
