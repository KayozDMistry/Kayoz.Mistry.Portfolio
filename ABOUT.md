# v6-e: the phone edition moves like the desktop site

v6-e is a copy of v6-d. Only `mobile.html` changed; `index.html` is identical to v6-d's. v6-d itself was not edited.

## What changed on phones

| # | Section | Before (v6-d) | Now (v6-e) |
|---|---------|---------------|------------|
| 1 | "From first meeting to on air." | Letters dropped in once | Same scroll-driven tumbling letters as desktop |
| 2 | Channel 1, Tuesday | Card stuck to the top while step text scrolled underneath it | Card holds still with a progress bar; one step at a time appears below it, never underneath |
| 3 | First, / Then, / Today, | Separate blocks you scrolled past | Same one-word-at-a-time story as desktop, with progress dots |
| 4 | 12 go-lives globe | Globe stuck to the top and country text slid under it | Globe holds still; one country at a time below it, plus a row of country buttons you can tap |
| 5 | Toolkit | Swipe carousel | Same 3D deck as desktop, dealt away as you scroll |

- **Sideways or reduced motion:** with the phone held sideways, or with "reduce motion" switched on, these sections show as simple stacked sections so nothing is cut off. Rotating back brings the animations back.
- **Testing:** tested on 320, 375, 412 and 430px-wide phones and in desktop-site mode. No overlaps, no sideways scrolling, no errors.
- **Google Analytics:** confirmed sending page views from the phone page (G-8XKXJR31YP).

To publish, replace `mobile.html` on GitHub with this one. Nothing else needs uploading again.

---

# Earlier: v6-d, a dedicated phone edition

v6-d is a copy of v6-c with a separate page built for phones.

## How it works

| Device | What it gets |
|--------|--------------|
| Phone (any browser, including Edge or Chrome in "desktop site" mode) | `mobile.html`, the phone edition |
| iPad / tablet | `index.html`, the full site |
| Laptop / desktop | `index.html`, the full site |

- **Spotting a phone:** the main page (`index.html`) checks the device's real screen size and touch screen. It doesn't rely on the browser's name, which changes in desktop-site mode. A phone is sent straight to `mobile.html`.
- **Phone edition:** a single column with a big photo, big text, full-width buttons, a schedule that builds as you scroll, journey cards, a globe that turns to each country, the track record as easy-to-read cards, a **swipeable** toolkit deck, and the "What's next?" orb finale.
- **Desktop-site mode:** the phone edition sizes everything as a share of the screen width. Even if the browser forces a computer-sized layout, text and photos stay big. Body text measured 17.5px on screen in both modes.
- **Phone held sideways:** the photo and headline sit side by side and nothing is cut off.
- **Full site from a phone:** "View the full site" (menu and footer) opens the full site and stays there for that visit.
- **A computer opening `mobile.html`** (a shared link) is sent to the full site.
- **Google Analytics:** both pages use the same tracker (G-8XKXJR31YP).

## Publishing

Upload **all** of these to the GitHub repository, next to each other:
- `index.html`
- `mobile.html`
- the `toolkit` folder
- `profile.webp`, `profile.png`, `favicon.svg` and `og-image.png`

---

# Earlier: v6-c, v6-b that fits every device

v6-c is a copy of v6-b, adjusted so the page fits every screen it's opened on.

## What's new in v6-c

Tested on 17 screen sizes: a foldable phone, small and large phones, phones held sideways, iPad mini / Air / Pro in both orientations, small laptops, a 1920px monitor and a 2560px monitor.

| Device | What was wrong | Fixed by |
|--------|----------------|----------|
| Small phones and foldables | Toolkit card text ran off the bottom of the card | Cards take their height from the screen and use slightly smaller text on narrow phones |
| Phones held sideways | Card deck, globe list and contact panel didn't fit the short screen | Those sections switch to a normal stacked layout when the screen is short and wide; rotating back restores the scroll scenes |
| Small phones | Contact panel cut off at the bottom | Headline and spacing scale with screen height |
| Big monitors (1920px and up) | Page looked small in the middle of the screen | Everything scales up gently, about 37% larger at 2560px |

After the fixes, on all 17 sizes:
- no sideways scrolling
- nothing cut off
- every button big enough to tap
- no errors

Smoothness is the same as v6-b.

---

# Earlier: v6-b, v6 with photos on the toolkit cards

v6-b is a copy of v6 with one change: the toolkit cards have photos. v6 itself was not edited.

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
