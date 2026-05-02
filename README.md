# Plex Acquisition — Landing Page

A bilingual (French / English) one-page website to collect leads from people who might be interested in selling their plex (multi-unit residential building) in Quebec.

The page is meant to be the destination for cards / flyers distributed to plex owners. It tells them who you are, what you do, how the process works, and lets them either contact you directly or leave their info so you can call them back.

---

## What's in the file

`index.html` is a single, self-contained file. No build step, no dependencies, no backend. It works as-is by double-clicking it.

The page contains five sections:

1. **Nav bar** — logo, section links, and a FR / EN language toggle (defaults to French)
2. **Hero** — pitch, headline, two call-to-action buttons
3. **About** — who we are + three values (Discretion, Speed, Respect for tenants)
4. **Process** — three-step explanation of how the transaction works
5. **Contact** — your direct phone / email + a callback request form

---

## Design choices

- **Aesthetic**: warm Quebec / Plateau-inspired palette — cream paper background, terre cuite (brick red) accent, deep ink for text. Editorial / refined feel rather than corporate or salesy.
- **Typography**: Fraunces (serif display) paired with Inter Tight (body) — both loaded from Google Fonts.
- **Bilingual**: every translatable element has `data-fr` and `data-en` attributes; the toggle swaps them via JavaScript without reloading the page.
- **Mobile-friendly**: the layout collapses cleanly on phones (single column, navigation simplifies).

---

## Before going live — what you need to change

Open `index.html` in any text editor and replace these placeholders:

| Placeholder | What to replace with |
|---|---|
| `Maison Habitat` | Your actual brand name (appears in nav and footer) |
| `514 000-0000` | Your real phone number |
| `contact@example.com` | Your real email |
| `YOUR_FORM_ID` | Your Formspree form ID — see next section |
| `Grand Montréal & Rive-Sud` | Your actual service area, if different |
| Hero / About / Process copy | Adjust tone and details to match your real positioning |

---

## Connecting the form to your inbox (Formspree)

The form needs an endpoint to send submissions to your email. The simplest free option is Formspree:

1. Go to [formspree.io](https://formspree.io) and sign up with the email where you want to receive leads
2. Click **New Form**, name it (e.g. "Plex Leads"), and choose the inbox email
3. Formspree gives you an endpoint like `https://formspree.io/f/abc123xyz`
4. In `index.html`, find this line:
   ```html
   <form id="leadForm" action="https://formspree.io/f/YOUR_FORM_ID" method="POST">
   ```
5. Replace `YOUR_FORM_ID` with your real ID (e.g. `abc123xyz`)
6. Save the file — submissions will now arrive in your inbox

The free tier covers 50 submissions per month, which is generally enough for a card-based outreach campaign.

**Alternatives** if you don't want Formspree: [Web3Forms](https://web3forms.com), [Formsubmit](https://formsubmit.co), or a Google Form embedded via iframe.

---

## Hosting the page (getting a public URL)

You need a hosting service to give the page a real internet address. Free options:

### Option A — Netlify Drop (fastest, no account needed)

1. Go to [app.netlify.com/drop](https://app.netlify.com/drop)
2. Drag `index.html` onto the page
3. You get a URL like `random-name-12345.netlify.app` instantly
4. (Optional) Create a free Netlify account to claim the site, change the subdomain, or attach a custom domain

### Option B — Tiiny.host

1. Go to [tiiny.host](https://tiiny.host)
2. Upload the file, choose a subdomain
3. Done

### Option C — Custom domain (recommended for the cards)

A real domain like `maisonhabitat.ca` looks more professional on printed cards.

1. Buy the domain (~$15–20/year) at [Namecheap](https://www.namecheap.com), [OVH](https://www.ovh.com), or [Hover](https://www.hover.com)
2. Host the page on Netlify (free)
3. In Netlify's dashboard, go to **Domain settings** and follow their instructions to point your domain to the site
4. SSL (https) is automatic and free via Netlify

---

## Putting it on the cards

Once the page is live, add to the card:

- The clean URL (e.g. `maisonhabitat.ca`)
- A QR code that points to the same URL — generate one for free at [qr-code-generator.com](https://www.qr-code-generator.com)
- The phone number, in case people want to skip the page entirely

Two paths for the recipient: **call directly** or **scan / type the URL → leave info → get called back**.

---

## Iterating later

Things you might want to add as the project grows:

- **Testimonials** from past sellers (huge trust signal)
- **"What we look for"** criteria (number of units, regions, condition) so you filter out bad leads earlier
- **FAQ** (do you buy occupied buildings, do you cover transfer tax, etc.)
- **Photos** of the team or buildings you've acquired
- **Analytics** — add Plausible or Google Analytics to see how many visitors come from cards
- **A second language flow** if you want different content per language rather than a direct translation

---

## Files in this project

- `index.html` — the entire website (HTML + CSS + JavaScript in one file)
- `README.md` — this document

---

*Built for: Quebec plex acquisition lead-gen via printed cards.*
