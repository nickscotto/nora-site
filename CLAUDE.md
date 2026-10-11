# norascotto.com

One-page site for Nora Scotto, hair colorist inside Euphoria Day Spa & Salon, 7640 N Wickham Rd, Suite 111, Melbourne, FL 32940. Phone/text 443-521-0660.

## Deploy
Edit `index.html`, commit, `git push`. GitHub Actions (`.github/workflows/deploy.yml`) publishes to GitHub Pages in about a minute and refreshes `/count.json` (launch-spot count from Square) every 30 minutes. New top-level files must be added to the `cp` line in the workflow.

## Brand
- Colors: tokens in `:root` (`--brunette` #4A2F22, `--honey` #B9822F, `--champagne` #E6CFA0, paper/ink). Dark mode redefines them; primary buttons turn honey in dark mode.
- Fonts: Bodoni Moda (headings, `--serif`), Manrope (body, `--sans`).
- Tone: warm, plain, specific. State facts: what she does, where, what it costs, what's included.

## Copy rules
- No fake contrasts ("chosen with you, not assumed", "nothing is decided until..."). Every colorist consults first.
- No invented scarcity, stats, values or guarantees. Only true limits: first 20 new clients, through Oct 31, 2026. No weekly cap.
- Guarantee (Nora's words): "If the color isn't what you discussed at your consult, Nora will fix it that same day."
- Reviews are real and from Maryland (Europa Salon, Nourish & Style). Never present them as Florida reviews.
- Before/after photos are real work; never AI-generated results.
- Ad text and the top of the page must use the same offer wording.

## Current offer (through Oct 31, 2026)
All-over or root color + haircut for $100 (normally $200 = $145 color + $55 cut), first 20 new clients. The $100 can go toward highlights ($155+), balayage ($190+) or gray blending ($175+); client pays the difference, cut included. Elements with `.promo-only` / `.regular-only` swap automatically when the promo ends (date passes or 20 claimed). The countdown shows only at 5+ claimed (`showCountAt`).

## Page structure
Hero (avatar line, offer H1, claim form: first name + mobile -> FormSubmit email to nick@claritypoint.io, then a step-2 booking link) -> credentials -> her work -> offer details -> reviews -> how the visit goes -> shade quiz -> services menu -> about -> FAQ -> visit -> final CTA. Mobile bar: Text + Claim.

## Tracking
Meta Pixel 2106667673542937 (PageView, Schedule on Book taps, Contact on Text taps, Lead on claims). Microsoft Clarity yu7jktz1hq with custom events tapped_book, tapped_text, tapped_photo, claimed_100.

## SEO rules
- One H1. Unique title and meta description. Alt text on every image.
- JSON-LD in `<head>`: HairSalon (keep name, address, phone identical to the Google Business Profile) and FAQPage (keep in sync with the FAQ section).
- `robots.txt`, `sitemap.xml` (update `lastmod` on content changes), canonical https://norascotto.com/.
- Check every change in a browser at phone width before pushing.
