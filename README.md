# Shakthi — Tanvika Karupakula Bharathanatyam Arangetram

Single-page site for Tanvika's Arangetram, served at **https://tanvikakarupakula.com** via GitHub Pages.

## How to publish a change

There is no build step. Any push to the `main` branch redeploys the live site automatically:

```bash
git add -A
git commit -m "describe the change"
git push origin main
```

GitHub Pages rebuilds within about a minute. You can watch progress under this repo's
**Settings → Pages**, or just reload the site after a minute.

That's it — pushing to `main` *is* publishing. There's nothing else to trigger.

## About the domain

The domain (Namecheap) and the GitHub Pages hosting are two separate things that were connected
to each other **once**, and that connection doesn't need to be touched again:

- The `CNAME` file in the repo root contains `tanvikakarupakula.com` — this is what tells GitHub
  Pages which domain to serve. **Don't delete or rename it**, or the custom domain breaks.
- Namecheap's DNS for tanvikakarupakula.com already points at GitHub's servers (four `A` records
  on `@`, one `CNAME` on `www`). This was configured directly in the Namecheap dashboard and has
  nothing to do with this repo — you'd only ever revisit it if you moved hosting off GitHub Pages
  entirely.

So day to day: **editing the site = editing files in this repo and pushing.** No domain or DNS
step involved.

## Structure

Everything lives in one file, `index.html` — sections are marked with banner comments
(HERO, ABOUT, GURU, ARTISTS, PROGRAM, KNOW YOUR ROOTS, GALLERY, CONTACT, FOOTER) and linked from
the nav via `#anchor` links. Images are in `assets/images/`. Styling is Tailwind, loaded from a
CDN `<script>` tag in `<head>` — no `npm install` needed to edit or preview.

To preview locally, just open `index.html` in a browser, or run a tiny local server from the repo
root (`python3 -m http.server`) if you want relative paths to behave exactly like production.

## Still placeholder

Search the file for `[PLACEHOLDER` to find what's left to fill in — currently: the featured-event
title/date/venue, the About section body, the "Know Your Roots" cards, the Gallery intro, and the
Contact section's email/location. The Program section currently lists a Rama-themed margam carried
over from the template; swap in the actual program once it's set.
