# SEO setup — manual steps

The on-page work (meta tags, sitemap, robots.txt, structured data) is done and committed. What's left requires your Google/Bing accounts, so I can't do it for you.

## 1. Google Search Console
1. Go to search.google.com/search-console → Add property → choose "URL prefix" → `https://papertrailacademic.com/`
2. Verify ownership. Since this is GitHub Pages with a CNAME on Cloudflare/your registrar, the easiest method is **DNS verification** (add a TXT record) or the **HTML file** method (drop the verification file in the repo root and push).
3. Once verified, go to Sitemaps → submit `https://papertrailacademic.com/sitemap.xml`
4. Use URL Inspection on `/`, `/inspect/`, `/stylematch/`, `/verify/` to request indexing directly — speeds up first crawl instead of waiting.

## 2. Bing Webmaster Tools
1. bing.com/webmasters → Add site → same URL
2. Bing lets you **import verified sites straight from Search Console** — use that instead of re-verifying manually.
3. Submit the same sitemap URL.

## 3. Google Business / knowledge panel (optional, later)
Not essential now since this isn't a local business, but if you want a knowledge panel eventually, having consistent `sameAs` links (Twitter/X, LinkedIn, etc.) pointing back to papertrailacademic.com helps — I left `sameAs` out of the Organization schema since none of those exist yet. Add them to the JSON-LD in `index.html` if/when you create social profiles.

## 4. Analytics (optional)
No analytics is currently wired up, so you have no way to see if any of this is working. Plausible or GoatCounter are lightweight, privacy-respecting options that fit a GitHub Pages/no-backend site — a single `<script>` tag, no cookie banner needed. Worth adding before you invest more time in SEO, otherwise you're flying blind.

## 5. Backlinks (ongoing, no tooling needed)
The single highest-leverage thing for a niche B2B education product: get listed on education-tech directories (Common Sense Education, EdTech reviews sites) and get a mention from any teacher blogs/newsletters you can reach. Domain authority from a handful of relevant links will outweigh on-page tweaks at this stage.

---

## What was already implemented in this pass
- `robots.txt` at site root (allows all, disallows `/confirmed/`, `/success/`, `/reset-password/`, points to sitemap)
- Open Graph + Twitter Card tags on all 10 public pages, using a new shared `og-image.png` (1200×630, on-brand)
- JSON-LD structured data: `Organization` + `WebSite` on the homepage, `SoftwareApplication` (with pricing where known) on Inspect/StyleMatch/Verify/Citations/Oral/Write, `WebPage` on For Schools
- `noindex, nofollow` added to the three utility pages (confirmed, success, reset-password) so they don't compete for search visibility
- Titles, meta descriptions, and canonical URLs were already in good shape site-wide — no changes needed there
