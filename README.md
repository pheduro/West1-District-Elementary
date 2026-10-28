# West I District Elementary Site

Mandaue City, West I District, SY 2026-2027. Six schools: Bakilid, Cabancalan I, Maguikay, Tabok 1, Tabok 2, Tingub.

Two tabs on the site:

- **SAPM**: the School Assessment Performance Monitoring dashboard (data in `West1_SAPM_data.json`, embedded in `index.html`).
- **CID Monitoring Checklist**: School Heads answer each monitoring checklist at every PSDS visit using a personal school code; the PSDS has a PIN-protected summary. Responses can be edited and downloaded as the original Excel forms.

## Files

- `index.html`: the whole site (SAPM + CID tab)
- `forms.json`: checklist structure (sections, items, Excel cell map)
- `templates/*.xlsx`: original Excel checklists used for downloads
- `js/cid.js`, `js/cid-xlsx.js`, `css/cid.css`, `vendor/jszip.min.js`: CID tab code

## CID backend

Separate Supabase project (`west1-district`), not shared with other projects. Tables have no public access; the browser only calls SECURITY DEFINER functions (`cid_login`, `cid_list`, `cid_save`, `cid_delete`, `cid_psds_all`) that check hashed school codes / PSDS PIN, with attempt throttling. The anon key in `js/cid.js` is public by design. School codes and the PSDS PIN are NOT stored in this repo.

## Deploy

Static site, no build step. On Vercel: Framework Preset "Other", leave build and output settings empty. Pushing to `main` redeploys.
