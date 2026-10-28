# CP Systems website

The Astro source for [cp-infra.in](https://cp-infra.in).

## Commands

Run from this `website` folder (Node 22.12 or later):

| Command | What it does |
| :-- | :-- |
| `npm install` | Install dependencies |
| `npm run dev` | Dev server at `localhost:4321` |
| `npm run build` | Build the site into `dist/`, which is what gets deployed |
| `npm run preview` | Serve the built `dist/` locally |

## Where things are

- `src/pages/` — one file per page: home, about, contact, `projects/*`, `speciality/*`.
- `src/components/` — header, footer, page hero, project chapters, clips, photo rows, the India map and the granite picker.
- `src/data/clips.json` — every site clip, its caption and the source video it was cut from.
- `src/data/sites.ts` — offices and project sites shown on the home-page map.
- `src/data/indiaMap.json` — India and Bhutan outlines, state lines and rivers for the map (Natural Earth, India point of view, public domain).
- `public/media/` — the web clips and their poster frames.
- `public/images/` — site photographs; `public/images/map/` holds the Natural Earth relief image.
