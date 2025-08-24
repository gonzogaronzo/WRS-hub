# Wilson Ninja Hub (GitHub Pages)
Static, Google-free hub for Wilson lessons. Data lives in `/data/*.json`. Buttons generate files client-side (PPTX via PptxGenJS).

## Host on GitHub Pages
1. Create a new GitHub repo, upload this folder.
2. Settings → Pages → Source: `main`, Folder: `/root`.
3. Your hub lives at `https://<you>.github.io/<repo>/`.

## Configure
- Edit `/data/scrolls.json`, `/data/secret_words.json`, `/data/story_scrolls.json`.
- Set active substep in `/js/app.js` (ACTIVE_SUBSTEP_DEFAULT) or via the dropdown.

## Generate Lesson Deck (PPTX)
This repo references PptxGenJS from a CDN (no server required). Click “Summon Lesson Deck” to download a deck built from your JSON for the chosen Substep.
