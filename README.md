# gpt-ocr

A small **static web app** that sends slide or document photos to the **OpenAI vision API** and returns transcribed text while trying to preserve layout (Markdown or HTML). It runs entirely in the browser, stores your API key in **localStorage**, and can be hosted on **GitHub Pages** with automatic deploys from `main`.

## Features

- **Batch image upload** (up to 150 images) with a progress bar and sequential or batched API calls (`gpt-4o-mini` by default).
- **OpenAI API key** field with persistence in `localStorage` (this device only).
- **Output format**: Markdown (editors, Git repos) or **HTML** (richer structure for Word-style paste).
- **Tuning**: images per request (1–5), optional lower image `detail` for fewer image tokens.
- **Copy** as plain text, **Copy for Word** (HTML + plain on the clipboard), and **Preview for Google Docs** (opens a new tab so you can copy from rendered HTML—Google Docs often ignores synthetic clipboard HTML).
- **GitHub Actions** workflow uploads [`public/`](public/) to GitHub Pages on every push to `main`.

## Live site

If Pages is enabled for this repository, the app is served from your usual GitHub Pages URL, for example:

`https://<user>.github.io/<repo>/`

## Run locally

No build step. Serve the `public` folder over HTTP (file URLs may block some features):

```bash
cd public
python3 -m http.server 8080
```

Then open `http://localhost:8080`.

## Deploy (GitHub Pages)

1. Push this repository to GitHub on the **`main`** branch.
2. **Settings → Pages → Build and deployment**: set **Source** to **GitHub Actions** (not “Deploy from branch” for this workflow).
3. After the first successful run of [`.github/workflows/deploy-pages.yml`](.github/workflows/deploy-pages.yml), the site URL appears on the workflow run and under **Environments → github-pages**.

The workflow publishes only the [`public/`](public/) directory as the site root.

## Security and privacy

- The app calls **OpenAI** directly from the browser with your key. The key is **visible in DevTools** and in **localStorage**; anyone who can use the deployed page could abuse a key stored there.
- For a **public** repo or shared URL, prefer a **restricted** OpenAI key (low limits, no billing surprises), rotate if exposed, or put the API behind your own **server proxy** instead of this static pattern.
- Slide images are read in the browser and sent to OpenAI per your account’s data/API policies.

## Project layout

| Path | Purpose |
|------|---------|
| [`public/index.html`](public/index.html) | Page structure |
| [`public/app.js`](public/app.js) | OpenAI calls, UI logic, preview |
| [`public/styles.css`](public/styles.css) | Layout and theme |
| [`.github/workflows/deploy-pages.yml`](.github/workflows/deploy-pages.yml) | Pages deploy |

## Requirements

- A modern browser with **Fetch** and **FileReader**.
- An **OpenAI API key** with access to vision-capable chat models (the code uses **`gpt-4o-mini`**).

## License

Licensed under the **Apache License, Version 2.0**. See [`LICENSE`](LICENSE).

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an “AS IS” basis, without warranties or conditions of any kind, either express or implied. See the License for the specific language governing permissions and limitations.

OpenAI, ChatGPT, and related marks are trademarks of their respective owners. This project is not affiliated with or endorsed by OpenAI.
