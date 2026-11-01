# docxkeun-fonts

Open fonts with the same character widths as the fonts Word documents usually name, so [docxkeun](https://github.com/kevinkjs/docxkeun) lays out lines and pages like Word does and can embed them in PDFs without shipping proprietary fonts.

| Files | Stands in for | Source | License |
|---|---|---|---|
| `Carlito-*` | Calibri, Calibri Light | Google Fonts, `ofl/carlito` | SIL OFL 1.1 ([LICENSE-Carlito.txt](LICENSE-Carlito.txt)) |
| `Caladea-*` | Cambria | Google Fonts, `ofl/caladea` | SIL OFL 1.1 ([LICENSE-Caladea.txt](LICENSE-Caladea.txt)) |
| `LiberationSans-*` | Arial, Helvetica | Liberation 2.1.5, github.com/liberationfonts | SIL OFL 1.1 ([LICENSE-Liberation.txt](LICENSE-Liberation.txt)) |
| `LiberationSerif-*` | Times New Roman | Liberation 2.1.5 | SIL OFL 1.1 |
| `LiberationMono-*` | Courier New | Liberation 2.1.5 | SIL OFL 1.1 |

Each family has Regular, Bold, Italic and BoldItalic (20 files, about 7 MB). The fonts are unmodified.

## Install

```bash
npm install docxkeun docxkeun-fonts
```

## Use

```ts
import { FontRegistry, loadBundledFonts } from 'docxkeun';
import { fontUrls } from 'docxkeun-fonts';

const fonts = new FontRegistry();
await loadBundledFonts(fonts, { urls: fontUrls });                         // browser (with a bundler)
await loadBundledFonts(fonts, { urls: fontUrls, families: ['Carlito'] });  // only what you need
```

`fontUrls` uses `new URL('./file.ttf', import.meta.url)`, which Vite, webpack 5 and Rollup recognise: they copy the font files into your build. In Node, the URLs are `file:` URLs:

```ts
import { readFile } from 'node:fs/promises';
await loadBundledFonts(fonts, { load: (file) => readFile(fontUrls[file]) });
```

You can also host the `.ttf` files yourself and pass `{ baseUrl: 'https://cdn.example.com/fonts/' }`.

## License

The fonts are licensed under the SIL Open Font License 1.1 (see the `LICENSE-*.txt` files). The OFL lets you bundle and redistribute them with software, but the fonts themselves must not be sold on their own.
