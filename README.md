# Excel Diff

Excel / CSV files are compared locally in the browser. No file uploads, accounts or usage limits.

- Japanese: https://makiabe.github.io/ExcelDiff/
- English: https://makiabe.github.io/ExcelDiff/en/

## 多言語対応 / Localization

`src/index.template.html` is the shared HTML template. `locales/ja.json` and `locales/en.json` contain visible page text, accessibility labels, comparison messages, errors and exported report text. The comparison code is shared by every language; uploaded spreadsheet contents are never translated or modified.

The generated `index.html`, `en/index.html`, `locales/*.js` and `sitemap.xml` are committed so the site works with ordinary GitHub Pages branch publishing. Do not edit generated files directly.

### Add another language

1. Copy `locales/en.json` to a new file such as `locales/es.json`. Translate its values, preserving keys, HTML tags and placeholders such as `{p0}` and `{locale}`. Do not translate or remove CSS embedded in the report template.
2. Add a registry entry in `site.config.json`, for example:

```json
{"code":"es","label":"Español","path":"es/","dir":"ltr","artwork":"international"}
```

3. Run `npm run build` and `npm run check` using Node.js 20 or later. No package installation is required for these commands.
4. Review desktop and mobile layouts, then commit the dictionary, configuration and generated files together.

The builder automatically creates the language-specific URL, language switcher links, canonical URL, reciprocal hreflang links, structured data and sitemap entries. It rejects missing translations, mismatched interpolation placeholders and duplicate locale paths. Do not publish a language before its translation has been reviewed.

All non-Japanese languages share the supplied text-free mascot artwork. Speech-bubble text is real HTML positioned with `site.css`; it is not baked into the images. Japanese artwork remains unchanged.

The dictionaries are trusted application resources. Values used for reports contain HTML, so translations should be reviewed as application code. Spreadsheet strings continue to be escaped by the comparison UI.

## Development

```sh
npm run build
npm run check
python -m http.server 8000
```

Open the Japanese root page or `/en/`. Tests cover translation completeness, static SEO output, reciprocal links, CSV comparison, ignored changes, header comparison, duplicate-key errors, interpolation and adding a third locale.

## Supported scope

XLSX values, types, formulas, standard cell formatting, merged cells, dimensions and sheet metadata are compared. Formulas are not recalculated. CSV values remain strings so leading zeros are preserved. Legacy XLS uses a compatibility library downloaded only when needed, without transmitting the selected file; its formatting comparison is limited to number formats.

Limits: 50 MB and 500,000 cells per file, and 2 million compared cell pairs per sheet. Unsupported features and safety limits are reported rather than silently dropping data.

## Search indexing

The language pages contain their headings, instructions and FAQ directly in HTML, without requiring JavaScript to display that content. The sitemap is `https://makiabe.github.io/ExcelDiff/sitemap.xml`. Search Console submission and search-engine indexing are separate from deployment; no indexing or ranking is guaranteed.
