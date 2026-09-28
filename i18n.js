/* UI and report translations are compiled from locales/*.json, never spreadsheet data. */
(() => {
  'use strict';
  const data = globalThis.ExcelDiffLocale;
  if (!data || !data.messages) throw new Error('Missing Excel Diff locale data.');
  const messages = Object.freeze(data.messages);
  function t(key, values = {}) {
    if (!Object.hasOwn(messages, key)) throw new Error(`Missing translation: ${key}`);
    return messages[key].replace(/\{(p\d+|locale)\}/g, (_, name) => {
      if (name === 'locale') return data.locale;
      if (!Object.hasOwn(values, name)) throw new Error(`Missing placeholder ${name} in ${key}`);
      return String(values[name] ?? '');
    });
  }
  globalThis.ExcelDiffI18n = Object.freeze({locale: data.locale, t});
})();
