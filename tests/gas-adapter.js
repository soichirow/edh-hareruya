import { readFileSync } from 'node:fs';
import { URL } from 'node:url';
import vm from 'node:vm';

const GAS_ROOT = new URL('../gas/', import.meta.url);

function load(file, globals = {}) {
  const context = vm.createContext({ console, ...globals });
  vm.runInContext(readFileSync(new URL(file, GAS_ROOT), 'utf8'), context, { filename: file });
  return context;
}

export const getJoined = (...args) => load('scryfall.js').getJoined(...args);
export const getImageNormal = (...args) => load('scryfall.js').getImageNormal(...args);

export function findImageCardPreferEn(card, fetchOptions, fetchJson) {
  const context = load('scryfall.js');
  context.fetchJsonWithRetry_ = fetchJson;
  return context.findImageCardPreferEn(card, fetchOptions);
}

export const convertISO8601ToTime = (...args) => load('youtube.js').convertISO8601ToTime(...args);
export const isOtakuCardVideo = (...args) => load('youtube.js').isOtakuCardVideo_(...args);

export function extractVideoData(item, Utilities) {
  return load('youtube.js', { Utilities }).extractVideoRow_(item);
}

export function createUpdateLatestVideos({ SpreadsheetApp, YouTube, Utilities, Logger }) {
  const context = load('youtube.js', { SpreadsheetApp, YouTube, Utilities, Logger });
  return { updateLatestVideos: context.newUpdate };
}

export function extractEpisodeNumber(title) {
  return load('youtube.js').extractEpisodeNumber_(title);
}

export function createFetchJsonWithRetry(UrlFetchApp, Utilities) {
  return load('scryfall.js', { UrlFetchApp, Utilities }).fetchJsonWithRetry_;
}

export function createFetchMtgCardDataJa({ SpreadsheetApp, UrlFetchApp, Utilities, now }) {
  const RuntimeDate = now ? class extends Date { static now() { return now(); } } : Date;
  return load('scryfall.js', { SpreadsheetApp, UrlFetchApp, Utilities, Date: RuntimeDate }).fetchMtgCardDataJa;
}

export function createDoGet({ SpreadsheetApp, ContentService }) {
  const context = load('app/doGet.js', { SpreadsheetApp, ContentService });
  return { doGet: context.doGet, fetchDatabaseJson: context.fetchDatabaseJson };
}
