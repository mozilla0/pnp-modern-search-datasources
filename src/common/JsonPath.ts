// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type JsonObject = { [key: string]: any };

/**
 * Reads a value by path (a.b.c, items[0].x). Keys that contain dots themselves
 * (e.g. "@odata.count") are matched directly.
 */
export function getByPath(obj: JsonObject, path: string): any { // eslint-disable-line @typescript-eslint/no-explicit-any
  if (!obj || !path) { return undefined; }
  if (Object.prototype.hasOwnProperty.call(obj, path)) { return obj[path]; }

  const parts: string[] = path.replace(/\[(\d+)\]/g, '.$1').split('.').filter(p => p !== '');
  let current: any = obj; // eslint-disable-line @typescript-eslint/no-explicit-any
  for (let i = 0; i < parts.length; i++) {
    if (current === undefined || current === null) { return undefined; }
    const rest: string = parts.slice(i).join('.');
    if (typeof current === 'object' && Object.prototype.hasOwnProperty.call(current, rest)) {
      return current[rest];
    }
    current = current[parts[i]];
  }
  return current;
}

/** Returns the values of a field as strings (arrays and lookup/person objects are resolved). */
export function valuesOf(item: JsonObject, path: string): string[] {
  const value = getByPath(item, path);
  if (value === undefined || value === null || value === '') { return []; }
  const list = Array.isArray(value) ? value : [value];
  return list
    .map(v => {
      if (v && typeof v === 'object') {
        return v.title || v.Title || v.lookupValue || v.LookupValue || v.Label || v.label
          || v.displayName || v.name || JSON.stringify(v);
      }
      return String(v);
    })
    .filter(v => v !== '');
}

/** Returns the first match from several possible paths. */
export function firstByPaths(obj: JsonObject, paths: string[]): any { // eslint-disable-line @typescript-eslint/no-explicit-any
  for (const p of paths) {
    const v = getByPath(obj, p);
    if (v !== undefined && v !== null && v !== '') { return v; }
  }
  return undefined;
}
