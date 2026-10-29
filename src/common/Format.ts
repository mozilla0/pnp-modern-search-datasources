/** Replaces {0}, {1} … in a localized string */
export function format(text: string, ...args: (string | number)[]): string {
  return (text || '').replace(/\{(\d+)\}/g, (match: string, index: string) => {
    const value = args[parseInt(index, 10)];
    return value !== undefined ? String(value) : match;
  });
}
