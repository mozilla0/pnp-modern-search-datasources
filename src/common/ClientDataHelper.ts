import {
  FilterComparisonOperator,
  FilterConditionOperator,
  IDataFilter,
  IDataFilterConfiguration,
  IDataFilterResult,
  IDataFilterResultValue,
  IDataFilterValue
} from '@pnp/modern-search-extensibility';
import { JsonObject, valuesOf } from './JsonPath';

export interface ISortSpec { field: string; descending: boolean; }

export function compareValues(a: string, b: string): number {
  const na: number = Number(a);
  const nb: number = Number(b);
  if (a !== '' && b !== '' && !isNaN(na) && !isNaN(nb)) { return na - nb; }
  const da: number = Date.parse(a);
  const db: number = Date.parse(b);
  if (/^\d{4}-\d{2}-\d{2}/.test(a) && /^\d{4}-\d{2}-\d{2}/.test(b) && !isNaN(da) && !isNaN(db)) { return da - db; }
  return a.localeCompare(b);
}

function matches(itemValues: string[], filterValue: IDataFilterValue): boolean {
  const target: string = String(filterValue.value);
  const op: FilterComparisonOperator = filterValue.operator !== undefined ? filterValue.operator : FilterComparisonOperator.Eq;

  if (op === FilterComparisonOperator.Neq) {
    return itemValues.every(v => v !== target);
  }
  return itemValues.some(v => {
    switch (op) {
      case FilterComparisonOperator.Contains: return v.toLowerCase().indexOf(target.toLowerCase()) !== -1;
      case FilterComparisonOperator.Gt: return compareValues(v, target) > 0;
      case FilterComparisonOperator.Geq: return compareValues(v, target) >= 0;
      case FilterComparisonOperator.Lt: return compareValues(v, target) < 0;
      case FilterComparisonOperator.Leq: return compareValues(v, target) <= 0;
      default: return v === target;
    }
  });
}

export function applyClientFilters(items: JsonObject[], filters: IDataFilter[], operator: FilterConditionOperator): JsonObject[] {
  if (!filters || filters.length === 0) { return items; }
  return items.filter(item => {
    const results: boolean[] = filters.map(filter => {
      const itemValues: string[] = valuesOf(item, filter.filterName);
      const checks: boolean[] = (filter.values || []).map(v => matches(itemValues, v));
      if (checks.length === 0) { return true; }
      return filter.operator === FilterConditionOperator.AND ? checks.every(c => c) : checks.some(c => c);
    });
    return operator === FilterConditionOperator.OR ? results.some(r => r) : results.every(r => r);
  });
}

/**
 * Computes filter values with counts. For each filter the other selected filters are
 * applied so the counts match the current selection.
 */
export function buildFilterResults(
  items: JsonObject[], configs: IDataFilterConfiguration[], applied: IDataFilter[], operator: FilterConditionOperator
): IDataFilterResult[] {
  return (configs || []).map(cfg => {
    const others: IDataFilter[] = (applied || []).filter(f => f.filterName !== cfg.filterName);
    const base: JsonObject[] = applyClientFilters(items, others, operator);

    const counts: { [value: string]: number } = {};
    base.forEach(item => valuesOf(item, cfg.filterName).forEach(v => { counts[v] = (counts[v] || 0) + 1; }));

    let values: IDataFilterResultValue[] = Object.keys(counts).map(v => ({ name: v, value: v, count: counts[v] }));
    values.sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
    if (cfg.maxBuckets && cfg.maxBuckets > 0) { values = values.slice(0, cfg.maxBuckets); }
    return { filterName: cfg.filterName, values };
  });
}

export function applyClientSort(items: JsonObject[], sort: ISortSpec[]): JsonObject[] {
  if (!sort || sort.length === 0) { return items; }
  return items.slice().sort((a, b) => {
    for (const s of sort) {
      const result: number = compareValues(valuesOf(a, s.field)[0] || '', valuesOf(b, s.field)[0] || '');
      if (result !== 0) { return s.descending ? -result : result; }
    }
    return 0;
  });
}
