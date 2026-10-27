import type { CellValue } from './index.js';
import type { SheetAction, TableRange } from './table.js';

type ChartAction = Extract<SheetAction, { type: 'chart' }>;
export interface ChartSummary {
  categories: (string | number)[];
  series: { name: string; values: number[] }[];
  sourceRows: number;
  ignoredBlankRows: number;
  convertedTextCells: number;
  trimmedCategories: number;
}
interface Decimal {
  units: bigint;
  scale: number;
}
function numberText(value: number): string {
  const raw = String(value);
  if (!/e/i.test(raw)) return raw;
  const [mantissa = '', exponent] = raw.toLowerCase().split('e');
  const sign = mantissa.startsWith('-') ? '-' : '';
  const parts = mantissa.replace(/^-/, '').split('.');
  const digits = parts.join('');
  const point = parts[0]!.length + Number(exponent);
  return (
    sign +
    (point <= 0
      ? '0.' + '0'.repeat(-point) + digits
      : point >= digits.length
        ? digits + '0'.repeat(point - digits.length)
        : digits.slice(0, point) + '.' + digits.slice(point))
  );
}

// Exact base-ten arithmetic is shared by bridge validation and the host. Never
// rewrite source cells or ask the model to calculate the chart's totals.
function decimal(value: CellValue): Decimal {
  let raw: string;
  if (typeof value === 'number' && Number.isFinite(value)) {
    raw = numberText(value);
  } else if (typeof value === 'string') {
    raw = value.trim().replace(/^'/, '').trim();
    if (!/^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(raw))
      throw new Error('金额须为十进制数字；不自动推断货币、百分比或单位。');
    raw = raw.replace(/,/g, '');
  } else throw new Error('金额缺失或不是数值，不能当作零。');
  const parts = raw.split('.');
  const fraction = (parts[1] ?? '').replace(/0+$/, '');
  const significant = (parts[0] + fraction)
    .replace(/^[+-]?0*/, '')
    .replace(/0+$/, '');
  if (significant.length > 15 || fraction.length > 15)
    throw new Error('金额超过宿主支持的 15 位有效数字或小数精度。');
  return {
    units: BigInt(parts[0] + fraction),
    scale: fraction.length,
  };
}
function add(a: Decimal, b: Decimal): Decimal {
  const scale = Math.max(a.scale, b.scale);
  return {
    units:
      a.units * BigInt(10) ** BigInt(scale - a.scale) +
      b.units * BigInt(10) ** BigInt(scale - b.scale),
    scale,
  };
}
function numeric(value: Decimal): number {
  const negative = value.units < BigInt(0);
  const digits = String(negative ? -value.units : value.units).padStart(
    value.scale + 1,
    '0',
  );
  const raw =
    (negative ? '-' : '') +
    (value.scale
      ? digits.slice(0, -value.scale) + '.' + digits.slice(-value.scale)
      : digits);
  decimal(raw); // Validate the sum's precision before crossing into native doubles.
  const result = Number(raw);
  if (!Number.isFinite(result)) throw new Error('汇总结果超出数值范围。');
  return result;
}
export function decimalNumber(value: CellValue): number {
  return numeric(decimal(value));
}
export function decimalSum(values: CellValue[]): number {
  return numeric(
    values.reduce<Decimal>((a, v) => add(a, decimal(v)), {
      units: BigInt(0),
      scale: 0,
    }),
  );
}
function quote(value: unknown): string {
  const text = String(value);
  if (/[\u0000-\u001f\u007f]/.test(text))
    throw new Error('图表分类或系列名称含不支持的控制字符。');
  return '"' + text.replace(/"/g, '""') + '"';
}
export function summarySeriesFormula(
  summary: ChartSummary,
  index: number,
): string {
  const series = summary.series[index];
  if (!series) throw new Error('图表系列无效。');
  const result =
    '=SERIES(' +
    quote(series.name) +
    ',{' +
    summary.categories.map(quote).join(',') +
    '},{' +
    series.values.map(numberText).join(',') +
    '},' +
    (index + 1) +
    ')';
  if (result.length > 8192)
    throw new Error('汇总图表系列超过 8192 字符，请缩小范围或减少分类。');
  return result;
}
export function summarizeChart(
  range: TableRange,
  action: ChartAction,
): ChartSummary {
  if (action.aggregation !== 'sum') throw new Error('未指定分组求和。');
  if (
    range.rowCount < 2 ||
    !Number.isSafeInteger(action.categoryColumn) ||
    action.categoryColumn < 0 ||
    action.categoryColumn >= range.columnCount ||
    !Array.isArray(action.valueColumns) ||
    !action.valueColumns.length ||
    action.valueColumns.some(
      (column) =>
        !Number.isSafeInteger(column) ||
        column < 0 ||
        column >= range.columnCount ||
        column === action.categoryColumn,
    )
  )
    throw new Error('分类或数值列超出本轮选区。');
  if (typeof BigInt !== 'function')
    throw new Error('当前宿主不支持精确汇总，请升级 WPS 内置浏览器。');
  const result: ChartSummary = {
    categories: [],
    series: action.valueColumns.map((column, index) => ({
      name: String(range.values[0]![column] ?? '系列 ' + (index + 1)),
      values: [],
    })),
    sourceRows: 0,
    ignoredBlankRows: 0,
    convertedTextCells: 0,
    trimmedCategories: 0,
  };
  const groups = new Map<string | number, Decimal[]>();
  let categoryType: string | undefined;
  const blank = (v: CellValue) =>
    v === null || (typeof v === 'string' && !v.trim());
  range.values.slice(1).forEach((row, index) => {
    const raw = row[action.categoryColumn]!;
    const values = action.valueColumns.map((column) => row[column]!);
    if (blank(raw) && values.every(blank)) {
      result.ignoredBlankRows++;
      return;
    }
    try {
      const category = typeof raw === 'string' ? raw.trim() : raw;
      if (
        (typeof category !== 'string' && typeof category !== 'number') ||
        category === '' ||
        (categoryType && categoryType !== typeof category)
      )
        throw new Error('分类缺失或类型混合，请先明确分类数据。');
      if (
        /^(合计|小计|总计|汇总|total|subtotal|grand total)$/i.test(
          String(category),
        )
      )
        throw new Error('分类包含合计或小计，请只选择明细，避免重复计算。');
      quote(category);
      categoryType = typeof category;
      const amounts = values.map(decimal);
      const previous = groups.get(category);
      groups.set(
        category,
        previous ? previous.map((sum, i) => add(sum, amounts[i]!)) : amounts,
      );
      result.sourceRows++;
      if (raw !== category) result.trimmedCategories++;
      result.convertedTextCells += values.filter(
        (v) => typeof v === 'string',
      ).length;
    } catch (error) {
      throw new Error(
        `选区第 ${index + 2} 行：${error instanceof Error ? error.message : String(error)}`,
      );
    }
  });
  if (!groups.size) throw new Error('没有可汇总的明细数据。');
  const entries = [...groups.entries()];
  if (action.categoryOrder === 'ascending')
    entries.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
  entries.forEach(([category, amounts]) => {
    result.categories.push(category);
    amounts.forEach((amount, i) =>
      result.series[i]!.values.push(numeric(amount)),
    );
  });
  result.series.forEach((_, i) => summarySeriesFormula(result, i));
  return result;
}
