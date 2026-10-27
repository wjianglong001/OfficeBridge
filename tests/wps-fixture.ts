import { readFile } from 'node:fs/promises';
import { createContext, runInContext } from 'node:vm';

/** Mutable WPS test double, never used by the app or real Agent. */
export function createWpsModel(
  initial: Record<string, unknown[][]> = {
    Data: [
      ['name', 'amount'],
      [' a ', 10],
      ['b', 20],
    ],
  },
) {
  const sheets: any[] = [];
  const app: any = {
    DisplayAlerts: true,
    ActiveWorkbook: null,
    ActiveSheet: null,
    Selection: null,
  };
  const book: any = { Name: 'test.xlsx', FullName: 'E:\\fixtures\\test.xlsx' };
  book.Names = { Count: 0 };
  book.ProtectStructure = false;
  function letters(column: number) {
    let value = '';
    while (column) {
      value = String.fromCharCode(65 + ((column - 1) % 26)) + value;
      column = Math.floor((column - 1) / 26);
    }
    return value;
  }
  function shiftedFormula(
    formula: string,
    owner: any,
    target: any,
    start: number,
    count: number,
    insert: boolean,
    rows = false,
  ) {
    return formula.replace(
      /"(?:[^"]|"")*"|(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))!)?(\$?)([A-Z]{1,3})(\$?\d+)/g,
      (match, quoted, name, dollar, column, row) => {
        if (match.startsWith('"')) return match;
        if (
          quoted || name ? (quoted || name) !== target.Name : owner !== target
        )
          return match;
        let index = rows ? Number(row.replace('$', '')) : 0;
        if (!rows)
          for (const char of column)
            index = index * 26 + char.charCodeAt(0) - 64;
        if (index < start) return match;
        if (!insert && index < start + count) return '#REF!';
        return (
          (quoted || name ? (quoted ? "'" + quoted + "'" : name) + '!' : '') +
          dollar +
          (rows ? column : letters(index + (insert ? count : -count))) +
          (rows
            ? (row.startsWith('$') ? '$' : '') +
              (index + (insert ? count : -count))
            : row)
        );
      },
    );
  }
  function makeRange(
    row: number,
    column: number,
    rows: number,
    columns: number,
  ) {
    return {
      Row: row,
      Column: column,
      Rows: { Count: rows },
      Columns: { Count: columns },
      Areas: { Count: 1 },
    };
  }
  function sheet(name: string, data: unknown[][]) {
    const cells = new Map<string, any>();
    const columnProperties = new Map<
      number,
      { width: number; hidden: boolean }
    >();
    const rowProperties = new Map<
      number,
      { height: number; hidden: boolean }
    >();
    const chartList: any[] = [];
    const result: any = {
      Name: name,
      ProtectContents: false,
      Rows: { Count: 1048576 },
      Columns: { Count: 16384 },
      Shapes: { Count: 0 },
      ListObjects: { Count: 0 },
      PivotTables: { Count: 0 },
      AutoFilterMode: false,
      Calculate() {},
    };
    function get(row: number, column: number) {
      const key = row + ':' + column;
      if (!cells.has(key)) {
        const cell: any = {
          _value: null,
          _formula: null,
          NumberFormat: 'General',
          Font: { Bold: false },
          Interior: { Color: 16777215, Pattern: -4142 },
          MergeCells: false,
          HasArray: false,
          failNextWrite: false,
          writes: 0,
          ClearContents() {
            cell.Value2 = null;
          },
        };
        Object.defineProperties(cell, {
          Value2: {
            configurable: true,
            get() {
              if (!cell._formula) {
                if (cell._value === null || cell._value === '') {
                  if (app.blankReadback === 'null') return null;
                  if (app.blankReadback === 'empty-string') return '';
                  if (app.blankReadback === 'undefined') return undefined;
                }
                return cell._value;
              }
              if (cell._formula === '=""') return '';
              if (cell._formula === '=1/0') return '#DIV/0!';
              if (cell._formula.includes('#REF!')) return '#REF!';
              const sum = /^=SUM\(([A-Z]+)(\d+):\1(\d+)\)$/i.exec(
                cell._formula,
              );
              if (sum) {
                let total = 0;
                const column = Array.from(sum[1]!.toUpperCase()).reduce(
                  (n, letter) => n * 26 + letter.charCodeAt(0) - 64,
                  0,
                );
                for (let r = Number(sum[2]); r <= Number(sum[3]); r++)
                  total += Number(get(r, column).Value2 ?? 0);
                return total;
              }
              return 42;
            },
            set(value) {
              cell.writes++;
              if (result.failNextCellWrite) {
                result.failNextCellWrite = false;
                throw new Error('simulated inserted-column write failure');
              }
              if (cell.failNextWrite) {
                cell.failNextWrite = false;
                throw new Error('simulated write failure');
              }
              cell._formula = null;
              cell._value = value;
              if (
                typeof value === 'string' &&
                value.startsWith('=') &&
                cell.NumberFormat !== '@'
              ) {
                cell._formula = value;
                cell._value = null;
              }
            },
          },
          Formula: {
            get() {
              return cell._formula ?? cell._value;
            },
            set(value) {
              cell._formula = value;
              cell._value = null;
            },
          },
          HasFormula: {
            get() {
              return cell._formula !== null;
            },
          },
          Text: {
            get() {
              return String(cell.Value2 ?? '');
            },
          },
        });
        cells.set(key, cell);
      }
      return cells.get(key);
    }
    result.Cells = { Item: get };
    result._cells = cells;
    const charts: any = {
      get Count() {
        return chartList.length;
      },
      Item(key: number | string) {
        const object =
          typeof key === 'number'
            ? chartList[key - 1]
            : chartList.find((item) => item.Name === key);
        if (!object) throw new Error('Chart not found');
        return object;
      },
      Add(left: number, top: number, width: number, height: number) {
        if (result.failNextChartAdd) {
          result.failNextChartAdd = false;
          throw new Error('chart add failed');
        }
        const seriesList: any[] = [];
        const collection: any = {
          get Count() {
            return seriesList.length;
          },
          Item(index: number) {
            return seriesList[index - 1];
          },
          NewSeries() {
            if (result.failNextSeries) {
              result.failNextSeries = false;
              throw new Error('series failed');
            }
            const item: any = {
              Formula: '=SERIES()',
              Delete() {
                seriesList.splice(seriesList.indexOf(item), 1);
              },
            };
            seriesList.push(item);
            return item;
          },
        };
        const object: any = {
          Name: 'Chart ' + (chartList.length + 1),
          Left: left,
          Top: top,
          Width: width,
          Height: height,
          Chart: {
            ChartType: 51,
            HasTitle: false,
            HasLegend: true,
            ChartStyle: 1,
            ChartTitle: { Text: '' },
            SeriesCollection: () => collection,
          },
          Delete() {
            if (result.failNextChartDelete) {
              result.failNextChartDelete = false;
              throw new Error('chart delete failed');
            }
            chartList.splice(chartList.indexOf(object), 1);
          },
        };
        chartList.push(object);
        if (result.failChartAddAfterMutation) {
          result.failChartAddAfterMutation = false;
          throw new Error('partial chart add');
        }
        return object;
      },
    };
    result.ChartObjects = () => charts;
    let extraShapes = 0;
    result.Shapes = {
      get Count() {
        return chartList.length + extraShapes;
      },
      set Count(value: number) {
        extraShapes = value - chartList.length;
      },
    };
    let currentName = name;
    Object.defineProperty(result, 'Name', {
      configurable: true,
      get() {
        return currentName;
      },
      set(value: string) {
        if (result.failNextRename) {
          result.failNextRename = false;
          throw new Error('rename failed');
        }
        if (
          sheets.some(
            (item) =>
              item !== result &&
              item.Name.toLowerCase() === value.toLowerCase(),
          ) ||
          /[\\/?*\[\]:]/.test(value)
        )
          throw new Error('invalid sheet name');
        const old = currentName;
        const pattern = new RegExp(
          "(?:'" +
            old.replace(/'/g, "''").replace(/[.*+?^${}()|[\]\\]/g, '\\$&') +
            "'|" +
            old.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') +
            ')!',
          'g',
        );
        currentName = value;
        for (const owner of sheets) {
          const replacement = /^[A-Za-z_][A-Za-z0-9_]*$/.test(value)
            ? value + '!'
            : "'" + value.replace(/'/g, "''") + "'!";
          for (const cell of owner._cells.values())
            if (cell._formula)
              cell._formula = cell._formula.replace(pattern, replacement);
          const collection = owner.ChartObjects();
          for (let i = 1; i <= collection.Count; i++) {
            const series = collection.Item(i).Chart.SeriesCollection();
            for (let j = 1; j <= series.Count; j++)
              series.Item(j).Formula = series
                .Item(j)
                .Formula.replace(pattern, replacement);
          }
        }
        if (result.failRenameAfterMutation) {
          result.failRenameAfterMutation = false;
          throw new Error('partial rename');
        }
      },
    });
    function rowProps(row: number) {
      if (!rowProperties.has(row))
        rowProperties.set(row, { height: 15, hidden: false });
      return rowProperties.get(row)!;
    }
    let filter: any = null;
    Object.defineProperties(result, {
      AutoFilterMode: {
        configurable: true,
        get() {
          return filter !== null;
        },
        set(value: boolean) {
          if (!value && filter) {
            for (
              let r = filter.Range.Row + 1;
              r < filter.Range.Row + filter.Range.Rows.Count;
              r++
            )
              rowProps(r).hidden = false;
            filter = null;
          } else if (value && !filter) filter = {};
        },
      },
      AutoFilter: {
        get() {
          return filter;
        },
      },
    });
    const filterMatch = (
      value: any,
      criteria: any,
      operator = 0,
      second?: any,
    ): boolean => {
      if (operator === 7 && Array.isArray(criteria))
        return criteria.some(
          (item) =>
            String(item).toLowerCase() === String(value ?? '').toLowerCase(),
        );
      const match = /^(>=|<=|<>|>|<|=)?([\s\S]*)$/.exec(String(criteria))!;
      const operation = match[1] ?? '=',
        literal = match[2]!;
      let found: boolean;
      if (['>', '>=', '<', '<='].includes(operation)) {
        const n = Number(literal);
        found =
          typeof value === 'number' &&
          (operation === '>'
            ? value > n
            : operation === '>='
              ? value >= n
              : operation === '<'
                ? value < n
                : value <= n);
      } else {
        let pattern = '';
        for (let i = 0; i < literal.length; i++) {
          let ch = literal[i]!;
          if (ch === '~' && i + 1 < literal.length) {
            ch = literal[++i]!;
            pattern += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          } else
            pattern +=
              ch === '*'
                ? '.*'
                : ch === '?'
                  ? '.'
                  : ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        }
        found = new RegExp('^' + pattern + '$', 'i').test(String(value ?? ''));
        if (operation === '<>') found = !found;
      }
      return second === undefined
        ? found
        : operator === 2
          ? found || filterMatch(value, second)
          : found && filterMatch(value, second);
    };
    result.Range = (address: string) => {
      const rowMatch = /^(\d+):(\d+)$/.exec(address);
      if (rowMatch) {
        const start = Number(rowMatch[1]),
          end = Number(rowMatch[2]),
          count = end - start + 1;
        const target: any = {
          get RowHeight() {
            return app.hiddenHeightZero && rowProps(start).hidden
              ? 0
              : rowProps(start).height;
          },
          set RowHeight(value: number) {
            for (let r = start; r <= end; r++) rowProps(r).height = value;
          },
          get Hidden() {
            return rowProps(start).hidden;
          },
          set Hidden(value: boolean) {
            for (let r = start; r <= end; r++) rowProps(r).hidden = value;
          },
          Insert() {
            shiftRows(true);
          },
          Delete() {
            shiftRows(false);
          },
        };
        function shiftRows(insert: boolean) {
          if (result.failNextRowOperation) {
            result.failNextRowOperation = false;
            throw new Error('row API failure');
          }
          const moved = new Map<string, any>();
          for (const [key, cell] of cells) {
            const [r, c] = key.split(':').map(Number) as [number, number];
            if (!insert && r >= start && r <= end) continue;
            const destination = r >= start ? r + (insert ? count : -count) : r;
            moved.set(destination + ':' + c, cell);
          }
          cells.clear();
          for (const [key, cell] of moved) cells.set(key, cell);
          const props = new Map<number, { height: number; hidden: boolean }>();
          for (const [r, item] of rowProperties) {
            if (!insert && r >= start && r <= end) continue;
            props.set(r >= start ? r + (insert ? count : -count) : r, item);
          }
          rowProperties.clear();
          for (const [r, item] of props) rowProperties.set(r, item);
          for (const owner of sheets)
            for (const cell of owner._cells.values())
              if (cell._formula)
                cell._formula = shiftedFormula(
                  cell._formula,
                  owner,
                  result,
                  start,
                  count,
                  insert,
                  true,
                );
          if (result.failRowAfterMutation) {
            result.failRowAfterMutation = false;
            throw new Error('partial row mutation');
          }
        }
        target.EntireRow = target;
        return target;
      }
      const cellMatch = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(address);
      if (cellMatch) {
        const num = (s: string) =>
          [...s].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
        const top = Number(cellMatch[2]),
          left = num(cellMatch[1]!),
          bottom = Number(cellMatch[4] ?? cellMatch[2]),
          right = num(cellMatch[3] ?? cellMatch[1]!);
        const range: any = {
          ...makeRange(top, left, bottom - top + 1, right - left + 1),
          Left: (left - 1) * 60,
          Top: (top - 1) * 15,
          Width: (right - left + 1) * 60,
          AutoFilter(
            field?: number,
            criteria?: any,
            operator = 0,
            second?: any,
          ) {
            if (!filter) {
              const conditions = Array.from(
                { length: right - left + 1 },
                () => ({ On: false }),
              );
              filter = {
                Range: range,
                Filters: {
                  Count: conditions.length,
                  Item(index: number) {
                    return conditions[index - 1];
                  },
                },
                _conditions: conditions,
              };
            }
            if (field !== undefined) {
              filter._conditions[field - 1] = {
                On: true,
                Criteria1: criteria,
                Operator: operator,
                ...(second !== undefined ? { Criteria2: second } : {}),
              };
            }
            for (let r = top + 1; r <= bottom; r++)
              rowProps(r).hidden = !filter._conditions.every(
                (item: any, index: number) =>
                  !item.On ||
                  filterMatch(
                    get(r, left + index).Value2,
                    item.Criteria1,
                    item.Operator,
                    item.Criteria2,
                  ),
              );
            if (result.failNextFilter) {
              result.failNextFilter = false;
              throw new Error('filter partial failure');
            }
          },
        };
        return range;
      }
      const match = /^([A-Z]+):([A-Z]+)$/.exec(address);
      if (!match) throw new Error('Fixture only supports whole-column ranges');
      const number = (value: string) =>
        [...value].reduce((sum, c) => sum * 26 + c.charCodeAt(0) - 64, 0);
      const start = number(match[1]!),
        end = number(match[2]!),
        count = end - start + 1;
      function props(column: number) {
        if (!columnProperties.has(column))
          columnProperties.set(column, { width: 8.43, hidden: false });
        return columnProperties.get(column)!;
      }
      function shift(insert: boolean) {
        if (result.failNextColumnOperation) {
          result.failNextColumnOperation = false;
          throw new Error('simulated column API failure');
        }
        const moved = new Map<string, any>();
        for (const [key, cell] of cells) {
          const [row, column] = key.split(':').map(Number) as [number, number];
          if (!insert && column >= start && column <= end) continue;
          const destination =
            column >= start ? column + (insert ? count : -count) : column;
          if (destination > 16384) throw new Error('Column overflow');
          moved.set(row + ':' + destination, cell);
        }
        cells.clear();
        for (const [key, cell] of moved) cells.set(key, cell);
        const widths = new Map<number, { width: number; hidden: boolean }>();
        for (const [column, property] of columnProperties) {
          if (!insert && column >= start && column <= end) continue;
          const destination =
            column >= start ? column + (insert ? count : -count) : column;
          if (destination <= 16384) widths.set(destination, property);
        }
        columnProperties.clear();
        for (const [column, property] of widths)
          columnProperties.set(column, property);
        for (const owner of sheets)
          for (const cell of owner._cells.values())
            if (cell._formula)
              cell._formula = shiftedFormula(
                cell._formula,
                owner,
                result,
                start,
                count,
                insert,
              );
        result.columnOperations = (result.columnOperations ?? 0) + 1;
        if (result.failColumnAfterMutation) {
          result.failColumnAfterMutation = false;
          throw new Error('simulated partial native mutation');
        }
      }
      return {
        get ColumnWidth() {
          return app.hiddenWidthZero && props(start).hidden
            ? 0
            : props(start).width;
        },
        set ColumnWidth(value) {
          for (let c = start; c <= end; c++) props(c).width = value;
        },
        get Hidden() {
          return props(start).hidden;
        },
        set Hidden(value) {
          for (let c = start; c <= end; c++) props(c).hidden = value;
        },
        Insert() {
          shift(true);
        },
        Delete() {
          shift(false);
        },
      };
    };
    result.Activate = () => {
      app.ActiveSheet = result;
    };
    result.Delete = () => {
      if (sheets.length <= 1) throw new Error('Cannot delete last worksheet');
      const index = sheets.indexOf(result);
      if (index < 0) throw new Error('Sheet not found');
      sheets.splice(index, 1);
      if (app.ActiveSheet === result) app.ActiveSheet = sheets[0];
    };
    Object.defineProperty(result, 'UsedRange', {
      get() {
        let minRow = Infinity,
          minColumn = Infinity,
          maxRow = 0,
          maxColumn = 0;
        for (const [key, cell] of cells) {
          if (
            cell._value === null &&
            !cell._formula &&
            cell.NumberFormat === 'General' &&
            !cell.Font.Bold &&
            cell.Interior.Pattern === -4142
          )
            continue;
          const [row, column] = key.split(':').map(Number) as [number, number];
          minRow = Math.min(minRow, row);
          minColumn = Math.min(minColumn, column);
          maxRow = Math.max(maxRow, row);
          maxColumn = Math.max(maxColumn, column);
        }
        return maxRow
          ? makeRange(
              minRow,
              minColumn,
              maxRow - minRow + 1,
              maxColumn - minColumn + 1,
            )
          : makeRange(1, 1, 1, 1);
      },
    });
    data.forEach((row, r) =>
      row.forEach((value, c) => {
        get(r + 1, c + 1).Value2 = value;
      }),
    );
    sheets.push(result);
    return result;
  }
  const collection: any = {
    Item(key: string | number) {
      const found =
        typeof key === 'number'
          ? sheets[key - 1]
          : sheets.find((item) => item.Name === key);
      if (!found) throw new Error('Worksheet not found: ' + key);
      return found;
    },
    Add() {
      let index = 1;
      while (sheets.some((item) => item.Name === 'Sheet' + index)) index++;
      const created = sheet('Sheet' + index, []);
      app.ActiveSheet = created;
      return created;
    },
  };
  Object.defineProperty(collection, 'Count', {
    get() {
      return sheets.length;
    },
  });
  book.Worksheets = collection;
  Object.entries(initial).forEach(([name, values]) => sheet(name, values));
  app.ActiveWorkbook = book;
  app.ActiveSheet = sheets[0];
  app.Selection = makeRange(1, 1, 3, 2);
  return {
    app,
    book,
    sheets,
    select: (row: number, column: number, rows = 1, columns = 1) => {
      app.Selection = makeRange(row, column, rows, columns);
    },
  };
}

export async function loadWpsHost(initial?: Record<string, unknown[][]>) {
  const model = createWpsModel(initial);
  const window: any = {};
  const context = createContext({
    Application: model.app,
    window,
    Blob,
    TextEncoder,
    setTimeout,
  });
  for (const file of [
    '.generated/contract.js',
    'host.js',
    'host-operations.js',
  ])
    runInContext(
      await readFile(
        new URL('../apps/wps-addin/' + file, import.meta.url),
        'utf8',
      ),
      context,
    );
  return {
    ...model,
    host: window.OfficeBridgeHost,
    clone: <T>(value: T): T => JSON.parse(JSON.stringify(value)),
  };
}
