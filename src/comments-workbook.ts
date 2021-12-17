import { program } from 'commander';
import ExcelJS, { Cell, Column, Workbook, Worksheet } from 'exceljs';
import { Connection } from 'oracledb';

export default class CommentsWorkbook {
  private readonly connection: Connection;
  private readonly workbook: Workbook;
  private tableNames: string[] | undefined;

  constructor(connection: Connection) {
    this.connection = connection;
    this.workbook = new ExcelJS.Workbook();
  }

  async oc2e(): Promise<void> {
    await this.createToc();
    await this.createTables();
    await this.outputWorkbook();
  }

  async createToc(): Promise<void> {
    console.log('Exporting toc ...');
    const toc = this.workbook.addWorksheet('目录');
    const result = await this.connection.execute<string[]>(
      'SELECT TABLE_NAME, TABLE_TYPE, COMMENTS FROM USER_TAB_COMMENTS ORDER BY TABLE_TYPE, TABLE_NAME'
    );
    if (result.rows === undefined || result.rows.length === 0) {
      throw new Error('USER_TAB_COMMENTS is empty');
    }
    this.tableNames = result.rows.filter(row => row[1] === 'TABLE').map(row => row[0].toString());
    toc.addTable({
      name: 'toc',
      ref: 'A1',
      headerRow: true,
      columns: [{ name: '表名' }, { name: '类型' }, { name: '注释' }],
      rows: result.rows,
    });
    this._setTableStyle(toc);
    toc.getColumn(1).eachCell?.((cell: Cell, rowNumber) => {
      if (rowNumber !== 1) {
        cell.value = { text: cell.text, hyperlink: `#'${cell.text}'!A1` };
        cell.style.font = { underline: 'single', color: { argb: '5780C7' }, italic: true };
      }
    });
  }

  async createTables(): Promise<void> {
    if (this.tableNames !== undefined) {
      for (const tableName of this.tableNames) {
        await this._createTable(tableName);
      }
    }
  }

  async outputWorkbook(): Promise<void> {
    await this.workbook.xlsx.writeFile(`${program.opts().username}.xlsx`);
    console.log(`${program.opts().username}.xlsx exported successfully`);
  }

  async _createTable(tableName: string): Promise<void> {
    console.log(`Exporting ${tableName} ...`);
    const sheet = this.workbook.addWorksheet(tableName);
    const result = await this.connection.execute<string[]>(
      `SELECT UTC.COLUMN_NAME, DATA_TYPE || '(' || DATA_LENGTH || ')', COMMENTS FROM USER_TAB_COLUMNS UTC LEFT JOIN USER_COL_COMMENTS UCC ON UTC.TABLE_NAME = UCC.TABLE_NAME AND UTC.COLUMN_NAME = UCC.COLUMN_NAME WHERE UTC.TABLE_NAME = '${tableName}' ORDER BY COLUMN_ID`
    );
    if (result.rows === undefined || result.rows.length === 0) {
      throw new Error(`${tableName} is no field`);
    }
    sheet.addTable({
      name: tableName,
      ref: 'A1',
      headerRow: true,
      columns: [{ name: '字段名' }, { name: '字段类型' }, { name: '注释' }],
      rows: result.rows,
    });
    this._setTableStyle(sheet);
    sheet.getCell('D1').value = { text: '返回目录', hyperlink: `#'目录'!A1` };
    sheet.getCell('D1').style.font = { underline: 'single', color: { argb: '5780C7' }, italic: true };
  }

  private _setTableStyle(sheet: Worksheet): void {
    this._autosizeColumnCells(sheet.getColumn(1), (max: number) => max + 4);
    this._autosizeColumnCells(sheet.getColumn(2));
    this._autosizeColumnCells(sheet.getColumn(3), (max: number) => max * 1.5);
    sheet.getRow(1).eachCell?.((cell: Cell) => {
      cell.style.font = { color: { argb: 'EDECF9' }, bold: true, size: 14 };
      cell.style.alignment = { vertical: 'middle', horizontal: 'center' };
    });
  }

  private _autosizeColumnCells(col: Column, calculate = (max: number) => max, minWidth = 12): void {
    const dataMax: number[] = [];
    col.eachCell?.((cell: Cell) => {
      dataMax.push(cell.value?.toString().length || 0);
    });
    const max = Math.max(...dataMax);
    col.width = max < minWidth ? minWidth : calculate(max);
  }
}
