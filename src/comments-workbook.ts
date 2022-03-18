import { program } from 'commander';
import ExcelJS, { Cell, Column, Workbook, Worksheet } from 'exceljs';
import { Connection } from 'oracledb';

export default class CommentsWorkbook {
  private readonly connection: Connection;
  private readonly workbook: Workbook;
  private tableNames: string[] | undefined;
  private titles: string[] | undefined;

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
    this.titles = result.rows
      .filter(row => row[1] === 'TABLE')
      .map(row => `${row[0].toString()}(${row[2].toString()})`);
    toc.addTable({
      name: 'toc',
      ref: 'A2',
      headerRow: true,
      columns: [{ name: '表名' }, { name: '类型' }, { name: '注释' }],
      rows: result.rows,
    });
    toc.mergeCells('A1:C1');
    toc.getCell('A1').value = '目录';
    this._setTableStyle(toc);
    toc.getColumn(1).eachCell?.((cell: Cell, rowNumber) => {
      if (rowNumber > 2) {
        cell.value = { text: cell.text, hyperlink: `#'${cell.text}'!A1` };
        cell.style.font = { underline: 'single', color: { argb: '5780C7' }, italic: true };
      }
    });
  }

  async createTables(): Promise<void> {
    if (this.tableNames !== undefined && this.titles !== undefined) {
      for (let i = 0; i < this.tableNames.length; i++) {
        await this._createTable(this.tableNames[i], this.titles[i]);
      }
    }
  }

  async outputWorkbook(): Promise<void> {
    await this.workbook.xlsx.writeFile(`${program.opts().username}.xlsx`);
    console.log(`${program.opts().username}.xlsx exported successfully`);
  }

  async _createTable(tableName: string, title: string): Promise<void> {
    console.log(`Exporting ${title} ...`);
    const sheet = this.workbook.addWorksheet(title);
    const result = await this.connection.execute<string[]>(
      `SELECT UTC.COLUMN_NAME, CASE WHEN DATA_PRECISION IS NOT NULL THEN DATA_TYPE || '(' || DATA_PRECISION || ',' || DATA_SCALE || ')' ELSE DATA_TYPE || '(' || DATA_LENGTH || ')' END, COMMENTS FROM USER_TAB_COLUMNS UTC LEFT JOIN USER_COL_COMMENTS UCC ON UTC.TABLE_NAME = UCC.TABLE_NAME AND UTC.COLUMN_NAME = UCC.COLUMN_NAME WHERE UTC.TABLE_NAME = '${tableName}' ORDER BY COLUMN_ID`
    );
    if (result.rows === undefined || result.rows.length === 0) {
      throw new Error(`${tableName} is no field`);
    }
    sheet.addTable({
      name: title,
      ref: 'A2',
      headerRow: true,
      columns: [{ name: '字段名' }, { name: '字段类型' }, { name: '注释' }],
      rows: result.rows,
    });
    sheet.mergeCells('A1:C1');
    sheet.getCell('A1').value = title;
    this._setTableStyle(sheet);
    sheet.getCell('D1').value = { text: '返回目录', hyperlink: `#'目录'!A1` };
    sheet.getCell('D1').style.font = { underline: 'single', color: { argb: '5780C7' }, italic: true };
  }

  private _setTableStyle(sheet: Worksheet): void {
    this._autosizeColumnCells(sheet.getColumn(1), (max: number) => max + 4);
    this._autosizeColumnCells(sheet.getColumn(2));
    this._autosizeColumnCells(sheet.getColumn(3), (max: number) => max * 1.5);
    sheet.getCell('A1').style.font = { bold: true, size: 15 };
    sheet.getCell('A1').style.alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.getRow(2).eachCell?.((cell: Cell) => {
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
