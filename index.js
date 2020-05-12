const fse = require('fs-extra');
const path = require('path');
const ExcelJs = require('exceljs');

const ExcelLib = {
  loadFile: async ({ filePath, headers, emptyCellValue = undefined }) => {
    let buffer = await fse.readFile(filePath);
    let workbook = await new ExcelJs.Workbook().xlsx.load(buffer);
    let worksheet = workbook.getWorksheet(1);
    let lines = [];
    let rowOptions = { includeEmpty: false };
    let cellOptions = { includeEmpty: true };

    worksheet.eachRow(rowOptions, function (row) {
      let cells = [];

      row.eachCell(cellOptions, function (cell) {
        if (cell.value === null) {
          cell.value = emptyCellValue;
        }
        cells.push(cell.value);
      });

      lines.push(cells);
    });

    // TODO valid empty file
    if (lines.length === 0) {

    }

    if (headers) {
      let actual_headers = lines[0];
      // TODO valid headers

      let rows = lines.slice(1);
      if (typeof headers[0] === 'object' && headers[0].key) {
        rows = rows.map(row => {
          const new_row = {};

          for (let i in headers) {
            let cell = row[i];
            if (cell && typeof cell === 'object' && cell.result) {
              cell = cell.result;
            }
            new_row[headers[i].key] = cell;
          }

          return new_row;
        })
      }
      lines = rows
    }

    return lines;
  },

  init: async (config) => {
    const stream = await ExcelLib.Multi.init(config);
    const origin = {
      end: stream.end
    };
    stream.end = async function end(...args) {
      const files = await origin.end(...args);

      return files[0];
    }
    return stream;
  },

  Multi: {
    init: async ({ host, dir, fileName, workbook = {}, worksheet = {}, skipError = false, cloudAPI = 'export', autoDestroy = true, options }) => {
      await fse.ensureDir(dir);
      workbook.options = { useStyles: false, ...workbook.options };

      const Private = {
        eventHandlers: {
          error: [error => {
            Private.error = error;
            if (!Private.skipError && Private.autoDestroy) {
              Public.destroy();
            }
          }]
        },
        fileName, host, dir, cloudAPI, workbook, skipError, autoDestroy,
        writeCount: 0,
        files: [],
        curFile: null,
        worksheet: { name: 'Sheet 1', ...worksheet },
        createFile: () => {
          const fileName = Private.fileName.replace('{i}', Private.files.length + 1);

          const filePath = path.join(path.resolve(dir), fileName);

          let downloadLink = (path.join(Private.dir, fileName)).replace(/\\+/g, '/');
          if (host) { downloadLink = `${host}/${downloadLink}` }

          const stream = fse.createWriteStream(filePath, options);

          Private.addEventHandlers(stream);

          workbook = new ExcelJs.stream.xlsx.WorkbookWriter({ stream, ...Private.workbook.options });
          worksheet = workbook.addWorksheet(Private.worksheet.name, Private.worksheet.options);
          if (Array.isArray(Private.worksheet.columns)) {
            worksheet.columns = Private.worksheet.columns;
          }
          return { filePath, downloadLink, fileName, stream, workbook, worksheet };
        },
        addEventHandlers(stream) {
          for (let event in Private.eventHandlers) {
            stream.on(event, (...args) => Private.emit(event, ...args));
          }
        },
        emit(event, ...args) {
          if (Array.isArray(Private.eventHandlers[event]) && Private.eventHandlers[event].length > 0) {
            for (let handler of Private.eventHandlers[event]) {
              handler(...args);
            }
          }
        }
      }

      const Public = {
        write: (row) => {
          if (false) {
            Private.curFile.worksheet.commit();
            Private.curFile._endPromise = Private.curFile.workbook.commit();
            Private.writeCount = 0;
          }
          if (Private.writeCount == 0) {
            Private.curFile = Private.createFile();
            Private.files.push(Private.curFile);
          }

          Private.curFile.worksheet.addRow(row).commit();

          Private.writeCount++;

          return Public;
        },
        on(event, handler) {
          if (!Private.eventHandlers[event]) {
            Private.eventHandlers[event] = [];
          }
          Private.eventHandlers[event].push(handler);

          return Public;
        },
        end: () => {
          if (Private.curFile && !Private.curFile._endPromise) {
            Private.curFile._endPromise = Private.curFile.workbook.commit();
          }

          if (Private.cloudAPI) {

          }
          return Private.files;
        },
        async destroy(reason) {
          if (Private.files.length > 0) {
            return await Promise.all(Private.files.map(async file => {
              if (file.is_deleting || file.is_deleted) {
                return;
              }
              file.is_deleting = true;
              if (!(file.stream.destroyed || file.stream.closed || file.stream.writableFinished)) {
                file.stream.destroy(reason);
              }
              await fse.unlink(file.filePath);
              file.is_deleted = true;
            }));
          }
        }
      }
      return Public;
    }
  }
}
module.exports = { ExcelLib }

const test = {
  load: async () => {
    try {
      let filePath = path.resolve('./files/test.xlsx');
      let headers = [
        { header: 'SKU', key: 'sku' },
        { header: 'Tên sản phẩm', key: 'name' },
        { header: 'Ngày tạo', key: 'created_at' },
        { header: 'Giá', key: 'price' },
      ]
      let data = await ExcelLib.loadFile({ filePath, headers })
      console.log(data)
    }
    catch (e) {
      console.log(e)
    }
  },
  export: async () => {
    const moment = require('moment');
    const uuid = require('uuid');
    const excel = await ExcelLib.init({
      dir: `./download/${moment().year()}/${moment().format('M-DD')}`,
      fileName: `export-list-product-wait-process-file-{i}-${moment().utc(7).format('DD-MM-YYYY_HH-mm-ss')}-${uuid()}.xlsx`,
      worksheet: {
        name: 'sheet1',
        columns: [
          { header: 'SKU', key: 'sku', width: 20 },
          { header: 'Tên sản phẩm', key: 'name', width: 20 },
          { header: 'Giá', key: 'price', width: 20 },
          { header: 'Ngày tạo', key: 'created_at', width: 20 },
        ]
      },
      limit: 1000
    });

    await excel.write({ sku: 'test', name: 'san pham 1', price: 100000, created_at: 100000 });
    await excel.write({ sku: 'test2', name: 'san pham 2', price: 200000, created_at: 100000 });

    const { downloadLink } = await excel.end();
    console.log(downloadLink)
  }
}

// test.load()
// test.export()