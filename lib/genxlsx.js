//
// officegen: All the code to generate XLSX files.
//
// Please refer to README.md for this module's documentations.
//
// NOTE:
// - Before changing this code please refer to the hacking the code section on README.md.
//
// Copyright (c) 2013 Ziv Barber;
//
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// 'Software'), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
// IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
// CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
// TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
// SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//

var fs = require('fs'),
    mustache = require('mustache'),
    baseobj  = require('./basicgen.js'),
    msdoc    = require('./msofficegen.js');

// Extend require to load xml files as string
require.extensions['.xml'] = function (module, filename) {
  module.exports = fs.readFileSync(filename, 'utf8');
};

// Templates
var appTemplate = require('./templates/app.xml'),
    sheetTemplate  = require('./templates/sheet.xml'),
    stylesTemplate = require('./templates/styles.xml'),
    workBookTemplate = require('./templates/workbook.xml'),
    sharedStringsTemplate = require('./templates/sharedStrings.xml');

var FONT_STYLES = {
  'default': 0,
  'normal': 0,
  'bold'  : 1
};

var TYPES_CODES = {
  'number': 'n',
  'string': 's',
  'default': 's'
};

///
/// @brief Extend officegen object with XLSX support.
///
/// This method extending the given officegen object to create XLSX document.
///
/// @param[in] genobj The object to extend.
/// @param[in] new_type The type of object to create.
/// @param[in] options The object's options.
/// @param[in] gen_private Access to the internals of this object.
/// @param[in] type_info Additional information about this type.
///
function makeXlsx (genobj, new_type, options, gen_private, type_info) {
  ///
  /// @brief Create the shared string resource.
  ///
  /// This resource holding all the text strings of any Excel document.
  ///
  /// @param[in] data Ignored by this callback function.
  /// @return Text string.
  ///
  // TODO: Remove result variable
  function cbMakeXlsSharedStrings() {
    var result = mustache.render(sharedStringsTemplate, {
      xmlDocType:gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(),
      count: genobj.generate_data.total_strings,
      uniqueCount: genobj.generate_data.shared_strings.length,
      items: genobj.generate_data.shared_strings
    });

    return result;
  }

  ///
  /// @brief Prepare everything to generate XLSX files.
  ///
  /// This method working on all the Excel cells to find out information needed by the generator engine.
  ///
  function cbPrepareXlsxToGenerate () {

    genobj.generate_data = {
      shared_strings: [],
      lookup_strings: {},
      total_strings : 0,
      cell_strings  : []
    };

    gen_private.pages.forEach(function(currentPage, i) {
      var currentPageData = currentPage.sheet.data;

      currentPageData.forEach(function(currentRow, rowId) {
        currentRow.forEach(function (currentColumn, columnId) {
          var value = getCellValue(currentColumn);
          switch (typeof value) {
            case 'string':
              genobj.generate_data.total_strings++;

              if (!genobj.generate_data.cell_strings[i]) {
                genobj.generate_data.cell_strings[i] = [];
              }

              if (!genobj.generate_data.cell_strings[i][rowId]) {
                genobj.generate_data.cell_strings[i][rowId] = [];
              }

              if (value in genobj.generate_data.lookup_strings) {
                genobj.generate_data.cell_strings[i][rowId][columnId] = genobj.generate_data.lookup_strings[value];

              } else {
                var shared_str_position = genobj.generate_data.shared_strings.length;

                genobj.generate_data.cell_strings[i][rowId][columnId] = shared_str_position;
                genobj.generate_data.lookup_strings[value]            = shared_str_position;
                genobj.generate_data.shared_strings[shared_str_position] = value;
              }

              break;
          }
        });
      });
    });

    if (genobj.generate_data.total_strings) {
      gen_private.plugs.intAddAnyResourceToParse('xl\\sharedStrings.xml', 'buffer', null, cbMakeXlsSharedStrings, false);
      gen_private.type.msoffice.files_list.push({
        name: '/xl/sharedStrings.xml',
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
        clear: 'generate'
      });

      gen_private.type.msoffice.rels_app.push({
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
        target: 'sharedStrings.xml',
        clear: 'generate'
      });
    }
  }

  ///
  /// @brief ???.
  ///
  /// ???.
  ///
  /// @param[in] data Ignored by this callback function.
  /// @return Text string.
  ///
  // TODO: Remove result variable
  function cbMakeXlsStyles (data) {
    var result = mustache.render(stylesTemplate, {
      xmlDocType: gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data)
    });

    return result;
  }

  ///
  /// @brief ???.
  ///
  /// ???.
  ///
  /// @param[in] data Ignored by this callback function.
  /// @return Text string.
  ///
  function cbMakeXlsApp ( data ) {
    var result = mustache.render(appTemplate, {
      xmlDocType: gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data),
      userName  : genobj.options.creator || 'officegen',
      pagesCount: gen_private.pages.length,
      sheets: (function (totalPages) {
        result = []

        for(var i = 0; i < totalPages; i++) {
          result.push(i + 1);
        }

        return result;
      })(gen_private.pages.length)
    });

    return result;
  }

  ///
  /// @brief ???.
  ///
  /// ???.
  ///
  /// @param[in] data Ignored by this callback function.
  /// @return Text string.
  ///
  function cbMakeXlsWorkbook (data) {
    var sheets = [];

    gen_private.pages.forEach(function(currentPage, index) {
      sheets.push({
        name: currentPage.sheet.name || 'Sheet' + (index + 1),
        sheetId: (index + 1),
        rId: currentPage.relId
      });
    });

    var result = mustache.render(workBookTemplate, {
      xmlDocType: gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data),
      sheets: sheets
    });

    return result;
  }

  ///
  /// @brief Translate from the Excel displayed row name into index number.
  ///
  /// ???.
  ///
  /// @param[in] cell_string Either the cell displayed position or the row displayed position.
  /// @return The cell's row Id.
  ///
  function cbCellToNumber ( cell_string, ret_also_column ) {
    var cellNumber = 0;
    var cellIndex = 0;
    var cellMax = cell_string.length;
    var rowId = 0;

    // Converted from C++ (from DuckWriteC++):
    while (cellIndex < cellMax) {
      var curChar = cell_string.charCodeAt (cellIndex);
      if ((curChar >= 0x30) && (curChar <= 0x39)) {
        rowId = parseInt(cell_string.slice(cellIndex), 10);
        rowId = (rowId > 0) ? (rowId - 1) : 0;
        break;

      } else if ((curChar >= 0x41) && (curChar <= 0x5A)) {
        if (cellIndex > 0) {
          cellNumber++;
          cellNumber *= (0x5B-0x41);
        }

        cellNumber += (curChar - 0x41);

      } else if ((curChar >= 0x61) && (curChar <= 0x7A)) {
        if (cellIndex > 0) {
          cellNumber++;
          cellNumber *= (0x5B-0x41);
        }

        cellNumber += (curChar - 0x61);
      }

      cellIndex++;
    }

    if (ret_also_column) {
      return {
        row: rowId,
        column: cellNumber
      };
    }

    return cellNumber;
  }

  ///
  /// @brief ???.
  ///
  /// ???.
  ///
  /// @param[in] cell_number ???.
  /// @return ???.
  ///
  function cbNumberToCell (cell_number) {
    var outCell = '';
    var curCell = cell_number;

    while (curCell >= 0) {
      outCell = String.fromCharCode((curCell % (0x5B-0x41)) + 0x41 ) + outCell;
      if (curCell >= (0x5B-0x41)) {
        curCell = Math.floor(curCell / (0x5B-0x41)) - 1;
      } else {
        break;
      }
    }

    return outCell;
  }

  function getCellValue(cell) {
    if (typeof cell === 'object'
        && cell['value']) {
      return cell['value'];
    }

    return cell;
  }

  function getCellStyleId(cell) {
    if (typeof cell === 'object' && cell['style']) {

      var style = cell['style'].toLowerCase();
      if (FONT_STYLES[style]) {
        return FONT_STYLES[style];
      }
    }

    return FONT_STYLES['default'];
  }

  ///
  /// @brief ???.
  ///
  /// ???.
  ///
  /// @param[in] data The main sheet object.
  /// @return Text string.
  ///
  function cbMakeXlsSheet (data) {
    var maxX = 0,
        maxY = 0,
        curColMax,
        rowId,
        columnId;

    // Find the maximum cells area:
    maxY = data.sheet.data.length ? (data.sheet.data.length - 1) : 0;

    for (var rowId = 0, total_size_y = data.sheet.data.length; rowId < total_size_y; rowId++) {
      if ( data.sheet.data[rowId] ) {
        curColMax = data.sheet.data[rowId].length ? (data.sheet.data[rowId].length - 1) : 0;
        maxX = maxX < curColMax ? curColMax : maxX;
      }
    }

    var rows = [];

    for (var rowId = 0, total_size_y = data.sheet.data.length; rowId < total_size_y; rowId++) {
      if (data.sheet.data[rowId]) {
        var rowLines = 1;
        data.sheet.data[rowId].forEach(function (cellData) {
          if (typeof cellData === 'string') {
            var candidate = cellData.split('\n').length;
            rowLines = Math.max(rowLines, candidate);
          }
        });

        var currentRow = {
          columns: [],
          rowId:   rowId + 1,
          height:  rowLines * 15,
          spansDimension: '1:' + data.sheet.data[rowId].length
        };

        for (var columnId = 0, total_size_x = data.sheet.data[rowId].length; columnId < total_size_x; columnId++) {
          var cellData = getCellValue(data.sheet.data[rowId][columnId]);
          if (typeof  cellData !== 'undefined') {
            var value,
                type = TYPES_CODES['default'];

            var cellValueType = typeof cellData;

            if (TYPES_CODES[cellValueType]) {
              type = TYPES_CODES[cellValueType];
            }

            switch (typeof cellData) {
              case 'number':
                value = cellData;
                break;

              case 'string':
                value = genobj.generate_data.cell_strings[data.id][rowId][columnId];
                break;
            }

            var currentColumn = {
              cellName: cbNumberToCell(columnId) + (rowId + 1),
              type: type,
              value: value,
              styleId: getCellStyleId(data.sheet.data[rowId][columnId])
            };

            currentRow.columns.push(currentColumn);
          }
        }

        rows.push(currentRow);
      }
    }

    var result = mustache.render(sheetTemplate, {
      xmlDocType: gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(),
      dimension : 'A1:' + cbNumberToCell(maxX) + '' + (maxY + 1),
      rows: rows
    });

    return result;
  }

  // Prepare genobj for MS-Office:
  msdoc.makemsdoc ( genobj, new_type, options, gen_private, type_info );
  gen_private.plugs.type.msoffice.makeOfficeGenerator ('xl', 'workbook', {});

  gen_private.features.page_name = 'sheets'; // This document type must have pages.

  // On each generate we'll prepare the shared strings list:
  genobj.on('beforeGen', cbPrepareXlsxToGenerate);

  gen_private.type.msoffice.files_list.push({
      name: '/xl/styles.xml',
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
      clear: 'type'
    },
    {
      name: '/xl/workbook.xml',
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
      clear: 'type'
    }
  );

  gen_private.type.msoffice.rels_app.push({
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
      target: 'styles.xml',
      clear: 'type'
    },
    {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
      target: 'theme/theme1.xml',
      clear: 'type'
    }
  );

  gen_private.plugs.intAddAnyResourceToParse('docProps\\app.xml', 'buffer', null, cbMakeXlsApp, true );
  gen_private.plugs.intAddAnyResourceToParse('xl\\styles.xml', 'buffer', null, cbMakeXlsStyles, true );
  gen_private.plugs.intAddAnyResourceToParse('xl\\workbook.xml', 'buffer', null, cbMakeXlsWorkbook, true );
  gen_private.plugs.intAddAnyResourceToParse('xl\\_rels\\workbook.xml.rels', 'buffer', gen_private.type.msoffice.rels_app, gen_private.plugs.type.msoffice.cbMakeRels, true );


  // ----- API for Excel documents: -----

  ///
  /// @brief Create a new sheet.
  ///
  /// This method creating a new Excel sheet.
  ///
  genobj.makeNewSheet = function () {
    var pageNumber = gen_private.pages.length,
        sheetObj = {
           data: []
        };

    gen_private.pages[pageNumber] = {
      id: pageNumber,
      relId: gen_private.type.msoffice.rels_app.length + 1,
      sheet: sheetObj
    };

    gen_private.type.msoffice.rels_app.push ({
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
      target: 'worksheets/sheet' + (pageNumber + 1) + '.xml',
      clear: 'data'
    });

    gen_private.type.msoffice.files_list.push({
      name: '/xl/worksheets/sheet' + (pageNumber + 1) + '.xml',
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      clear: 'data'
    });

    sheetObj.setCell = function (position, data_val) {
      var rel_pos = cbCellToNumber (position, true);

      if (!sheetObj.data[rel_pos.row]) {
        sheetObj.data[rel_pos.row] = [];
      }

      sheetObj.data[rel_pos.row][rel_pos.column] = data_val;
    };

    gen_private.plugs.intAddAnyResourceToParse ( 'xl\\worksheets\\sheet' + (pageNumber + 1) + '.xml', 'buffer', gen_private.pages[pageNumber], cbMakeXlsSheet, false );

    return sheetObj;
  };
}

baseobj.plugins.registerDocType('xlsx', makeXlsx, {}, baseobj.docType.SPREADSHEET, "Microsoft Excel Document");

