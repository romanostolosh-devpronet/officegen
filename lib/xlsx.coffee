fs = require 'fs'
_  = require 'underscore'
mustache = require 'mustache'

baseobj  = require './basicgen.js'
msdoc    = require './msofficegen.js'

ObservableMixin = require './mixins/observableMixin'

# Extend require to load xml files as string
require.extensions['.xml'] = (module, filename)->
  module.exports = fs.readFileSync filename, 'utf8'

#Templates
appTemplate    = require './templates/app.xml'
sheetTemplate  = require './templates/sheet.xml'
stylesTemplate = require './templates/styles.xml'
workBookTemplate = require './templates/workbook.xml'
sharedStringsTemplate = require './templates/sharedStrings.xml'

class Xlsx

  # @mixes EventDispatcherMixin
  _.extend @prototype, ObservableMixin

  FONT_STYLES:
    'default': 0
    'normal' : 0
    'bold'   : 1

  TYPES_CODES:
    'number' : 'n'
    'string' : 's'
    'default': 's'

  constructor:->
    @name = null
    @data = null

  setCell:(name, value)->

  setRow:(number, value, style)->

  generate:(genobj, new_type, options, gen_private, type_info) ->

      #/
      #/ @brief Create the shared string resource.
      #/
      #/ This resource holding all the text strings of any Excel document.
      #/
      #/ @param[in] data Ignored by this callback function.
      #/ @return Text string.
      #/
      # TODO: Remove result variable
  cbMakeXlsSharedStrings = ->
    result = mustache.render(sharedStringsTemplate,
      xmlDocType: gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml()
      count: genobj.generate_data.total_strings
      uniqueCount: genobj.generate_data.shared_strings.length
      items: genobj.generate_data.shared_strings
    )
    result

  #/
  #/ @brief Prepare everything to generate XLSX files.
  #/
  #/ This method working on all the Excel cells to find out information needed by the generator engine.
  #/
  cbPrepareXlsxToGenerate = ->
    genobj.generate_data =
      shared_strings: []
      lookup_strings: {}
      total_strings: 0
      cell_strings: []

    gen_private.pages.forEach (currentPage, i) ->
      currentPageData = currentPage.sheet.data
      currentPageData.forEach (currentRow, rowId) ->
        currentRow.forEach (currentColumn, columnId) ->
          value = getCellValue(currentColumn)
          switch typeof value
          when "string"
            genobj.generate_data.total_strings++
            genobj.generate_data.cell_strings[i] = []  unless genobj.generate_data.cell_strings[i]
            genobj.generate_data.cell_strings[i][rowId] = []  unless genobj.generate_data.cell_strings[i][rowId]
            if value of genobj.generate_data.lookup_strings
              genobj.generate_data.cell_strings[i][rowId][columnId] = genobj.generate_data.lookup_strings[value]
            else
              shared_str_position = genobj.generate_data.shared_strings.length
              genobj.generate_data.cell_strings[i][rowId][columnId] = shared_str_position
              genobj.generate_data.lookup_strings[value] = shared_str_position
              genobj.generate_data.shared_strings[shared_str_position] = value

        return

      return

    if genobj.generate_data.total_strings
      gen_private.plugs.intAddAnyResourceToParse "xl\\sharedStrings.xml", "buffer", null, cbMakeXlsSharedStrings, false
      gen_private.type.msoffice.files_list.push
        name: "/xl/sharedStrings.xml"
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
        clear: "generate"

      gen_private.type.msoffice.rels_app.push
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        target: "sharedStrings.xml"
        clear: "generate"

    return

  #/
  #/ @brief ???.
  #/
  #/ ???.
  #/
  #/ @param[in] data Ignored by this callback function.
  #/ @return Text string.
  #/
  # TODO: Remove result variable
  cbMakeXlsStyles = (data) ->
    result = mustache.render(stylesTemplate,
      xmlDocType: gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data)
    )
    result

  #/
  #/ @brief ???.
  #/
  #/ ???.
  #/
  #/ @param[in] data Ignored by this callback function.
  #/ @return Text string.
  #/
  cbMakeXlsApp = (data) ->
    result = mustache.render(appTemplate,
      xmlDocType: gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data)
      userName: genobj.options.creator or "officegen"
      pagesCount: gen_private.pages.length
      sheets: ((totalPages) ->
        result = []
        i = 0

        while i < totalPages
          result.push i + 1
          i++
        result
      )(gen_private.pages.length)
    )
    result

  #/
  #/ @brief ???.
  #/
  #/ ???.
  #/
  #/ @param[in] data Ignored by this callback function.
  #/ @return Text string.
  #/
  cbMakeXlsWorkbook = (data) ->
    sheets = []
    gen_private.pages.forEach (currentPage, index) ->
      sheets.push
        name: currentPage.sheet.name or "Sheet" + (index + 1)
        sheetId: (index + 1)
        rId: currentPage.relId

      return

    result = mustache.render(workBookTemplate,
      xmlDocType: gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data)
      sheets: sheets
    )
    result

  #/
  #/ @brief Translate from the Excel displayed row name into index number.
  #/
  #/ ???.
  #/
  #/ @param[in] cell_string Either the cell displayed position or the row displayed position.
  #/ @return The cell's row Id.
  #/
  cbCellToNumber = (cell_string, ret_also_column) ->
    cellNumber = 0
    cellIndex = 0
    cellMax = cell_string.length
    rowId = 0

    # Converted from C++ (from DuckWriteC++):
    while cellIndex < cellMax
      curChar = cell_string.charCodeAt(cellIndex)
      if (curChar >= 0x30) and (curChar <= 0x39)
        rowId = parseInt(cell_string.slice(cellIndex), 10)
        rowId = (if (rowId > 0) then (rowId - 1) else 0)
        break
      else if (curChar >= 0x41) and (curChar <= 0x5A)
        if cellIndex > 0
          cellNumber++
          cellNumber *= (0x5B - 0x41)
        cellNumber += (curChar - 0x41)
      else if (curChar >= 0x61) and (curChar <= 0x7A)
        if cellIndex > 0
          cellNumber++
          cellNumber *= (0x5B - 0x41)
        cellNumber += (curChar - 0x61)
      cellIndex++
    if ret_also_column
      return (
        row: rowId
        column: cellNumber
      )
    cellNumber

  #/
  #/ @brief ???.
  #/
  #/ ???.
  #/
  #/ @param[in] cell_number ???.
  #/ @return ???.
  #/
  cbNumberToCell = (cell_number) ->
    outCell = ""
    curCell = cell_number
    while curCell >= 0
      outCell = String.fromCharCode((curCell % (0x5B - 0x41)) + 0x41) + outCell
      if curCell >= (0x5B - 0x41)
        curCell = Math.floor(curCell / (0x5B - 0x41)) - 1
      else
        break
    outCell
  getCellValue = (cell) ->
    return cell["value"]  if typeof cell is "object" and cell["value"]
    cell
  getCellStyleId = (cell) ->
    if typeof cell is "object" and cell["style"]
      style = cell["style"].toLowerCase()
      return FONT_STYLES[style]  if FONT_STYLES[style]
    FONT_STYLES["default"]

  #/
  #/ @brief ???.
  #/
  #/ ???.
  #/
  #/ @param[in] data The main sheet object.
  #/ @return Text string.
  #/
  cbMakeXlsSheet = (data) ->
    maxX = 0
    maxY = 0
    curColMax = undefined
    rowId = undefined
    columnId = undefined

    # Find the maximum cells area:
    maxY = (if data.sheet.data.length then (data.sheet.data.length - 1) else 0)
    rowId = 0
    total_size_y = data.sheet.data.length

    while rowId < total_size_y
      if data.sheet.data[rowId]
        curColMax = (if data.sheet.data[rowId].length then (data.sheet.data[rowId].length - 1) else 0)
        maxX = (if maxX < curColMax then curColMax else maxX)
      rowId++
    rows = []
    rowId = 0
    total_size_y = data.sheet.data.length

    while rowId < total_size_y
      if data.sheet.data[rowId]
        rowLines = 1
        data.sheet.data[rowId].forEach (cellData) ->
          if typeof cellData is "string"
            candidate = cellData.split("\n").length
            rowLines = Math.max(rowLines, candidate)
          return

        currentRow =
          columns: []
          rowId: rowId + 1
          height: rowLines * 15
          spansDimension: "1:" + data.sheet.data[rowId].length

        columnId = 0
        total_size_x = data.sheet.data[rowId].length

        while columnId < total_size_x
          cellData = getCellValue(data.sheet.data[rowId][columnId])
          if typeof cellData isnt "undefined"
            value = undefined
            type = TYPES_CODES["default"]
            cellValueType = typeof cellData
            type = TYPES_CODES[cellValueType]  if TYPES_CODES[cellValueType]
            switch typeof cellData
            when "number"
              value = cellData
            when "string"
              value = genobj.generate_data.cell_strings[data.id][rowId][columnId]
            currentColumn =
              cellName: cbNumberToCell(columnId) + (rowId + 1)
              type: type
              value: value
              styleId: getCellStyleId(data.sheet.data[rowId][columnId])

            currentRow.columns.push currentColumn
          columnId++
        rows.push currentRow
      rowId++
    result = mustache.render(sheetTemplate,
      xmlDocType: gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml()
      dimension: "A1:" + cbNumberToCell(maxX) + "" + (maxY + 1)
      rows: rows
    )
    result

  # Prepare genobj for MS-Office:
  msdoc.makemsdoc genobj, new_type, options, gen_private, type_info
  gen_private.plugs.type.msoffice.makeOfficeGenerator "xl", "workbook", {}
  gen_private.features.page_name = "sheets" # This document type must have pages.

  # On each generate we'll prepare the shared strings list:
  genobj.on "beforeGen", cbPrepareXlsxToGenerate
  gen_private.type.msoffice.files_list.push
    name: "/xl/styles.xml"
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
    clear: "type"
  ,
    name: "/xl/workbook.xml"
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    clear: "type"

  gen_private.type.msoffice.rels_app.push
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    target: "styles.xml"
    clear: "type"
  ,
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    target: "theme/theme1.xml"
    clear: "type"

  gen_private.plugs.intAddAnyResourceToParse "docProps\\app.xml", "buffer", null, cbMakeXlsApp, true
  gen_private.plugs.intAddAnyResourceToParse "xl\\styles.xml", "buffer", null, cbMakeXlsStyles, true
  gen_private.plugs.intAddAnyResourceToParse "xl\\workbook.xml", "buffer", null, cbMakeXlsWorkbook, true
  gen_private.plugs.intAddAnyResourceToParse "xl\\_rels\\workbook.xml.rels", "buffer", gen_private.type.msoffice.rels_app, gen_private.plugs.type.msoffice.cbMakeRels, true

  # ----- API for Excel documents: -----

  #/
  #/ @brief Create a new sheet.
  #/
  #/ This method creating a new Excel sheet.
  #/
  genobj.makeNewSheet = ->
    pageNumber = gen_private.pages.length
    sheetObj = data: []
    gen_private.pages[pageNumber] =
      id: pageNumber
      relId: gen_private.type.msoffice.rels_app.length + 1
      sheet: sheetObj

    gen_private.type.msoffice.rels_app.push
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
      target: "worksheets/sheet" + (pageNumber + 1) + ".xml"
      clear: "data"

    gen_private.type.msoffice.files_list.push
      name: "/xl/worksheets/sheet" + (pageNumber + 1) + ".xml"
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
      clear: "data"

    sheetObj.setCell = (position, data_val) ->
      rel_pos = cbCellToNumber(position, true)
      sheetObj.data[rel_pos.row] = []  unless sheetObj.data[rel_pos.row]
      sheetObj.data[rel_pos.row][rel_pos.column] = data_val
      return

    gen_private.plugs.intAddAnyResourceToParse "xl\\worksheets\\sheet" + (pageNumber + 1) + ".xml", "buffer", gen_private.pages[pageNumber], cbMakeXlsSheet, false
    sheetObj

  return

module.exports = Xlsx