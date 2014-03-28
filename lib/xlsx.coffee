fs = require 'fs'
_  = require 'underscore'
mustache = require 'mustache'

Sheet = require './sheet'
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

  XMLDOCTYPE: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

  FONT_STYLES:
    'default': 0
    'normal' : 0
    'bold'   : 1

  TYPES_CODES:
    'number' : 'n'
    'string' : 's'
    'default': 's'

  constructor:->
    @sheets = []

    @_sharedStrings = []
    @_totalStrings  = 0

    @creator = "test"

  createSheet:(name)->
    name?= "Sheet #{@sheets.length}"
    sheet = new Sheet name
    @sheets.push sheet

    sheet

  # Returns cell value
  # in case if value is object returns
  # value field
  _getCellValue: (cell) ->
    return cell["value"] if typeof cell is "object" and cell.value?
    cell

  _generateSharedStrings:->
    @sheets.forEach (sheet) =>
      sheetData = sheet.data

      sheetData.forEach (row) =>
        row.forEach (column) =>
          value = @_getCellValue column

          return unless typeof(value) is 'string'
          @_totalStrings++
          @_sharedStrings.push(value) if @_sharedStrings.indexOf(value) is -1

    result = mustache.render(sharedStringsTemplate,
      xmlDocType : @XMLDOCTYPE
      uniqueCount: @_sharedStrings.length
      count: @_totalStrings
      items: @_sharedStrings
    )

    result

  _generateRelations:->


  # Generates styles.xml
  # TODO: Remove result variable
  _generateXlsStyles:(data) ->
      result = mustache.render(stylesTemplate,
        xmlDocType: @XMLDOCTYPE
      )
      result

  #
  # Generates app.xml
  _generateXlsApp:(data) ->
      result = mustache.render(appTemplate,
        xmlDocType: @XMLDOCTYPE
        userName: @creator or "officegen"
        pagesCount: @sheets.length
        sheets: ((totalPages) ->
          result = []
          i = 0

          while i < totalPages
            result.push i + 1
            i++
          result
        )(@sheets.length)
      )

      result

  # @param[in] data Ignored by this callback function.
  # @return Text string.
  #
  _generateXlsWorkbook:() ->
    sheets = []
    @sheets.forEach (sheet, index) ->
      sheets.push
        name: sheet.name or "Sheet" + (index + 1)
        sheetId: (index + 1)
        rId: sheet.relId

      return

    result = mustache.render(workBookTemplate,
      xmlDocType: @XMLDOCTYPE
      sheets: sheets
    )
    result

  # Returns cell style
  # in case if value is object returns
  # style field
  # otherwise return default style
  _getCellStyleId: (cell) ->
    if typeof cell is "object" and cell.style?
      style = cell.style.toLowerCase()
      return @FONT_STYLES[style] if @FONT_STYLES[style]?

    @FONT_STYLES["default"]

  _generateXlsSheets: ->
    @sheets.forEach (sheet, index)=>
      fs.writeFileSync "./tmp/xl/worksheets/sheet#{index + 1}.xml", @_generateXlsSheet(sheet)
      # TODO: Add creating relations

  _generateXlsSheet: (sheet) ->
    xSize = 0
    ySize = 0
    rows = []

    # Find the maximum cells area:
    ySize = sheet.data.length - 1 if sheet.data.length

    sheet.data.forEach (row)->
      currentColumnSize = row.length - 1 if row.length
      xSize = Math.max currentColumnSize, xSize

    sheet.data.forEach (row, rowIndex)=>
      rowLines = 1

      currentRow =
        columns: []
        rowId: rowIndex + 1
        height: rowLines * 15
        spansDimension: "1:#{row.length}"

      row.forEach (column, columnIndex) =>
        columnValue = @_getCellValue column

        if typeof columnValue isnt "undefined"
          value = undefined
          type = @TYPES_CODES["default"]

          cellValueType = typeof columnValue
          type = @TYPES_CODES[cellValueType] if @TYPES_CODES[cellValueType]

          switch typeof columnValue

            when "number"
              value = columnValue

            when "string"
              # Calculate row height
              # depend on max lines in cell
              # TODO: move height calculation to another method
              candidate = columnValue.split("\n").length
              rowLines = Math.max rowLines, candidate
              currentRow.height = rowLines * 15

              value = @_sharedStrings.indexOf columnValue

          currentColumn =
            cellName: "#{Sheet::numberToCell(columnIndex)}#{rowIndex + 1}"
            type: type
            value: value
            styleId: @_getCellStyleId(columnValue)

          currentRow.columns.push currentColumn

        rows.push currentRow

    result = mustache.render(sheetTemplate,
      xmlDocType: @XMLDOCTYPE
      dimension: "A1:#{Sheet::numberToCell(xSize)}#{(ySize + 1)}"
      rows: rows
    )

    result

  generate:() ->

    # TODO: Temporary folders
    fs.mkdirSync "./tmp"       unless fs.existsSync "./tmp"
    fs.mkdirSync "./tmp/xl"    unless fs.existsSync "./tmp/xl"
    fs.mkdirSync "./tmp/_rels" unless fs.existsSync "./tmp/_rels"
    fs.mkdirSync "./tmp/docProps" unless fs.existsSync "./tmp/docProps"
    fs.mkdirSync "./tmp/xl/worksheets" unless fs.existsSync "./tmp/xl/worksheets"

    fs.writeFileSync "./tmp/docProps/app.xml", @_generateXlsApp()
    fs.writeFileSync "./tmp/xl/sharedStrings.xml", @_generateSharedStrings()
    fs.writeFileSync "./tmp/xl/styles.xml",  @_generateXlsStyles()
    fs.writeFileSync "./tmp/xl/workbook.xml", @_generateXlsWorkbook()

    @_generateXlsSheets()

    #TODO: Clarify how to generate this
    #gen_private.plugs.intAddAnyResourceToParse('xl\\_rels\\workbook.xml.rels', 'buffer', gen_private.type.msoffice.rels_app, gen_private.plugs.type.msoffice.cbMakeRels, true );


module.exports = Xlsx

# test
xlsxDocument = new Xlsx()

# First sheet
sheet = xlsxDocument.createSheet()

sheet.name = "Excel Test"

# The direct option - two-dimensional array:
sheet.data[0] = []
sheet.data[0][0] = 1
sheet.data[1] = []
sheet.data[1][3] =
  value: "abc"
  style: "BOLD"

sheet.data[1][4] =
  value: "More"
  style: "bOld"

sheet.data[1][5] = "Text\n test"
sheet.data[1][6] = "Here"
sheet.data[2] = []
sheet.data[2][5] = "abc"
sheet.data[2][6] = 900
sheet.data[6] = []
sheet.data[6][2] = 1972

# Using setCell:
sheet.setCell "E7", 340
sheet.setCell "I1", -3
sheet.setCell "I2", 31.12
sheet.setCell "G102",
  value: "Hello World!"
  style: "bold"

# Second sheet

sheet2 = xlsxDocument.createSheet()

sheet2.name = "Excel My"

# The direct option - two-dimensional array:
sheet2.data[0] = []
sheet2.data[0][0] = 1
sheet2.data[1] = []
sheet2.data[1][3] =
  value: "abc"
  style: "BOLD"

sheet2.data[1][4] =
  value: "More 1"
  style: "bOld"

sheet2.data[1][5] = "Text 2"
sheet2.data[1][6] = "Here 2"
sheet2.data[2] = []
sheet2.data[2][5] = "abc 4"
sheet2.data[2][6] = 900444
sheet2.data[6] = []
sheet2.data[6][2] = 19724213

# Using setCell:
sheet2.setCell "E7", 3404
sheet2.setCell "I1", -34
sheet2.setCell "I2", 31.125
sheet2.setCell "G102",
  value: "Hello World!"
  style: "bold"

xlsxDocument.generate()

#sheet.setRow ( '3', [] , {style:'bold'});
