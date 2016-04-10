var SocialCalc = require('socialcalc')
  , SpreadsheetColumn = require('spreadsheet-column')
  , column = new SpreadsheetColumn

const operationsList = [Set, InsertRow, DeleteRow, InsertCol, DeleteCol]
const operationsHash = operationsList.reduce(function(obj, op) {obj[(new op).type] = op; return obj},{})
const scInstance = new SocialCalc.SpreadsheetControl()

exports.create = function() {
  SocialCalc.ResetSheet(scInstance.sheet)
  var newScSnapshot = SocialCalc.CreateSheetSave(scInstance.sheet)
  return newScSnapshot
}

exports.apply = function(scSnapshot, ops) {
  // load snapshot into our global SocialCalc instance
  scInstance.sheet.ParseSheetSave(scSnapshot)

  // Turn ops into a command string
  var cmds = unpackOps(ops)
  .map((op) => op.serialize())
  .filter((op) => !!op)
  .forEach((cmd) => {
    var error = SocialCalc.ExecuteSheetCommand(scInstance.sheet, new SocialCalc.Parse(cmd), /*saveundo:*/false)
    if(error) throw new Error(error)
  })

  var newScSnapshot = SocialCalc.CreateSheetSave(scInstance.sheet)

  return newScSnapshot
}

exports.transform = function(ops1, ops2, side) {
  return unpackOps(ops1).map(function(op1) {
    unpackOps(ops2).forEach(function(op2) {
      op1.transformAgainst(op2, ('left'==side))
    })
    return op1
  })
}

exports.compose = function(ops1, ops2) {
  return ops1.concat(ops2)
}

exports.deserializeEdit = function(cmds) {
  return cmds.split('\n')
  .reduce((ops, cmd) => {
    var op 
    operationsList.some((Operation) => op = Operation.parse(cmd))
    if(op) op.forEach(op => ops.push(op)) // If nothing recognizes this, we filter it out.
    return ops
  }, [])
}

exports.serializeEdit = function(ops) {
  return unpackOps(ops).map((op) => op.serialize()).join('\n')
}

function unpackOps (ops) {
  return ops
  .map((op) => operationsHash[op.type].hydrate(op))
}


/**
----------------------------
OPERATIONS
----------------------------
*/
// These are SocialCalcs commands in the original format (we parse and serialize operations from7to this format):
//
//    set sheet attributename value (plus lastcol and lastrow)
//    set 22 attributename value
//    set B attributename value
//    set A1 attributename value1 value2... (see each attribute in code for details)
//    set A1:B5 attributename value1 value2...
//    erase/copy/cut/paste/fillright/filldown A1:B5 all/formulas/format
//    loadclipboard save-encoded-clipboard-data
//    clearclipboard
//    merge C3:F3
//    unmerge C3
//    insertcol/insertrow C5
//    deletecol/deleterow C5:E7
//    movepaste/moveinsert A1:B5 A8 all/formulas/format (if insert, destination must be in same rows or columns or else paste done)
//    sort cr1:cr2 col1 up/down col2 up/down col3 up/down
//    name define NAME definition
//    name desc NAME description
//    name delete NAME
//    recalc
//    redisplay
//    changedrendervalues
//    startcmdextension extension rest-of-command
//    sendemail ??? eddy ???

/**
All Operations implement the same Interface:

Operation#transformAgainst(op, side) : Operation // transforms this op against the passed on in-place. `side` is for tie-breaking.
Operation#serialize() : string // Returns the corresponding SocialCalc command (without newline)
Operation.parse(cmd:String) : Array<Operation>|false // Checks if the SocialCalc command is equivalent to the Operation type, if so: returns the corresponding operation(s), else it returns false.
Operation.hydrade(obj) : Operation // turns a plain object into an Operation instance
*/

/**
 * Set operation
 */
function Set(target, attribute, value) {
  this.type = 'Set'
  this.target = target
  this.attribute = attribute
  this.value = value
  this.hasEffect = true // Can become effectless upon transformation
}

Set.hydrate = function(obj) {
  return new Set(obj.target, obj.attribute, obj.value)
}

Set.prototype.transformAgainst = function(op2, side) {
  if(op2 instanceof Set) {
    if(op2.target !== this.target) return this
    if(side) {
      var obj = new Set(this.target, this.attribute, this.value)
      obj.hasEffect = false
      return obj
    }else return this
  }else if(op2 instanceof InsertRow && this.target !== 'sheet') {
    var otherRow = parseCell(op2.newRow)[1]
    // If this target is a cell
    if (this.target.match(/[a-z]+[0-9]+/)) {
      var myCell = parseCell(this.target)
        , thisRow = myCell[1]
      if (otherRow <= thisRow) return new Set(column.fromInt(myCell[0])+(thisRow+1), this.attribute, this.value)
      else return this
    }else
    // this target is a row
    if (parseInt(this.target) !== NaN) {
      var thisRow = parseInt(this.target)
      if (otherRow <= thisRow) return new Set(thisRow+1, this.attribute, this.value)
      else return this
    }
    // if this target is a column
    else return this
  }else
  if (op2 instanceof DeleteRow && this.target !== 'sheet'){
    var otherRow = parseCell(op2.row)[1]
    // If this target is a cell
    if (this.target.match(/[a-z]+[0-9]+/)) {
      var myCell = parseCell(this.target)
        , thisRow = myCell[1]
      if (otherCol < thisCol) return new Set(column.fromInt(myCell[0])+(thisRow-1), this.attribute, this.value)
      else if (otherRow === thisRow) {
        var obj = new Set(this.target, this.attribute, this.value)
	obj.hasEffect = false
	return obj
      }
      else return this
    }else
    // this target is a row
    if (parseInt(this.target) === NaN) {
      var thisRow = this.target
      if (otherRow <= thisRow) return new Set(String(thisRow-1), this.attribute, this.value)
      else if (otherCol === thisCol) {
        var obj = new Set(this.target, this.attribute, this.value)
	obj.hasEffect = false
	return obj
      }
      else return this
    }
    // if this target is a row
    else return this
  }
  if (op2 instanceof InsertCol && this.target !== 'sheet') {
    var otherCol = parseCell(op2.newCol)[0]
    // If this target is a cell
    if (this.target.match(/[a-z]+[0-9]+/)) {
      var myCell = parseCell(this.target)
        , thisCol = myCell[0]
      if (otherCol <= thisCol) return new Set(column.fromInt(thisCol+1)+myCell[1], this.attribute, this.value)
      else return this
    }else
    // this target is a col
    if (parseInt(this.target) === NaN) {
      var thisCol = column.fromStr(this.target)
      if (otherCol <= thisCol) return new Set(column.fromInt(thisCol+1), this.attribute, this.value)
      else return this
    }
    // if this target is a row
    else return this
  }else
  if (op2 instanceof DeleteCol && this.target !== 'sheet'){
    var otherCol = parseCell(op2.col)[0]
    // If this target is a cell
    if (this.target.match(/[a-z]+[0-9]+/)) {
      var myCell = parseCell(this.target)
        , thisCol = myCell[0]
      if (otherCol < thisCol) return new Set(column.fromInt(thisCol-1)+myCell[1], this.attribute, this.value)
      else if (otherCol === thisCol) {
        var obj = new Set(this.target, this.attribute, this.value)
	obj.hasEffect = false
	return obj
      }
      else return this
    }else
    // this target is a col
    if (parseInt(this.target) === NaN) {
      var thisCol = column.fromStr(this.target)
      if (otherCol <= thisCol) return new Set(column.fromInt(thisCol-1), this.attribute, this.value)
      else if (otherCol === thisCol) {
        var obj = new Set(this.target, this.attribute, this.value)
	obj.hasEffect = false
	return obj
      }
      else return this
    }
    // if this target is a row
    else return this
  }
  return this
}

Set.prototype.serialize = function() {
  if(!this.hasEffect) return ''
  return 'set '+this.target+' '+this.attribute+' '+this.value
}

Set.parse = function(cmdstr) {
  if(0 !== cmdstr.indexOf('set')) return
  var parts = cmdstr.split(' ')
    , cmd = parts[0]
    , target = parts[1]
    , attr = parts[2]
    , value = cmdstr.substr(cmd.length+1+target.length+1+attr.length+1)

  // if this a range?
  if(~target.indexOf(':')) {
    return resolveRange(target).map((target) => new Set(target, attr, value))
  }else {
    return [new Set(target, attr, value)]
  }
}

/**
 * InsertRow operation
 */
function InsertRow(newRow) {
  this.type = 'InsertRow'
  this.newRow = newRow
}

InsertRow.hydrate = function(obj) {
  return new InsertRow(obj.newRow)
}

InsertRow.prototype.transformAgainst = function(op, left) {
  if(op instanceof InsertRow) {
    var otherCell = parseCell(op.newRow)
     , myCell = parseCell(this.newRow)
    if (otherCell[1] < myCell[1]) {
      return new InsertRow(column.fromInt(myCell[0])+(myCell[1]+1))
    }else if (otherCell[1] === myCell[1]) {
      if(left) return new InsertRow(column.fromInt(myCell[0])+(myCell[1]+1))
      else return this
    }else{
      return this
    }
  }
  else
  if (op instanceof DeleteRow) {
    var otherRow = parseCell(op.row)[1]
      , mycell = parseCell(this.newRow)
    if (otherRow < myCell[1]) {
      return new DeleteCol(column.fromInt(myCell[0])+(myCell[1]-1))
    }else {
      return this
    }
  }
  return this
}

InsertRow.parse = function(cmd) {
  if(0 !== cmd.indexOf('insertrow ')) return false
  return [new InsertRow(cmd.substr('insertrow '.length))]
}

InsertRow.prototype.serialize = function() {
  return 'insertrow '+this.newRow
}

/**
 * DeleteRow operation
 */
function DeleteRow(row) {
  this.type = 'DeleteRow'
  this.row = row
}

DeleteRow.hydrate = function(obj) {
  return new DeleteRow(obj.row)
}

DeleteRow.prototype.transformAgainst = function(op, left) {
  if(op instanceof InsertRow) {
    var otherCell = parseCell(op.newRow)
     , myCell = parseCell(this.row)
    
    if (otherCell[1] === myCell[1]) {
      if(left) return new InsertCol(column.fromInt(myCell[0])+(myCell[1]+1))
      else return this
    }
    else if (otherCell[1] < myCell[1]) {
      return new InsertCol(column.fromInt(myCell[0])+(myCell[1]+1))
    }else{
      return this
    }
  }else
  if (op instanceof DeleteRow) {
    var otherRow = parseCell(op.row)[1]
      , mycell = parseCell(this.row)
    if (otherRow === myCell[1]) {
      // tie! break it!
      // If both happen to delete the same row, one wins, the other becomes a noop
      if (left) return new DeleteRow(null)
      else return this
    }
    else if (otherRow < myCell[1]) {
      return new DeleteCol(column.fromInt(myCell[0])+(myCell[1]-1))
    }else {
      return this
    }
  }
  return this
}

DeleteRow.parse = function(cmd) {
  if(0 !== cmd.indexOf('deleterow ')) return false
  var val = cmd.substr('deletecol '.length)
  if (~val.indexOf(':')) return resolveRowRange(val).map(cell => new DeleteRow(cell))
  return [new DeleteCol(val)]
}

DeleteRow.prototype.serialize = function() {
  if (!this.row) return ''
  return 'deleterow '+this.row
}

/**
 * InsertCol operation
 */
function InsertCol(newCol) {
  this.type = 'InsertCol'
  this.newCol = newCol
}

InsertCol.hydrate = function(obj) {
  return new InsertCol(obj.newCol)
}

InsertCol.prototype.transformAgainst = function(op, left) {
  if(op instanceof InsertCol) {
    var otherCell = parseCell(op.newRow)
     , myCell = parseCell(this.newRow)
    if (otherCell[0] === myCell[0]) {
      if(left) return new InsertCol(column.fromInt(myCell[0]+1)+myCell[1])
      else return this
    }
    else if (otherCell[0] < myCell[0]) {
      return new InsertCol(column.fromInt(myCell[0]+1)+myCell[1])
    }else{
      return this
    }
  }
  else if (op instanceof DeleteCol) {
    var otherCol = parseCell(op.col)[0]
      , mycell = parseCell(this.newCol)
    if (otherCol < myCell[0]) {
      return new InsertCol(column.fromInt(myCell[0]-1)+myCell[1])
    }else {
      return this
    }
  }
  return this
}

InsertCol.parse = function(cmd) {
  if(0 !== cmd.indexOf('insertcol ')) return false
  return [new InsertCol(cmd.substr('insertcol '.length))]
}

InsertCol.prototype.serialize = function() {
  return 'insertcol '+this.newCol
}

/**
 * DeleteCol operation
 */
function DeleteCol(col) {
  this.type = 'DeleteCol'
  this.col = col
}

DeleteCol.hydrate = function(obj) {
  return new DeleteCol(obj.col)
}

DeleteCol.prototype.transformAgainst = function(op, left) {
  if(op instanceof InsertCol) {
    var otherCell = parseCell(op.newRow)
     , myCell = parseCell(this.newRow)
    
    if (otherCell[0] === myCell[0]) {
      if(left) return new InsertCol(column.fromInt(myCell[0]+1)+myCell[1])
      else return this
    }
    else if (otherCell[0] < myCell[0]) {
      return new InsertCol(column.fromInt(myCell[0]+1)+myCell[1])
    }else{
      return this
    }
  }else
  if (op instanceof DeleteCol) {
    var otherCol = parseCell(op.col)[0]
      , mycell = parseCell(this.col)
    if (otherCol === myCell[0]) {
      // tie! break it!
      if (left) return new DeleteCol(null)
      else return this
    }
    else if (otherCol < myCell[0]) { // This now only catches 'less' (we caught equal already^^)
      return new DeleteCol(column.fromInt(myCell[0]-1)+myCell[1])
    }else {
      return this
    }
  }
  return this
}

DeleteCol.parse = function(cmd) {
  if(0 !== cmd.indexOf('deletecol ')) return false
  var val = cmd.substr('deletecol '.length)
  if (~val.indexOf(':')) return resolveColRange(val).map(cell => new DeleteCol(cell))
  return [new DeleteCol(val)]
}

DeleteCol.prototype.serialize = function() {
  if (!this.col) return ''
  return 'deletecol '+this.col
}

/**
 * Utility functions
 */

function parseCell(cell) {
  var match = cell.match(/([a-z]+)([0-9]+)/i)
  if(!match) throw new Error('invalid cell id '+cell)
  return [column.fromStr(match[1]), parseInt(match[2])]
}

function resolveRange(range) {
  if(!range.indexOf(':')) throw new Error('not a range.')
  var parts = range.split(':')
    , start = parseCell(parts[0])
    , end = parseCell(parts[1])
  var cells = []
  for (var i=start[0]; i<=end[0]; i++) {
    for (var j=start[1]; j <= end[1]; j++) {
      cells.push(column.fromInt(i)+j)
    }
  }
  return cells
}

function resolveColRange(range) {
  if(!range.indexOf(':')) throw new Error('not a range.')
  var parts = range.split(':')
    , start = parseCell(parts[0])
    , end = parseCell(parts[1])
  var cells = []
  for (var i=start[0]; i<=end[0]; i++) {
    cells.push(column.fromInt(i)+start[1])
  }
  cells.reverse()
  return cells
}
function resolveRowRange(range) {
  if(!range.indexOf(':')) throw new Error('not a range.')
  var parts = range.split(':')
    , start = parseCell(parts[0])
    , end = parseCell(parts[1])
    , col = column.fromInt(start[0])
  var cells = []
  for (var i=start[1]; i<=end[1]; i++) {
    cells.push(col+i)
  }
  cells.reverse()
  return cells
}
