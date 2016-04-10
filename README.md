# ot-socialcalc
This is a shareJS-compatible [OT](https://en-wikipedia.org/wiki/Operational_transformation) type. It defines Operations that relate to [socialcalc](https://npmjs.org/package/socialcalc)'s spreadsheet commands and can be serialized, applied on a socialcalc snapshot and, of course, transformed against each other.

## API
It implements shareJS' OT type interface

 * create() : snapshot
 * apply(op, snapshot): snapshot
 * transform(op1, op2, side): neOp
 * compose(op1, op2): newOP

Furthermore it allows you to convert a list of socialcalc commands into an operation and vice versa:

 * serializeEdit(op) : cmdStr
 * deserializeEdit(cmdStr) : op

## To-do
  - [x] set sheet operation
  - [x] set row
  - [x] set col
  - [x] set cell
  - [x] set cell range
  - [ ] erase/copy/cut/paste/fillright/filldown A1:B5 all/formulas/format
  - [ ] merge
  - [ ] unmerge
  - [x] insertcol/insertrow
  - [x] deletecol/deleterow C5:E7
  - [ ] movepaste/moveinsert A1:B5 A8 all/formulas/format (if insert, destination must be in same rows or columns or else paste done)
  - [ ] sort cr1:cr2 col1 up/down col2 up/down col3 up/down
  - [ ] name define|desc|delete


## Tests.
None yet. This is on the to-do!

## Legal
(c) 2016 by Marcel Klehr

Mozilla Public License 2.0 (see LICENSE.txt)
