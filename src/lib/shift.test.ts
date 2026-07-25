import assert from 'node:assert/strict'
import { test } from 'node:test'

import { type ShiftSpec, shiftFormula } from './shift.js'

const insertRows = (at: number, onCurrentSheet = true, editedSheet = 'Sheet1'): ShiftSpec => ({
  axis: 'row',
  at,
  delta: 1,
  editedSheet,
  onCurrentSheet,
})
const insertColumns = (at: number, onCurrentSheet = true, editedSheet = 'Sheet1'): ShiftSpec => ({
  axis: 'column',
  at,
  delta: 1,
  editedSheet,
  onCurrentSheet,
})
const deleteRows = (
  at: number,
  count = 1,
  onCurrentSheet = true,
  editedSheet = 'Sheet1',
): ShiftSpec => ({ axis: 'row', at, delta: -count, editedSheet, onCurrentSheet })

test('shiftFormula moves a row reference at or past the insertion point', () => {
  assert.equal(shiftFormula('A5', insertRows(3)), 'A6')
  assert.equal(shiftFormula('A3', insertRows(3)), 'A4')
  assert.equal(shiftFormula('A2', insertRows(3)), 'A2')
})

test('shiftFormula preserves absolute markers on both parts', () => {
  assert.equal(shiftFormula('$A$5', insertRows(3)), '$A$6')
  assert.equal(shiftFormula('$A5', insertRows(3)), '$A6')
  assert.equal(shiftFormula('A$5', insertRows(3)), 'A$6')
})

test('shiftFormula shifts each end of a range independently', () => {
  assert.equal(shiftFormula('SUM(A5:A10)', insertRows(3)), 'SUM(A6:A11)')
  assert.equal(shiftFormula('SUM(A1:A10)', insertRows(3)), 'SUM(A1:A11)')
})

test('shiftFormula leaves an unqualified reference alone when the formula lives elsewhere', () => {
  assert.equal(shiftFormula('A5', insertRows(3, false)), 'A5')
})

test('shiftFormula shifts a reference qualified with the edited sheet from any sheet', () => {
  assert.equal(shiftFormula('Sheet1!A5', insertRows(3, false)), 'Sheet1!A6')
  assert.equal(shiftFormula('sheet1!A5', insertRows(3, false)), 'sheet1!A6')
  assert.equal(shiftFormula('Sheet3!A5', insertRows(3, false)), 'Sheet3!A5')
})

test('shiftFormula matches a quoted sheet name', () => {
  assert.equal(shiftFormula("'My Sheet'!A5", insertRows(3, false, 'My Sheet')), "'My Sheet'!A6")
  assert.equal(shiftFormula("'Other'!A5", insertRows(3, false, 'My Sheet')), "'Other'!A5")
})

test('shiftFormula never touches a string literal', () => {
  assert.equal(shiftFormula('"A5"', insertRows(3)), '"A5"')
  assert.equal(shiftFormula('IF(A5="A5",A5,0)', insertRows(3)), 'IF(A6="A5",A6,0)')
})

test('shiftFormula leaves function names and defined names alone', () => {
  assert.equal(shiftFormula('SUM(A5)', insertRows(3)), 'SUM(A6)')
  assert.equal(shiftFormula('LOG10(A5)', insertRows(3)), 'LOG10(A6)')
  assert.equal(shiftFormula('myRange+A5', insertRows(3)), 'myRange+A6')
  assert.equal(shiftFormula('A5B+A5', insertRows(3)), 'A5B+A6')
})

test('shiftFormula leaves a table structured reference alone', () => {
  assert.equal(shiftFormula('Table1[Amount]+A5', insertRows(3)), 'Table1[Amount]+A6')
})

test('shiftFormula shifts columns and re-letters them', () => {
  assert.equal(shiftFormula('A5', insertColumns(1)), 'B5')
  assert.equal(shiftFormula('SUM(A5:C5)', insertColumns(2)), 'SUM(A5:D5)')
  assert.equal(shiftFormula('$C$5', insertColumns(2)), '$D$5')
})

test('shiftFormula shifts whole-row and whole-column ranges on the matching axis', () => {
  assert.equal(shiftFormula('SUM(5:5)', insertRows(3)), 'SUM(6:6)')
  assert.equal(shiftFormula('SUM(A:A)', insertColumns(1)), 'SUM(B:B)')
  assert.equal(shiftFormula('SUM(A:A)', insertRows(3)), 'SUM(A:A)')
  assert.equal(shiftFormula('SUM(5:5)', insertColumns(1)), 'SUM(5:5)')
})

test('shiftFormula turns a reference pushed past the sheet into #REF!', () => {
  assert.equal(shiftFormula('A1048576', insertRows(1)), '#REF!')
  assert.equal(shiftFormula('XFD5', insertColumns(1)), '#REF!')
})

test('shiftFormula shifts references below a deleted row up', () => {
  assert.equal(shiftFormula('A10', deleteRows(3)), 'A9')
  assert.equal(shiftFormula('A2', deleteRows(3)), 'A2')
})

test('shiftFormula turns a single cell on a deleted row into #REF!', () => {
  assert.equal(shiftFormula('A3', deleteRows(3)), '#REF!')
  assert.equal(shiftFormula('SUM(A3,B1)', deleteRows(3)), 'SUM(#REF!,B1)')
})

test('shiftFormula drops the sheet qualifier from a destroyed reference', () => {
  assert.equal(shiftFormula('Sheet1!A3', deleteRows(3, 1, false)), '#REF!')
  assert.equal(shiftFormula("'Sheet1'!A3:A3", deleteRows(3, 1, false)), '#REF!')
})

test('shiftFormula clamps a range that a deletion trims rather than destroying it', () => {
  assert.equal(shiftFormula('SUM(A5:A10)', deleteRows(5)), 'SUM(A5:A9)')
  assert.equal(shiftFormula('SUM(A3:A6)', deleteRows(5, 2)), 'SUM(A3:A4)')
})

test('shiftFormula destroys a range that a deletion covers entirely', () => {
  assert.equal(shiftFormula('SUM(A5:A6)', deleteRows(5, 2)), 'SUM(#REF!)')
  assert.equal(shiftFormula('SUM(5:5)', deleteRows(5)), 'SUM(#REF!)')
})

test('shiftFormula leaves a range or whole range qualified with another sheet alone', () => {
  assert.equal(shiftFormula('Sheet3!A5:A10', insertRows(3, false)), 'Sheet3!A5:A10')
  assert.equal(shiftFormula('Sheet3!5:5', insertRows(3, false)), 'Sheet3!5:5')
})

test('shiftFormula leaves an external workbook reference alone', () => {
  assert.equal(shiftFormula('[1]Sheet1!A5', insertRows(3, false)), '[1]Sheet1!A5')
  assert.equal(shiftFormula('[1]Sheet1!A5+A5', insertRows(3)), '[1]Sheet1!A5+A6')
})

test('shiftFormula copies through fragments that are not references', () => {
  assert.equal(shiftFormula('A5:', insertRows(3)), 'A5:')
  assert.equal(shiftFormula('5:', insertRows(3)), '5:')
  assert.equal(shiftFormula("'unclosed", insertRows(3)), "'unclosed")
  assert.equal(shiftFormula('Sheet1!+A5', insertRows(3)), 'Sheet1!+A6')
})
