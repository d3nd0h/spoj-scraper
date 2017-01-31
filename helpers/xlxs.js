'use strict';
var XLSX = require('xlsx');
var _ = require('lodash');

function parseProblems(workbook) {
  var firstSheetName = workbook.SheetNames[0];
  var worksheet = workbook.Sheets[firstSheetName];

  var cols = _.split('BCDEFGHIJKLMNOPQRSTUVWXYZ', '');
  var columns = _.reduce(cols, function (acc, val) {
    var cell = val + '1';
    if (!_.isEmpty(worksheet[cell])) {
      return acc.concat(val);
    }

    return acc;
  }, []);

  var row = 2;
  var allProblems = [];
  while (!_.isEmpty(worksheet['A'.concat(row)])) {
    var problemsThatDay = _.reduce(columns, function (acc, col) {
      var cell = worksheet[col.concat(row)];
      if (!_.isEmpty(cell)) {
        return acc.concat(cell.v);
      }

      return acc;
    }, []);

    allProblems.push({
      date: worksheet['A'.concat(row)].v,
      problems: problemsThatDay
    });

    row++;
  }

  return allProblems;
}

exports.read = function (filename) {
  var workbook = XLSX.readFile(filename);

};

exports.read('files/test.xlsx');
