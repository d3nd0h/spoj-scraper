'use strict';
var XLSX = require('xlsx');
var _ = require('lodash');

function parseProblems(workbook) {
  var sheetName = workbook.SheetNames[0];
  var worksheet = workbook.Sheets[sheetName];

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

function parseUsers(workbook) {
  var sheetName = workbook.SheetNames[1];
  var worksheet = workbook.Sheets[sheetName];

  // Should be put in the first column without any headings
  var col = 'A';
  var row = 1;
  var users = [];
  while (!_.isEmpty(worksheet[col.concat(row)])) {
    var cell = col.concat(row);
    users.push(worksheet[cell].v);
    row++;
  }

  return users;
}

exports.read = function (filename) {
  var workbook = XLSX.readFile(filename);

  return {
    problems: parseProblems(workbook),
    users: parseUsers(workbook)
  };
};
