'use strict';
var XLSX = require('xlsx');
var _ = require('lodash');
var moment = require('moment');

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

function createHeadings(data) {
  if (_.isEmpty(data)) {
    return null;
  }

  return _.reduce(data[0].results, function (acc, result) {
    return acc.concat(_.map(result.problems, 'code'));
  }, []);
}

function checkScore(firstACDate, problemDate) {
  var startDayTime = moment(problemDate, 'DD/MM/YYYY');
  var endDayTime = moment(problemDate, 'DD/MM/YYYY');
  startDayTime.hour('08');
  endDayTime.hour('17');

  var ret = {
    isRepeating: false,
    score: 0
  };

  if (_.isEmpty(firstACDate)) {
    return ret;
  } else {
    firstACDate = moment(firstACDate);
    ret.score = 1;
  }

  if (firstACDate.isAfter(startDayTime) && firstACDate.isBefore(endDayTime)) {
    ret.isRepeating = false
  } else {
    ret.isRepeating = true;
  }

  return ret;
}

function dumpDataToArray(data) {
  var headings = ['username'].concat(createHeadings(data));
  var dataInArray = _.map(data, function (userData) {
    return _.reduce(userData.results, function (acc, result) {
      return acc.concat(_.map(result.problems, function (problemResult) {
        return checkScore(problemResult.firstACDate, result.date);
      }));
    }, [userData.user]);
  });
  return [headings].concat(dataInArray);
}

function createWorkbook(array) {
  if (_.isEmpty(array)) {
    throw new Error('Array is empty.');
  }

  var repeatingSheet = {};
  var practiceSheet = {};
  var workbook = {
    SheetNames: ['Practice', 'Repeating'],
    Sheets: {
      Practice: practiceSheet,
      Repeating: repeatingSheet
    }
  };

  // Set sheet range
  var sheetRange = {
    s: {c: 0, r: 0},
    e: {c: array[0].length - 1, r: array.length - 1}
  };
  repeatingSheet['!ref'] = XLSX.utils.encode_range(sheetRange);
  practiceSheet['!ref'] = XLSX.utils.encode_range(sheetRange);

  for (var R = 0; R <= sheetRange.e.r; R++) {
    for (var C = 0; C <= sheetRange.e.c; C++) {
      var cellRef = XLSX.utils.encode_cell({c: C, r: R});
      var data = array[R][C];
      if (typeof(data) === 'string') {
        practiceSheet[cellRef] = { v: data };
        repeatingSheet[cellRef] = { v: data };
      } else {
        repeatingSheet[cellRef] = { v: data.score };
        if (!data.isRepeating) {
          practiceSheet[cellRef] = { v: data.score };
        } else {
          practiceSheet[cellRef] = { v: 0 };
        }
      }
    }
  }

  return workbook;
}

exports.read = function (filename) {
  var workbook = XLSX.readFile(filename);

  return {
    problems: parseProblems(workbook),
    users: parseUsers(workbook)
  };
};

exports.write = function (data, filename) {
  var workbook = createWorkbook(dumpDataToArray(data));
  XLSX.writeFile(workbook, filename);
};
