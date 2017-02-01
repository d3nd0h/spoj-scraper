'use strict';

var _ = require('lodash');
var Bluebird = require('bluebird');

var scraper = require('./modules/scraper');
var xlsxHelper = require('./helpers/xlsx');

function iterateAll(users, problems) {
  return Bluebird.map(users, function (user) {
    return Bluebird.props({
      user: user,
      result: Bluebird.map(problems, function (problemSpec) {
        return Bluebird.props({
          date: problemSpec.date,
          problems: Bluebird.map(problemSpec.problems, function (problemCode) {
            return scraper.getFirstAC(user, problemCode);
          })
        });
      })
    });
  });
}

exports.runAll = function () {
  var filename = 'files/data.xlsx';
  var data = xlsxHelper.read(filename);

  iterateAll(data.users, data.problems)
    .then(function (res) {
      res.forEach(function (a) { console.log(a.result) });
    })
    .catch(console.log)
};

exports.runAll();
