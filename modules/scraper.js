'use strict';
var Bluebird = require('bluebird');
var cheerio = require('cheerio');
var _ = require('lodash');
var moment = require('moment');
var request = Bluebird.promisify(require('request'));

exports.getFirstAC = function (username, problem) {
  var spojUrl = 'http://www.spoj.com/status/';
  var url = spojUrl + username.toLowerCase() + ',' + problem.toUpperCase() + '/';

  return request(url)
    .then(function (res) {
      var $ = cheerio.load(res.body);
      var firstAC = $('td[status="15"]').last();

      if (firstAC.length === 0) {
        return null;
      }

      var data = $(firstAC).parent().children().toArray();
      var date = moment.utc(_.trim($(data[1]).text()));
      date.subtract(1, 'hour');

      return date;
    })
};
