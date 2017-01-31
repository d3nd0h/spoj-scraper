'use strict';
var Bluebird = require('bluebird');
var cheerio = require('cheerio');
var _ = require('lodash');
var moment = require('moment');
var request = Bluebird.promisify(require('request'));

var spojUrl = 'http://www.spoj.com/status/';
var username = 'd3nd0h';
var problemCode = 'ILKQUERY';
var url = spojUrl + username + ',' + problemCode + '/';

request(url)
  .then(function (res) {
    var $ = cheerio.load(res.body);
    var firstAC = $('td[status="15"]').last();
    var data = $(firstAC).parent().children().toArray();
    var date = moment.utc(_.trim($(data[1]).text()));
    date.subtract(1, 'hour');
  })
  .catch(console.log);
