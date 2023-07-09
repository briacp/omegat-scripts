/**************************************************************************
 OmegaT - Computer Assisted Translation (CAT) tool
          with fuzzy matching, translation memory, keyword search,
          glossaries, and translation leveraging into updated projects.

 Copyright (C) 2023 Briac Pilpré
               Home page: https://www.omegat.org/
               Support center: https://omegat.org/support

 This file is part of OmegaT.

 OmegaT is free software: you can redistribute it and/or modify
 it under the terms of the GNU General Public License as published by
 the Free Software Foundation, either version 3 of the License, or
 (at your option) any later version.

 OmegaT is distributed in the hope that it will be useful,
 but WITHOUT ANY WARRANTY; without even the implied warranty of
 MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 GNU General Public License for more details.

 You should have received a copy of the GNU General Public License
 along with this program.  If not, see <https://www.gnu.org/licenses/>.
 **************************************************************************/
// jshint esversion:6

// Generates OmegaT changelog
// usage: node omegat_changelog [milestone or tag]

const https = require('https');

// If the tag is '6.1.0', the milestone is '6.1'
const versionTag = process.argv[2] || 'v6.1.0_0';
const milestone = versionTag.replace(/^v(\d+)\.(\d+)\.(\d)_.*/, '$1.$2.$3').replace(/\.0$/, '');

const trackers = {
    'bugs' : [],
    'feature-requests': [],
    //'documentation': [],
};

request(getFixedTicket('bugs', milestone))
.then( res => parseResponse(res, 'bugs'))
.then( () => request(getFixedTicket('feature-requests', milestone)))
.then( res => parseResponse(res, 'feature-requests'))
.then( () => {
    console.log(`----------------------------------------------------------------------`);
    console.log(` OmegaT ${versionTag}`);
    console.log(`----------------------------------------------------------------------`);
    console.log(`  ${trackers['feature-requests'].length} enhancements`);
    console.log(`  ${trackers.bugs.length} bug fixes`);
    console.log(`----------------------------------------------------------------------`);

    console.log('');
    console.log('  Implemented requests:');
    trackers['feature-requests'].forEach(rfe => {
        console.log('');
        console.log(`  - ${rfe.summary}`);
        console.log(`    ${rfe.url}`);
    });

    console.log('\n');
    console.log('  Bug fixes:');
    trackers.bugs.forEach(rfe => {
        console.log('');
        console.log(`  - ${rfe.summary}`);
        console.log(`    ${rfe.url}`);
    });

});

function parseResponse(res, tracker) {
    const data = JSON.parse(res.body);
    data.tickets.forEach(ticket => {
        trackers[tracker].push({
          id: ticket.ticket_num,
          summary: ticket.summary,
          labels: ticket.labels,
          url: `https://sourceforge.net/p/omegat/${tracker}/${ticket.ticket_num}/`
      });
  });
}

function getFixedTicket(tracker, milestone) {
    return `https://sourceforge.net/rest/p/omegat/${tracker}/search?q=status:open-fixed+AND+_milestone:${milestone}`;
}

function request(urlOptions, data = '') {
    return new Promise((resolve, reject) => {
    // Inspired from https://gist.github.com/ktheory/df3440b01d4b9d3197180d5254d7fb65
    const req = https.request(urlOptions, res => {
      // I believe chunks can simply be joined into a string
      const chunks = [];
  
      res.on('data', chunk => chunks.push(chunk));
      res.on('error', reject);
      res.on('end', () => {
        const { statusCode, headers } = res;
        const validResponse = statusCode >= 200 && statusCode <= 299;
        const body = chunks.join('');
  
        if (validResponse) resolve({ statusCode, headers, body });
        else reject(new Error(`Request failed. status: ${statusCode}, body: ${body}`));
      });
    });
    req.on('error', reject);
    req.write(data, 'binary');
    req.end();
  });
}