'use strict';

var fs = require('fs')
var JSZip = require('jszip')
var xml2js = require('xml2js')

function parseXmlAsync(text) {
    return new Promise((resolve, reject) => {
        new xml2js.Parser().parseString(text, (err, data) => {
            if (err) {
                reject(err)
            } else {
                resolve(data)
            }
        })
    })
}

function Workbook() {
    this.__xml = ''
    this.worksheets = []
}

Workbook.prototype.load = function(data) {
    var self = this
    return new JSZip().loadAsync(data)
    .then((zip) => {
        var qs = []
        qs.push(zip.file('xl/workbook.xml').async('string').then((text) => {
            parseXmlAsync(text).then((data) => this.__xml = data)
        }))
        zip.folder('xl/worksheets/').forEach((path, entry) => {
            if (path.indexOf('/') < 0) {
                qs.push(entry.async('string').then((text) => {
                    parseXmlAsync(text).then((data) => this.worksheets.push([path, data]))
                }))
            }
        })
        return Promise.all(qs)
    }).then(() => {
        return self
    })
}

Workbook.prototype.loadFile = function(path) {
    return new Promise((resolve, reject) => {
        fs.readFile(path, (err, data) => {
            if (err) {
                reject(err)
            } else {
                resolve(this.load(data))
            }
        })
    })
}

module.exports = {
    'Workbook': Workbook
}
