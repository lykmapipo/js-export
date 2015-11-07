js-export
==========

[![Build Status](https://travis-ci.org/lykmapipo/js-export.svg?branch=master)](https://travis-ci.org/lykmapipo/js-export)

Utilities to export js data into different acceptable format for [nodejs](https://github.com/nodejs)

## Installation
```sh
$ npm install --save js-export
```

## Usage
```js
var JSExport = require('js-export');
var jsexport = new JSExport(data, options);

//export data as excel file
jsexport.writeExcel(`<file>`, done);

//download data as excel through http requests
app.get('/exports', function(request, response){
   jsexport.downloadExcel(response, options); 
});
```

## Supported Formats
- [ ] excel
- [ ] csv
- [ ] text
- [ ] json