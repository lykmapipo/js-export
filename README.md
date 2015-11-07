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

## Plugin Export Engine
`js-export` support additional export engine through plugins. To add new export engine do as below and implement `write` and `download` methods

```js
var JSExport = require('js-export');
var jsexport = new JSExport(data, options);

//buffer engine
var bufferEngine = {
    write: function(path, done){
        //codes
        ...
    },

    download:function(response, options){
        //codes
        ...
    }
}

//use export engine
jsexport.use('buffer', bufferEngine);

//then use buffer export engine
jsexport.writeBuffer(path, done);

```

## Supported Formats
- [ ] excel
- [ ] csv
- [ ] text
- [ ] json