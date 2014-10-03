"use strict";

require("console-stamp")(console, {pattern: 'ddd dd HH:MM:ss'});

var AWS           = require('aws-sdk');
var child_process = require('child_process');
var co            = require('co');
var fs            = require('fs');
var nodefn        = require('when/node/function');
var path          = require('path');
var temp          = require('temp');
var when          = require('when');
var config        = require(process.env.CONFIG_FILE);

var BUCKET   = 'cbd.document-agent';
var WORD2PDF = path.join(process.cwd(), 'word-to-pdf', 'WordToPdf.exe');

var S3 = new AWS.S3({
    accessKeyId: config.awsAccessKeys.global.accessKeyId,
    secretAccessKey: config.awsAccessKeys.global.secretAccessKey,
    region: 'us-east-1',
    apiVersion: '2006-03-01',
});

S3.listObjects  = nodefn.lift(S3.listObjects .bind(S3));
S3.getObject    = nodefn.lift(S3.getObject   .bind(S3));
S3.putObject    = nodefn.lift(S3.putObject   .bind(S3));
S3.deleteObject = nodefn.lift(S3.deleteObject.bind(S3));

return when(co(function* () {

    while(true) {
        let hasDocument = yield poll();
        yield when(0).delay(hasDocument ? 5*1000 : 60*1000);
    }

})).catch(function (error) {

    console.log(`[ERROR] ${error}`);
});

//============================================================
//
//
//============================================================
function* poll() {

    console.log('polling...');

    let documents = yield S3.listObjects({ Bucket: BUCKET, Prefix: 'source/', MaxKeys: 2 });

    if(documents.Contents.length>1) {
        yield processKey(documents.Contents[1].Key);
        return true;
    }

    return false;
}

//============================================================
//
//
//============================================================
function* processKey(key) {

    console.log(`[info] Processing ${key}...`);

    let keyParts   = path.parse(key);
    let tempPath   = temp.path();
    let sourceFile = `${tempPath}${keyParts.ext}`;
    let targetFile = `${tempPath}.pdf`;
    let targetKey  = `target/${keyParts.name}.pdf`;

    // GET SOURCE FILE

    let source = yield S3.getObject({ Bucket: BUCKET, Key: key });
    fs.writeFileSync(sourceFile, source.Body);

    let md5 = source.ETag.replace(/"/g, '');

    // GENERATE PDF
    console.log(`${sourceFile} -> ${targetFile}`);

    let result = child_process.spawnSync(WORD2PDF, [sourceFile], { timeout : 30*1000 });

    console.log(result.stdout.toString('utf-8'));

    yield S3.putObject({ Bucket: BUCKET, Key: targetKey+'.txt', Body: result.stdout });


    // UPLOAD TARGET

    if(result.status===0) {
        let targetData = fs.readFileSync(targetFile);
        yield S3.putObject({ Bucket: BUCKET, Key: targetKey, Body: targetData, Metadata: { md5: md5 } });
    }

    // DELETE SOURCE

    yield S3.deleteObject({ Bucket: BUCKET, Key: key });

    console.log(`[info] Processing ${key}...`, result.status===0 ? 'DONE' : 'ERROR');
}
