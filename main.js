"use strict";

require("console-stamp")(console, {pattern: 'ddd dd HH:MM:ss'});

var AWS           = require('aws-sdk');
var child_process = require('child_process');
var fs            = require('fs');
var nodefn        = require('when/node/function');
var path          = require('path');
var temp          = require('temp');
var config        = require(process.env.CONFIG_FILE);

var BUCKET   = 'cbd.document-agent';
var WORD2PDF = path.join(process.cwd(), 'word-to-pdf', 'WordToPdf.exe');
let lastError = null;

var S3 = new AWS.S3({
    accessKeyId: config.awsAccessKeys.global.accessKeyId,
    secretAccessKey: config.awsAccessKeys.global.secretAccessKey,
    region: 'us-east-1',
    apiVersion: '2006-03-01'
});

S3.listObjects  = nodefn.lift(S3.listObjects .bind(S3));
S3.getObject    = nodefn.lift(S3.getObject   .bind(S3));
S3.putObject    = nodefn.lift(S3.putObject   .bind(S3));
S3.deleteObject = nodefn.lift(S3.deleteObject.bind(S3));

//============================================================
//
//
//============================================================
async function mainLoop() {

    let nextLoop = 60*1000;

    try {

        if(await poll())
            nextLoop = 5*1000;

        lastError = null;

    } catch (error) {

        nextLoop = 15*1000;

        lastError = error;

        if(error.stack) {
            lastError = {
                message: error.message,
                stack : error.stack
            }
        }

        console.error(`[ERROR] ${error}\n${error.stack}`);
    }

    setTimeout(mainLoop, nextLoop);
}


//============================================================
//
//
//============================================================
async function poll() {

    console.log('polling...');

    let documents = await S3.listObjects({ Bucket: BUCKET, Prefix: 'source/', MaxKeys: 2 });

    if(documents.Contents.length>1) {
        await processKey(documents.Contents[1].Key);
        return true;
    }

    return false;
}

//============================================================
//
//
//============================================================
async function processKey(key) {

    console.log(`[info] Processing ${key}...`);

    let keyParts   = path.parse(key);
    let tempPath   = temp.path();
    let sourceFile = `${tempPath}${keyParts.ext}`;
    let targetFile = `${tempPath}.pdf`;
    let targetKey  = `target/${keyParts.name}.pdf`;

    // GET SOURCE FILE

    let source = await S3.getObject({ Bucket: BUCKET, Key: key });
    fs.writeFileSync(sourceFile, source.Body);

    let md5 = source.ETag.replace(/"/g, '');

    // GENERATE PDF
    console.log(`${sourceFile} -> ${targetFile}`);

    let result = child_process.spawnSync(WORD2PDF, [sourceFile], { timeout : 5*60*1000 });

    console.log((result.stdout||'{ "error": "NO-STDOUT" }').toString('utf-8'));

    await S3.putObject({ Bucket: BUCKET, Key: targetKey+'.txt', Body: result.stdout });


    // UPLOAD TARGET

    if(result.status===0) {
        let targetData = fs.readFileSync(targetFile);
        await S3.putObject({ Bucket: BUCKET, Key: targetKey, Body: targetData, Metadata: { md5: md5 } });
    }

    // DELETE SOURCE

    await S3.deleteObject({ Bucket: BUCKET, Key: key });

    if(result.status!==0) {

        let error = null;

        try {
            error = JSON.parse(result.stdout);
        } catch(e) {
            error = result.stdout || { "error": "NO-STDOUT" } || e;
        }

        throw error
    }

    console.log(`[info] Processing ${key}...`, result.status===0 ? 'DONE' : 'ERROR');
}

//============================================================
//============================================================
//============================================================
//============================================================
mainLoop();

var app = require('express')();

app.get('/', function(req, res) {
    res.status(200).send(lastError ? { date: new Date(), error: lastError } : `OK\n${new Date().toISOString()}`);
});

app.listen(process.env.PORT || 8080, '0.0.0.0', function(){
    console.info('info: Listening on %j', this.address());
});
