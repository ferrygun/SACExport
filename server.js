"use strict";
const port = process.env.PORT || 3000;
const server = require("http").createServer();
const cors = require('cors');
const multer = require('multer')
const upload = multer()
const htmlparser2 = require("htmlparser2");
const cheerio = require('cheerio')
const sap = "data-sap-widget-id";
const express = require("express");
const app = express();

//Compression
app.use(require("compression")({
    threshold: "1b"
}));


let http = require("http");
//https://stackoverflow.com/questions/58864234/fetch-o-access-control-allow-origin-header-is-present-on-the-requested-resou
app.use(cors());
app.use(function(req, res, next) {
    res.header("Access-Control-Allow-Origin", "*"); // update to match the domain you will make the request from    
    res.header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    next();
});

app.get("/export", (req, res) => {
    res.sendStatus(200);
});


app.post('/export', upload.none(), function(req, res, next) {
    const jsonsetting = JSON.parse(req.body.export_settings_json);
    let metadata = JSON.parse(jsonsetting.metadata);
    let rawHtml = req.body.export_content;
    const $ = cheerio.load(rawHtml);
    let htmlfinal = "";
    let Keys = [];

    let obj = metadata.components;
    let keys = Object.keys(obj);

    let objentries = Object.entries(obj);

    if(jsonsetting.ppt_exclude !== "") {
        var widgetexclude = JSON.parse(jsonsetting.ppt_exclude);
    
        let component, isExcluded, type;
        for (let i=0; i<widgetexclude.length; i++) {
            component =  widgetexclude[i].component;
            isExcluded = widgetexclude[i].isExcluded;
            type = widgetexclude[i].type;

            if(!isExcluded && type !== "sdk_com_fd_djaja_sap_sac_export__0") {
               Keys.push(keys[i]);
            }
        }   
    } else {
        let type;
        for(let k = 0; k<keys.length; k++) {
            type = objentries[k][1].type;

            if(type !== "sdk_com_fd_djaja_sap_sac_export__0") {
                Keys.push(keys[k]);
            }
        }
    }

    const parser = new htmlparser2.Parser({
        onopentag(name, attribs) {
            for(let k = 0; k<Keys.length; k++) {
                if (attribs[sap] === Keys[k]) {
                    let html = cheerio.html($('#' + attribs.id));
                    htmlfinal += html;  
                
                }
            }
        }
    }, {
        decodeEntities: true
    });
    parser.write(
        rawHtml
    );
    parser.end(
        res.status(200).send(htmlfinal)
    );
})

//Start the Server 
server.on("request", app);
server.listen(port, function() {
    console.info(`HTTP Server: ${server.address().port}`);
});