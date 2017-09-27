var iter = 0;
var start = Date.now();
var valsArray = [];
(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            console.log("initialized");
            $('#load-chained-thens-flow').click(loadChainedThensOnCtx);
            $('#load-chained-runs-flow').click(loadChainedThensOnRun);
            $('#load-static-flow').click(loadDataFlow);
            $('#memory').click(memory);
        });
    };
    
    var t1, t0, partTime0, partTime1;

    function memory() {
        for (var i = 0; i < 11; i++) {
            valsArray.push(loadCSVFiles(i));
        }
    }

    function loadCSVFiles(i) {
        var filename = "onco/" + (i + 1) + ".csv";
        var vals;
        $.ajax({
            async: false,
            type: "GET",
            url: filename,
            dataType: "text",
            success: function(data) {vals = processData(data);}
         });
         return vals;
    }

    function processData(allText) {
        var allTextLines = allText.split(/\r\n|\n/);
        var headers = allTextLines[0].split('@');
        var lines = [];
    
        for (var i = 0; i < allTextLines.length; i++) {
            var data = allTextLines[i].split('@');
            var tarr = [];
            for (var j=0; j<headers.length; j++) {
                tarr.push(data[j]);
            }
            lines.push(tarr);
        }
        iter = allTextLines.length;
        return lines;
    }

    function loadFlow(i) {
        return Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var rangeString = "A" + (1000 * i + 1).toString() + ":FD" + ((i + 1) * 1000).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = values;
           
            return ctx.sync();
        });
    }

    function loadChainedThensOnRun() {
        var i = 0;
        var start = Date.now();
        return Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync(ctx);
        }).then(function(ctx) {
            console.log(i + " ended sync at " + (Date.now() - start));
            start = Date.now();
            ++i;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync(ctx);
        }).then(function(ctx) {
            console.log(i + " ended sync at " + (Date.now() - start));
            start = Date.now();
            ++i;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync(ctx);
        }).then(function(ctx) {
            console.log(i + " ended sync at " + (Date.now() - start));
            start = Date.now();
            ++i;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync(ctx);
        }).then(function(ctx) {
            console.log(i + " ended sync at " + (Date.now() - start));
            start = Date.now();
            ++i;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync(ctx);
        }).then(function(ctx) {
            console.log(i + " ended sync at " + (Date.now() - start));
            start = Date.now();
            ++i;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync(ctx);
        }).then(function(ctx) {
            console.log(i + " ended sync at " + (Date.now() - start));
            start = Date.now();
            ++i;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync(ctx);
        }).then(function(ctx) {
            console.log(i + " ended sync at " + (Date.now() - start));
            start = Date.now();
            ++i;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync(ctx);
        }).then(function(ctx) {
            console.log(i + " ended sync at " + (Date.now() - start));
            start = Date.now();
            ++i;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync(ctx);
        }).then(function(ctx) {
            console.log(i + " ended sync at " + (Date.now() - start));
            start = Date.now();
            ++i;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync(ctx);
        }).then(function(ctx) {
            console.log(i + " ended sync at " + (Date.now() - start));
            start = Date.now();
            ++i;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync(ctx);
        })
    }

    function loadChainedThensOnCtx() {
        var i = 0;
        var start = Date.now();
        return Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            iter = valsArray[i].length;
            var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            
            range.values = valsArray[i];
            console.log(i + " started sync at " + Date.now());
            return ctx.sync().then(function() {
                console.log(i + " ended sync at " + (Date.now() - start));
                start = Date.now();
                ++i;
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                iter = valsArray[i].length;
                var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
                console.log(rangeString);
    
                var range = sheet.getRange(rangeString);
                
                range.values = valsArray[i];
                console.log(i + " started sync at " + Date.now());
            }).then(ctx.sync)
            .then(function() {
                console.log(i + " ended sync at " + (Date.now() - start));
                start = Date.now();
                ++i;
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                iter = valsArray[i].length;
                var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
                console.log(rangeString);
    
                var range = sheet.getRange(rangeString);
                
                range.values = valsArray[i];
            }).then(ctx.sync).then(function() {
                console.log(i + " ended sync at " + (Date.now() - start));
                start = Date.now();
                ++i;
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                iter = valsArray[i].length;
                var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
                console.log(rangeString);
    
                var range = sheet.getRange(rangeString);
                
                range.values = valsArray[i];
            }).then(ctx.sync).then(function() {
                console.log(i + " ended sync at " + (Date.now() - start));
                start = Date.now();
                ++i;
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                iter = valsArray[i].length;
                var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
                console.log(rangeString);
    
                var range = sheet.getRange(rangeString);
                
                range.values = valsArray[i];
            }).then(ctx.sync).then(function() {
                console.log(i + " ended sync at " + (Date.now() - start));
                start = Date.now();
                ++i;
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                iter = valsArray[i].length;
                var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
                console.log(rangeString);
    
                var range = sheet.getRange(rangeString);
                
                range.values = valsArray[i];
            }).then(ctx.sync).then(function() {
                console.log(i + " ended sync at " + (Date.now() - start));
                start = Date.now();
                ++i;
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                iter = valsArray[i].length;
                var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
                console.log(rangeString);
    
                var range = sheet.getRange(rangeString);
                
                range.values = valsArray[i];
            }).then(ctx.sync).then(function() {
                console.log(i + " ended sync at " + (Date.now() - start));
                start = Date.now();
                ++i;
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                iter = valsArray[i].length;
                var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
                console.log(rangeString);
    
                var range = sheet.getRange(rangeString);
                
                range.values = valsArray[i];
            }).then(ctx.sync).then(function() {
                console.log(i + " ended sync at " + (Date.now() - start));
                start = Date.now();
                ++i;
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                iter = valsArray[i].length;
                var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
                console.log(rangeString);
    
                var range = sheet.getRange(rangeString);
                
                range.values = valsArray[i];
            }).then(ctx.sync).then(function() {
                console.log(i + " ended sync at " + (Date.now() - start));
                start = Date.now();
                ++i;
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                iter = valsArray[i].length;
                var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
                console.log(rangeString);
    
                var range = sheet.getRange(rangeString);
                
                range.values = valsArray[i];
            }).then(ctx.sync).then(function() {
                console.log(i + " ended sync at " + (Date.now() - start));
                start = Date.now();
                ++i;
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                iter = valsArray[i].length;
                var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
                console.log(rangeString);
    
                var range = sheet.getRange(rangeString);
                
                range.values = valsArray[i];
            }).then(ctx.sync).then(function() {
                console.log(i + " ended sync at " + (Date.now() - start));
                start = Date.now();
                ++i;
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                iter = valsArray[i].length;
                var rangeString = "A" + (iter * i + 1).toString() + ":GZ" + ((i + 1) * iter).toString();
                console.log(rangeString);
    
                var range = sheet.getRange(rangeString);
                
                range.values = valsArray[i];
            }).then(ctx.sync).then(console.log(i + " ended sync at " + (Date.now() - start)));
        }).catch(function() {console.log("catch block")});
    }
    
    function loadDataFlow() {
        start = Date.now();
        var promise = new OfficeExtension.Promise(function(resolve, reject) { resolve (null); });

        for (var i = 0; i < 11; i++) {
            (function(i) {
                promise = promise.then(function() {
                    partTime1 = performance.now();
                    console.log((i-1) + ". " + (partTime1 - partTime0) + "ms");                            
                    console.log((i-1) + " ended run at " + Date.now());
                    var nextPromise = loadFlow(i);
                    console.log(i + " started run at " + Date.now());
                    partTime0 = performance.now();                            
                    return nextPromise;
                })
            })(i);
        }
    }
})();