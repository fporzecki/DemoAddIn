var iter = 0;
var start = Date.now();

(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            console.log("initialized");
            $('#load-data-segments').click(load);
            $('#load-data-flow').click(loadDataFlow);
        });
    };
    var t1, t0, partTime0, partTime1;

    function load() {
        console.log("click");
        document.getElementById("load-data-segments").disabled = true;
        loadDataSegments(iter++);
    }

    function loadDataSegments(i) {
        t0 = performance.now();

        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();

            partTime0 = performance.now();

            var rangeString = "A" + (1000 * i + 1).toString() + ":FD" + ((i + 1) * 1000).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            range.values = values;

            partTime1 = performance.now();
            console.log((i+1) + ". " + (partTime1 - partTime0) + "ms");
            
            return ctx.sync();
        })
        .then(function () {
            document.getElementById("load-data-segments").disabled = false;
            t1 = performance.now();
            console.log((i+1) + ". " + (t1 - t0) + "ms");
            app.showNotification("Success " + (t1 - t0) + "ms");
            console.log("Success!");
        })
        .catch(function (error) {
            document.getElementById("load-data-segments").disabled = false;
            app.showNotification("Error: " + error);
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function loadFlow(i) {
        return Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();

            var rangeString = "A" + (1000 * i + 1).toString() + ":FD" + ((i + 1) * 1000).toString();
            console.log(rangeString);

            var range = sheet.getRange(rangeString);
            range.values = values;
           
            return ctx.sync();//.then(()=>console.log("synced"));
        });
    }
    
    function loadDataFlow() {
        start = Date.now();
        var promise = new OfficeExtension.Promise(function(resolve, reject) { resolve (null); });

        for (var i = 0; i < 16; i++) {
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