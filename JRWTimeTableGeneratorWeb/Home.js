(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;
    const INPUT = document.getElementById("url"),
        copyInfo = document.getElementById("copyInfo"),
        worksheetW = document.getElementById("worksheetW"),
        worksheetH = document.getElementById("worksheetH"),
        DESTINATION_URL = "https://script.google.com/macros/s/AKfycbwlrB8oXEt4hLlEjsQkVQ-tVjHdbVHRHzHeLryZnOq2zMzn_To6ymBqTW0QWfibbNk/exec",
        TRAIN_TYPE_URL = "https://script.google.com/macros/s/AKfycbz2YGPFbOk5RqOju0AONsqDU6AnRzA-X-hORI4qJBM7sjZErZHsvspHSIYPieW3SEkyqQ/exec";
    let destination, destinationColor, destinationArray;//Google SpreadSheet
    let trainType, trainTypeColor, trainTypeArray;//Google SpreadSheet
    let trainAllData = []; //hourで参照

    //列車情報読み込み
    const arrFunc = [];
    const fetchD = (resolve) => {
        fetch(DESTINATION_URL).then((response) => {
            return response.json();
        }).then((obj) => {
            return JSON.parse(JSON.stringify(obj, null, " "));
        }).then((jsonObj) => {
            destinationArray = jsonObj.allData;
            resolve();
        }).catch((error) => {
            console.log("スプレッドシートDが読み込めません");
            console.log(error);
        })
    };
    arrFunc.push(fetchD);
    const fetchT = (resolve) => {
        fetch(TRAIN_TYPE_URL).then((response) => {
            return response.json();
        }).then((obj) => {
            return JSON.parse(JSON.stringify(obj, null, " "));
        }).then((jsonObj) => {
            trainTypeArray = jsonObj.allData;
            resolve();
        }).catch((error) => {
            console.log("スプレッドシートTが読み込めません");
            console.log(error);
        })
    };
    arrFunc.push(fetchT);
    const arrPromise = arrFunc.map((func) => new Promise(func));
    Promise.all(arrPromise).then(() => {
        const loading = document.getElementById('loading');
        loading.classList.add('loaded');
    });

    
    function worksheet(day) {
        switch (day) {
            case "weekdays": return worksheetW;
            case "holidays": return worksheetH;
            case "reset":
                worksheetW.innerHTML = "更新中...";
                worksheetH.innerHTML = "更新中...";
            case "error":
                worksheetW.innerHTML = "データ取得後に出力・更新されます";
                worksheetH.innerHTML = "データ取得後に出力・更新されます";
                break;
        }
    }
    const dayList = ["日", "月", "火", "水", "木", "金", "土"];

    // 新しいページが読み込まれるたびに初期化関数を実行する必要があります。
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // 通知メカニズムを初期化して非表示にします
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // Excel 2016 を使用していない場合は、フォールバック ロジックを使用してください。
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                /*$("#template-description").text("このサンプルでは、スプレッドシートで選ばれたセルの値が表示されます。");
                $('#button-text').text("表示!");
                $('#button-desc').text("選択範囲が表示されます");

                $('#highlight-button').click(displaySelectedCells);*/
                //return;
            }
                
            loadSampleData();

            // 強調表示ボタンのクリック イベント ハンドラーを追加します。
            $('#highlight-button').click(hightlightHighestValue);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Excel オブジェクト モデルに対してバッチ操作を実行します
        Excel.run(function (ctx) {
            // 作業中のシートに対するプロキシ オブジェクトを作成します
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // ワークシートにサンプル データを書き込むコマンドをキューに入れます
            sheet.getRange("B3:D5").values = values;

            // キューに入れるコマンドを実行し、タスクの完了を示すために Promise を返します
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('選択されたテキスト:', '"' + result.value + '"');
                } else {
                    showNotification('エラー', result.error.message);
                }
            });
    }

    // エラーを処理するためのヘルパー関数
    function errorHandler(error) {
        // Excel.run の実行から浮かび上がってくるすべての累積エラーをキャッチする必要があります
        showNotification("エラー", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // 通知を表示するヘルパー関数
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
