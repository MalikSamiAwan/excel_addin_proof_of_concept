//var jsLib=require("https://appsforoffice.microsoft.com/lib/1/hosted/office.js");




function showMessageFunction(){
  alert('Hello World!')
}

function showSheetNames(){

//jsLib.Excel.run(function (context) {
//    var worksheets = context.workbook.worksheets;
//    worksheets.load('name');
//    return context.sync()
//    .then(function() {
//        for (var i = 0; i < worksheets.items.length; i++)
//        {
//            console.log(worksheets.items[i].name);
//            alert('Sheet Name :'+worksheets.items[i].name);
//        }
//    });
//})

}

/*
<!--    window.addEventListener('load', function(ev) {-->
<!--      // Download main.dart.js-->
<!--      _flutter.loader.loadEntrypoint({-->
<!--        serviceWorker: {-->
<!--          serviceWorkerVersion: serviceWorkerVersion,-->
<!--        },-->
<!--        onEntrypointLoaded: function(engineInitializer) {-->
<!--          engineInitializer.initializeEngine().then(function(appRunner) {-->
<!--            appRunner.runApp();-->
<!--          });-->
<!--        }-->
<!--      });-->
<!--    });-->
*/

/*
//whole page

<!DOCTYPE html>
<html>
<head>
  <!--
    If you are serving your web app in a path other than the root, change the
    href value below to reflect the base path you are serving from.

    The path provided below has to start and end with a slash "/" in order for
    it to work correctly.

    For more details:
    * https://developer.mozilla.org/en-US/docs/Web/HTML/Element/base

    This is a placeholder for base href that will be replaced by the value of
    the `--base-href` argument provided to `flutter build`.
  -->
  <base href="$FLUTTER_BASE_HREF">

  <meta charset="UTF-8">
  <meta content="IE=Edge" http-equiv="X-UA-Compatible">
  <meta name="description" content="A new Flutter project.">

  <!-- iOS meta tags & icons -->
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="black">
  <meta name="apple-mobile-web-app-title" content="testing4">
  <link rel="apple-touch-icon" href="icons/Icon-192.png">

  <!-- Favicon -->
  <link rel="icon" type="image/png" href="favicon.png"/>

  <title>testing4</title>
  <link rel="manifest" href="manifest.json">
<!--  <script src='index.js' defer></script>-->
<!--  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>-->

  <script>
    // The value below is injected by flutter build, do not touch.
    const serviceWorkerVersion = null;
  </script>
  <!-- This script adds the flutter initialization JS code -->
  <script src="flutter.js" defer></script>

  <!-- Yandex.RTB -->
  <script>
      window.yaContextCb = window.yaContextCb || [];
    </script>
  <script src="https://yandex.ru/ads/system/context.js" async></script>

</head>
<body>

<div id="loading">
  <style>
        body {
          inset: 0;
          overflow: hidden;
          margin: 0;
          padding: 0;
          position: fixed;
        }
        #loading {
          align-items: center;
          display: flex;
          height: 100%;
          justify-content: center;
          width: 100%;
        }
        #loading img {
          animation: 1s ease-in-out 0s infinite alternate breathe;
          opacity: 0.66;
          transition: opacity 0.4s;
        }
        #loading.main_done img {
          opacity: 1;
        }
        #loading.init_done img {
          animation: 0.33s ease-in-out 0s 1 forwards zooooom;
          opacity: 0.05;
        }
        @keyframes breathe {
          from {
            transform: scale(1);
          }
          to {
            transform: scale(0.95);
          }
        }
        @keyframes zooooom {
          from {
            transform: scale(1);
          }
          to {
            transform: scale(10);
          }
        }
      </style>
  <img src="icons/icon-192.png" alt="Loading indicator..." />
</div>

  <script>

   window.addEventListener("load", function (ev) {
        var loading = document.querySelector("#loading");
        // Download main.dart.js
        _flutter.loader
          .loadEntrypoint({
            serviceWorker: {
              serviceWorkerVersion: serviceWorkerVersion,
            },
          })
          .then(function (engineInitializer) {
            loading.classList.add("main_done");
            return engineInitializer.initializeEngine();
          })
          .then(function (appRunner) {
            loading.classList.add("init_done");
            return appRunner.runApp();
          })

          .then(function () {
            window.setTimeout(function () {
              loading.remove();
            }, 200);

            console.log("intializing office");
            const officeEl = document.getElementById("office");
            if (officeEl != null) return;

            const scriptTag = document.createElement("script");
            scriptTag.src =
              "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
            scriptTag.id = "office";
            scriptTag.addEventListener("load", () => {
              console.log("office loaded");
              class OfficeHelpers {
                runExcel = Excel.run;
                officeOnReady = Office.onReady;
              }
              window["getOfficeHelpers"] = () => new OfficeHelpers();
              console.log("helpers injected");
            });
            document.getElementsByTagName("head")[0].appendChild(scriptTag);
          });
      });



  </script>
</body>

</html>

*/