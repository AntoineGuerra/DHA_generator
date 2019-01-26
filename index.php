<!DOCTYPE html>
<html>
<head>
    <title>
        Google
        Calendar
        API
        Quickstart</title>
    <meta charset="utf-8"/>
</head>
<body>
<p>
    CSV DATA TO DHA</p>

<!--Add buttons to initiate auth sequence and sign out-->
<button id="authorize_button"
        style="display: none;">
    Authorize
</button>
<button id="signout_button"
        style="display: none;">
    Sign
    Out
</button>
<label>Acronyme
    :
    <input id="acronyme"
           type="text"
           placeholder="AGU"></label>
<label>Semaine
    N°
    <input id="week"
           type="number"
           placeholder="4"></label>
<button id="generate">
    Générer
</button>
<pre id="content"
     style="white-space: pre-wrap;"></pre>
<div>
    <span>
        Toutes les parties
        <p id="full">

        </p>
    </span>
    <span>
        Partie : Avant Ventes
        <p id="AV">

        </p>
    </span>
    <span>
        Partie : Vendus
        <p id="V">

        </p>
    </span>
    <span>
        Partie : Maintenances
        <p id="M">

        </p>
    </span>
    <span>
        Partie : Internes
        <p id="I">

        </p>
    </span>
    <span>
        Impossible à traiter
        <p id="err">

        </p>
    </span>

</div>
<!--<link href="http://cdn.grapecity.com/spreadjs/hosted/css/gc.spread.sheets.excel2013white.11.0.1.css" rel="stylesheet" type="text/css" />-->
<!--<script type="text/javascript" src="http://cdn.grapecity.com/spreadjs/hosted/scripts/gc.spread.sheets.all.11.0.1.min.js"></script>-->
<!--<script type="text/javascript" src="http://cdn.grapecity.com/spreadjs/hosted/scripts/interop/gc.spread.excelio.11.0.1.min.js"></script>-->
<script async
        defer
        src="https://apis.google.com/js/api.js"
        onload="this.onload=function(){};handleClientLoad()"
        onreadystatechange="if (this.readyState === 'complete') this.onload()">
</script>

<script type = "text/javascript" src = "public/js/script.js"></script>
<?php
//    $inputFileType1 = 'Excel2007';
//    $inputFileName1 = './public/assets/DHA/DHA-Part1.xlsx';
//    $inputFileType1 = 'Excel2007';
//    $inputFileName1 = './public/assets/DHA/DHA-Part1.xlsx';
//

?>
</body>
</html>