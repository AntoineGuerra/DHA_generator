<!DOCTYPE html>
<html>
<body>


<!--Add buttons to initiate auth sequence and sign out-->
<div class="container">
    <h1 class="text-center">DHA GENERATOR</h1>
    <div class="row">
        <button id="authorize_button"
                style="display: none;" class="btn-outline-success btn">
            Authorize
        </button>
        <button id="signout_button"
                style="display: none;" class="btn btn-outline-danger">
            Sign
            Out
        </button>
    </div>
    <div class="row">
        <form>
            <label class="col-12">Acronyme :</label>
            <input id="acronyme" type="text" placeholder="AGU" class="btn btn-light col-6">

            <label class="col-6">Semaine N° :</label>
            <input id="week" type="number" placeholder="4" class="btn btn-light col-6">

            <button id="generate" class="btn btn-outline-primary col-12">Générer</button>
        </form>
    </div>
    <div class="row d-none m-2" id="div_DHA">
        <div class="col-12 m-2">
            <span class="col-6 col-md-4 ">DHA :</span>
        </div>
        <div class="col-12">

            <a class="col-6" id="DHA">
                <button id="full" class="btn-outline-primary"></button>
            </a>
        </div>
    </div>
    <div class="row d-none" id="div_err">
        <hr>
        <span class="col-12">Votre Projet doit être sous cette forme :</span>
        <span class="text-success col-12">V (categorie) - Client - Projet - MEP (famille : optionnel Pour AV et INT)</span>
        <hr>
        <span class="col-12">
            Impossible à traiter
        <div class="col-12" id="err">

        </div>
    </span>
    </div>
</div>

<!--<pre id="content"-->
<!--     style="white-space: pre-wrap;"></pre>-->
<div>


</div>
<!--<link href="http://cdn.grapecity.com/spreadjs/hosted/css/gc.spread.sheets.excel2013white.11.0.1.css" rel="stylesheet" type="text/css" />-->
<!--<script type="text/javascript" src="http://cdn.grapecity.com/spreadjs/hosted/scripts/gc.spread.sheets.all.11.0.1.min.js"></script>-->
<!--<script type="text/javascript" src="http://cdn.grapecity.com/spreadjs/hosted/scripts/interop/gc.spread.excelio.11.0.1.min.js"></script>-->
<script type="text/javascript" src="public/js/xmlContainer.js"></script>
<script type="text/javascript" src="public/js/script.js"></script>
<script async defer src="https://apis.google.com/js/api.js" onload="this.onload=function(){};handleClientLoad()" onreadystatechange="if (this.readyState === 'complete') this.onload()"></script>

<?php
//    $inputFileType1 = 'Excel2007';
//    $inputFileName1 = './public/assets/DHA/DHA-Part1.xlsx';
//    $inputFileType1 = 'Excel2007';
//    $inputFileName1 = './public/assets/DHA/DHA-Part1.xlsx';
//

?>
</body>
<head>
    <title>
        DHA GENERATOR</title>
    <meta charset="utf-8"/>
    <link type="text/css" rel="stylesheet" href="public/css/bootstrap_4.css">
</head>
</html>