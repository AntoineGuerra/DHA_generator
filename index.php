<!DOCTYPE html>
<html>
<head>
    <title>
        DHA GENERATOR</title>
    <meta charset="utf-8"/>
    <link type="text/css" rel="stylesheet" href="public/css/bootstrap_4.css">
    <meta name="google-signin-scope" content="profile email">
    <meta name="google-signin-client_id" content="889910005804-93lmvk75fa0un1dpd7ju8usqqp1orf2g.apps.googleusercontent.com">
    <script src="https://apis.google.com/js/platform.js" async defer></script>
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous">
    <link rel="stylesheet" href="public/css/custom.css">
</head>
<body>


<!--Add buttons to initiate auth sequence and sign out-->
<div class="container">
    <h1 class="text-center">DHA GENERATOR</h1>
    <div class="row">
        <div class="g-signin2" data-onsuccess="onSignIn" data-theme="dark" id="connectGoogle"></div>
        <button class="btn btn-outline-danger rounded d-none mr-3" id="signOut">Sign Out</button>
        <button id="settings"><i class="fas fa-cog"></i></button>

        <script>

        </script>
    </div>
    <div id="settings-content" class="modal">
        <!-- Modal content -->
        <div class="modal-content">
            <span class="close">&times;</span>
            <button class="btn" id="autoDl">Disable auto Download</button><br>
            <span class="text-muted">Famille d'activité par défault</span>
            <select class="btn btn-light" id="defaultFamily">
                <option id="defaultFamilyMEP" value="MEP">Mise En Production</option>
                <option id="defaultFamilyGraphisme" value="Graphisme">Graphisme</option>
                <option id="defaultFamilyConception" value="Conception">Conception</option>
                <option id="defaultFamilyPDP" value="PDP">Pilotage de projet</option>
                <option id="defaultFamilyDirection" value="Direction">Direction conseil, technique et éditoriale</option>
                <option id="defaultFamilyContenus" value="Contenus">Contenus et CM</option>
                <option id="defaultFamilyMaintenance" value="Maintenance">Maintenance et interventions post-projet</option>
            </select>
        </div>
    </div>
    <div class="row" >

    </div>
    <div class="row d-none m-3" id="UserDiv">
        <img id="user-image" class="rounded-circle">
        <h3 id="user-name" style="line-height: 50px" class="pl-2"></h3>
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
<script type="text/javascript" src="public/js/modal-settings.js"></script>
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
</html>