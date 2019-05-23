







    autoDl = false;
    defaultFamily = 'MEP';
    const MODE = {
        DHA: 1,
        AGENDA: 2,
    };
    mode = MODE.DHA;
    // // Array of API discovery doc URLs for APIs used by the quickstart
    var DISCOVERY_DOCS = ["https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest", "https://sheets.googleapis.com/$discovery/rest?version=v4"];
    //
    // // Authorization scopes required by the API; multiple scopes can be
    // // included, separated by spaces.
    var SCOPES = "https://www.googleapis.com/auth/calendar https://www.googleapis.com/auth/spreadsheets.readonly profile";
    //
    // // Client ID and API key from the Developer Console
    var CLIENT_ID = '889910005804-tq4btqe01v6sapk7fk65telp3kcdle3p.apps.googleusercontent.com';
    var API_KEY = 'AIzaSyD42yhsKM9BM71lbkWy1LeX7m68oKxKCbw';

    document.getElementById('change_mode').addEventListener('click', function () {
        emptyTables();
        let generateBtn = document.getElementById('generate');
        switch (mode) {
            case MODE.DHA:
                mode = MODE.AGENDA;
                generateBtn.innerText = 'Générer ' + ' mon agenda';
                break;
            case MODE.AGENDA:
                mode = MODE.DHA;
                generateBtn.innerText = 'Générer ' + ' ma DHA';
                break;
        }
    });
    document.getElementById('acronyme').addEventListener('change', function () {
        document.cookie = 'acronyme=' + this.value.toUpperCase() + '; expires=Fri, 31 Dec 2030 23:59:59 GMT';
    });
    function emptyTables() {
        for(key in MODE) {
            console.log(key + ': ' + MODE[key]);
            let div = document.getElementById('div_' + key);
            if (div !== undefined && div !== null) {
                document.getElementById('div_' + key).classList.add('d-none');
            }
            document.getElementById('table-' + key).classList.add('d-none');
            document.getElementById('table-' + key + '-content').innerHTML = '';
            document.querySelector('.error-' + key).classList.add('d-none');
        }
        document.getElementById('macro_file').innerHTML = '';
        // document.getElementById('table-error-AGENDA-content').innerHTML = '';

    }







document.addEventListener("DOMContentLoaded", function () {
    document.getElementById('generate').addEventListener('click', function (click) {
        click.preventDefault();
        if (!isLogged) {
            return alert('Vous devez être connecté !');
        }
    });
    let acroCookie = getCookie('acronyme');
    if (acroCookie !== undefined && acroCookie.length === 3) {
        document.getElementById('acronyme').value = acroCookie;
    }
    let autoDlBtn = document.getElementById('autoDl');
    let autoDlCookie = getCookie('autoDl');
    if (autoDlCookie !== undefined) {
        autoDl = JSON.parse(autoDlCookie);

    }
    let defaultFamilyCookie = getCookie('defaultFamily');
    if (defaultFamilyCookie !== undefined && EventFilter.saveFamily(defaultFamilyCookie)) {
        defaultFamily = defaultFamilyCookie;
        document.getElementById('defaultFamily' + defaultFamilyCookie).setAttribute('selected', 'selected');
    }

    let defaultFamilySelect = document.getElementById('defaultFamily');
    defaultFamilySelect.addEventListener('change', function () {
        defaultFamily = EventFilter.saveFamily(this.value);
        EventFilter.saveFamilyCookie(this.value);
    });
    if (autoDl) {
        autoDlBtn.classList.add('btn-outline-danger');
        autoDlBtn.innerHTML = 'Désactiver Téléchargement automatique';
    } else {
        autoDlBtn.classList.add('btn-outline-success');
        autoDlBtn.innerHTML = 'Activer Téléchargement automatique';
    }





    // gapi.load("client:auth2", function() {
    //     gapi.auth2.init({
    //         apiKey: API_KEY,
    //         client_id: CLIENT_ID,
    //         discoveryDocs: DISCOVERY_DOCS,
    //         scope: SCOPES
    //     });
    // });
    // document.getElementById('Gauth').addEventListener('click', function () {
    //     authenticate().then(loadClient)
    //
    // });




    autoDlBtn.addEventListener('click', function () {
        if (autoDl) {
            this.classList.add('btn-outline-success');
            this.classList.remove('btn-outline-danger');
            autoDl = false;
            document.cookie = 'autoDl=false; expires=Fri, 31 Dec 2030 23:59:59 GMT';
            this.innerHTML = 'Activer Téléchargement automatique';
        } else {
            this.classList.add('btn-outline-danger');
            this.classList.remove('btn-outline-success');
            autoDl = true;
            document.cookie = 'autoDl=true; expires=Fri, 31 Dec 2030 23:59:59 GMT';
            this.innerHTML = 'Désactiver Téléchargement automatique';
        }
    })
});

// var endTable = '';
isLogged = false;
var getCookie = function (name) {
    var value = "; " + document.cookie;
    var parts = value.split("; " + name + "=");
    if (parts.length == 2) return parts.pop().split(";").shift();
};

function onSignIn(googleUser) {
    // Useful data for your client-side scripts:
    var profile = googleUser.getBasicProfile();
    let acroCookie = getCookie('acronyme');
    if (acroCookie === undefined || acroCookie.length !== 3) {
        let acronyme = profile.getGivenName().substr(0, 1).toUpperCase() + profile.getFamilyName().substr(0, 2).toUpperCase();
        document.getElementById('acronyme').value = acronyme;
        document.cookie = 'acronyme=' + acronyme + '; expires=Fri, 31 Dec 2030 23:59:59 GMT';
    }
    // document.getElementById('connectGoogle').classList.add('d-none');
    let userDiv = document.getElementById('UserDiv');
    userDiv.classList.remove('d-none');
    let userName = document.getElementById('user-name');
    userName.innerHTML = 'Hello ' + profile.getGivenName();
    let userImage = document.getElementById('user-image');
    userImage.setAttribute('src', profile.getImageUrl());
    userImage.style.maxWidth = '50px';
    userImage.style.maxHeight = '50px';
    // let signOut = document.getElementById('signOut');
    // signOut.addEventListener('click', signOut);
    // signOut.classList.remove('d-none');

    // The ID token you need to pass to your backend:
    // var id_token = googleUser.getAuthResponse().id_token;
    isLogged = true;
}

// function signOut() {
//     var auth2 = gapi.auth2.getAuthInstance();
//     auth2.signOut().then(function () {
//     });
// }

/**
 *  On load, called to load the auth2 library and API client library.
//  */
// function handleClientLoad() {
//
//     let weekInput = document.getElementById('week');
//     let current_week = getWeekNumber((new Date()))[1];
//     weekInput.value = current_week;
//
//
//
//
//     let generateBtn = document.getElementById('generate');
//
//     generateBtn.addEventListener('click', function (event) {
//         event.preventDefault();
//         if (!isLogged) {
//             return alert('Vous devez être connecté !')
//         }
//         let val = weekInput.value;
//         let isoDate;
//         if (val !== '' && parseInt(val) > 0) {
//             isoDate = getDateOfISOWeek(val, (new Date()).getFullYear());
//         } else {
//             let time = getWeekNumber((new Date()));
//             isoDate = getDateOfISOWeek(time[1], time[0])
//         }
//         listUpcomingEvents(isoDate);
//     });
//
//
// }







function replaceForObjectName(name) {
    name = name.replace(/[^\w]/gm, 'x');
    return name;
}

    /**
     *
     * @param week
     * @param classStorage class -> create class.agenda_api_events = [events]
     */
    function getEvents(week, classStorage) {
    if (week !== '' && parseInt(week) > 0) {
        week = getDateOfISOWeek(week, (new Date()).getFullYear());
    } else {
        let time = getWeekNumber((new Date()));
        week = getDateOfISOWeek(time[1], time[0])
    }
    let first;
    let end;
    if (week === false) {
        first = (new Date()).toISOString();
        end = (new Date(((new Date().getTime()) + 432000000))).toISOString();
    } else {
        first = week.toISOString();
        end = (new Date((week.getTime() + 432000000))).toISOString();
    }
    week = getWeekNumber((new Date(week)))[1];
    // let projects = {};
    // let projectDontProccess = [];

    gapi.client.calendar.events.list({
        'calendarId': 'primary',
        'timeMin': first,
        'timeMax': end,
        'showDeleted': false,
        'singleEvents': true,
        // 'maxResults': 20,
        'orderBy': 'startTime'
    }).then(function (response) {
        console.log('response events', response);
        // let fake_dha = {
        //     projects: [],
        //     errorProjects: [],
        // };
        // for (let i = 0; i < response.result.items.length; i++) {
        //     let event = response.result.items[i];
        //     let eventFiltered = new EventFilter(event, 'Mise en production', fake_dha);
        //     console.log('event filtered', eventFiltered);
        // }
        classStorage.agenda_api_events = response.result.items;
        console.log('eventsss', classStorage);
    });
}

/**
 * Print the summary and start datetime/date of the next ten events in
 * the authorized user's calendar. If no events are found an
 * appropriate message is printed.
 */
function listUpcomingEvents(week = false) {
    let first;
    let end;
    if (week === false) {
        first = (new Date()).toISOString();
        end = (new Date(((new Date().getTime()) + 432000000))).toISOString();
    } else {
        first = week.toISOString();
        end = (new Date((week.getTime() + 432000000))).toISOString();
    }
    week = getWeekNumber((new Date(week)))[1];
    // let projects = {};
    // let projectDontProccess = [];

    gapi.client.calendar.events.list({
        'calendarId': 'primary',
        'timeMin': first,
        'timeMax': end,
        'showDeleted': false,
        'singleEvents': true,
        // 'maxResults': 20,
        'orderBy': 'startTime'
    }).then(function (response) {
        var events = response.result.items;
        if (events.length > 0) {

            let acronymeInput = document.getElementById('acronyme');
            let acronyme = acronymeInput.value.trim().toUpperCase();
            if (acronyme.length !== 3) {
                removeLoader();
                return alert('Votre acronyme doit contenir UNIQUEMENT 3 Lettres')
            }

            let dha = new DhaBuilder(week, acronyme, events, defaultFamily);
            div_err = document.getElementById('err');
            let div_full = document.getElementById('full');
            // empty(div_full);
            document.getElementById('div_DHA').classList.remove('d-none');
            removeLoader()

        } else {
            removeLoader();
            return alert('Vous n\'avez rien pour cette semaine dans votre agenda !')
        }
    }).catch(function (error) {
        removeLoader();
        console.log('error', error);
        let error_message = '';
        if (error.result !== undefined && error.result.error !== undefined) {

            error_message = checkError(error.result.error);
        } else {
            error_message = 'Oups, il y a eu une erreur : \n'
                + 'ERREUR : ' + error;
        }
        alert(error_message);
    });

}

function checkError(error) {
    switch (error.code) {
        case -1:
            return 'S\'il vous plait vérifiez votre connexion';
        case 400:
            return 'La requête HTTP n\'a pas pu être comprise par le serveur en raison d\'une syntaxe erronée.';
        case 401:
            return 'La requête nécessite une identification de l\'utilisateur.';
        case 402:
            return 'Ce code n\'est pas encore mis en oeuvre dans le protocole HTTP.';
        case 403:
            return 'Google API a compris la requête, mais refuse de la traiter.\nLe document semble inaccessible';
        case 404:
            return 'Google API n\'a rien trouvé qui corresponde à l\'adresse (URI) demandée ( non trouvé ).';
        case 405:
            return 'Ce code indique que la méthode utilisée par le client n\'est pas supportée pour cet URI.';
        case 500:
            return 'Google API a rencontré une condition inattendue qui l\'a empêché de traiter la requête.';
        case 501:
            return 'Google API ne supporte pas la fonctionnalité nécessaire pour traiter la requête.';
        case 503:
            return 'Google API est actuellement incapable de traiter la requête en raison d\'une surcharge temporaire ou d\'une opération de maintenance.';
        case 505:
            return 'La version du protocole HTTP utilisée dans cette requête n\'est pas (ou plus) supportée par le serveur.';
    }
    return 'Oups, il y a eu une erreur : \n'
        + 'Code d\'erreur : ' + error.code.toString() + '\n' +
        'Status : ' + error.status + '\n' +
        'Message : ' + error.message + '\n';
}

function removeLoader() {
    document.getElementById('loader').classList.add('d-none');
    document.getElementById('generate').classList.remove('d-none');
}
function startLoader() {
    document.getElementById('loader').classList.remove('d-none');
    document.getElementById('generate').classList.add('d-none');
}

function empty(div) {
    for (let i = 0; i < div.children.length; i++) {
        div.removeChild(div.children[i]);
    }
}

// function parse_family(family) {
//     family = family.toUpperCase();
//     if (family.match(/MEP|Mises?\s?en\s?Productions?/i)) {
//         return 'Mise en production';
//     } else if (family.match(/Conceptions?/i)) {
//         return 'Conception';
//     } else if (family.match(/Graphismes?/i)) {
//         return 'Graphisme';
//     } else if (family.match(/Pilotages?\s?de\s?projets?/i)) {
//         return 'Pilotage de projet';
//     } else if (family.match(/Directions?|Conseils?|[ée]ditos?/i)) {
//         return 'Direction conseil, technique et éditoriale';
//     } else if (family.match(/Formation/i)) {
//         return 'Formation';
//     } else if (family.match(/Contenus|CM/i)) {
//         return 'Contenus et CM';
//     } else if (family.match(/Maintenance|interventions/i)) {
//         return 'Maintenance et interventions post-projet';
//     } else {
//         return false;
//     }
// }



/* For a given date, get the ISO week number
 *
 * Based on information at:
 *
 *    http://www.merlyn.demon.co.uk/weekcalc.htm#WNR
 *
 * Algorithm is to find nearest thursday, it's year
 * is the year of the week number. Then get weeks
 * between that date and the first day of that year.
 *
 * Note that dates in one year can be weeks of previous
 * or next year, overlap is up to 3 days.
 *
 * e.g. 2014/12/29 is Monday in week  1 of 2015
 *      2012/1/1   is Sunday in week 52 of 2011
 */
function getWeekNumber(d) {
    // Copy date so don't modify original
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    // Set to nearest Thursday: current date + 4 - current day number
    // Make Sunday's day number 7
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    // Get first day of year
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));

    // Calculate full weeks to nearest Thursday
    var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    // Return array of year and week number
    return [d.getUTCFullYear(), weekNo];
}

function getDateOfISOWeek(w, y) {
    var simple = new Date(y, 0, 1 + (w - 1) * 7);
    var dow = simple.getDay();
    var ISOweekStart = simple;
    if (dow <= 4)
        ISOweekStart.setDate(simple.getDate() - simple.getDay() + 1);
    else
        ISOweekStart.setDate(simple.getDate() + 8 - simple.getDay());
    return ISOweekStart;
}

tippy('#exceptions');
tippy('#change_mode');
tippy('#settings');
// tippy('.generate_agenda_all');
tippy('.agenda_generate_btn');
// let btns = document.getElementsByClassName('agenda_generate_btn');
//     for (let i = 0; i <btns.length ; i++) {
//         if ()
//     }
let exceptionsHTML = document.getElementById('exceptions');

let exceptions = getCookie('exceptions');

if (exceptions !== undefined) {
    exceptionsHTML.value = JSON.parse(exceptions);
} else {
    exceptionsHTML.value = 'e-commerce';
}

exceptionsHTML.addEventListener('change', function () {
    document.cookie = 'exceptions=' + JSON.stringify(this.value) + '; expires=Fri, 31 Dec 2030 23:59:59 GMT';
});

/**
 * Simulate a click event.
 * @public
 * @param {Element} elem  the element to simulate a click on
 */
var simulateClick = function (elem) {
    // Create our event (with options)
    var evt = new MouseEvent('click', {
        bubbles: true,
        cancelable: true,
        view: window
    });
    // If cancelled, don't dispatch our event
    var canceled = !elem.dispatchEvent(evt);
};


function test() {
    localStorage.clear();
    var sample_url = "https://spreadsheets.google.com/pub?key=0Ago31JQPZxZrdHF2bWNjcTJFLXJ6UUM5SldEakdEaXc&hl=en&output=html";
    var url_parameter = document.location.search.split(/\?url=/)[1]
    var url = url_parameter || sample_url;
    var googleSpreadsheet = new GoogleSpreadsheet();
    googleSpreadsheet.url(url);
    googleSpreadsheet.load(function(result) {
        $('#results').html(JSON.stringify(result).replace(/,/g,",\n"));
    });
}







    // // Client ID and API key from the Developer Console
    // var CLIENT_ID = '889910005804-tq4btqe01v6sapk7fk65telp3kcdle3p.apps.googleusercontent.com';
    // var API_KEY = 'AIzaSyD42yhsKM9BM71lbkWy1LeX7m68oKxKCbw';
    //
    // // Array of API discovery doc URLs for APIs used by the quickstart
    // var DISCOVERY_DOCS = ["https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest", "https://sheets.googleapis.com/$discovery/rest?version=v4"];
    //
    // // Authorization scopes required by the API; multiple scopes can be
    // // included, separated by spaces.
    // var SCOPES = "https://www.googleapis.com/auth/calendar.readonly https://www.googleapis.com/auth/spreadsheets.readonly";

    var authorizeButton = document.getElementById('authorize_button');
    var signoutButton = document.getElementById('signout_button');
    var connexion_status = document.getElementById('connexion_status');
    tippy(connexion_status);
    /**
     *  On load, called to load the auth2 library and API client library.
     */
    // function handleClientLoad() {
    //     gapi.load('client:auth2', initClient);
    // }
    function handleClientLoad() {
        gapi.load('client:auth2', initClient);

        let generateBtn = document.getElementById('generate');
        let weekInput = document.getElementById('week');
        let current_week = getWeekNumber((new Date()))[1];
        weekInput.value = current_week;
        let loader = document.getElementById('loader');
        generateBtn.addEventListener('click', function (event) {
            event.preventDefault();
            // loader = true;
            startLoader();
            emptyTables();
            switch (mode) {
                case MODE.DHA:
                    generateDHA();
                    break;
                case MODE.AGENDA:
                    let agenda_generator = new AgendaGenerator();
                    agenda_generator.createTable();
                    let val = document.getElementById('week').value;
                 
                    getEvents(val, agenda_generator);
                    break;
            }

        });
    }

    function generateDHA() {

        let val = document.getElementById('week').value;
        let isoDate;
        if (val !== '' && parseInt(val) > 0) {
            isoDate = getDateOfISOWeek(val, (new Date()).getFullYear());
        } else {
            let time = getWeekNumber((new Date()));
            isoDate = getDateOfISOWeek(time[1], time[0])
        }
        listUpcomingEvents(isoDate);
    }

    /**
     *  Initializes the API client library and sets up sign-in state
     *  listeners.
     */
    function initClient() {
        gapi.client.init({
            apiKey: API_KEY,
            clientId: CLIENT_ID,
            discoveryDocs: DISCOVERY_DOCS,
            scope: SCOPES
        }).then(function () {
            // Listen for sign-in state changes.
            gapi.auth2.getAuthInstance().isSignedIn.listen(updateSigninStatus);

            // Handle the initial sign-in state.
            updateSigninStatus(gapi.auth2.getAuthInstance().isSignedIn.get());
            authorizeButton.onclick = handleAuthClick;
            signoutButton.onclick = handleSignoutClick;
        }, function(error) {
            // appendPre(JSON.stringify(error, null, 2));
            console.log('ERROR ', JSON.stringify(error, null, 2));
        });
    }

    /**
     *  Called when the signed in status changes, to update the UI
     *  appropriately. After a sign-in, the API is called.
     */
    function updateSigninStatus(isSignedIn) {
        if (isSignedIn) {
            console.log('gapi', gapi, gapi.auth2);
            onSignIn(gapi.auth2.getAuthInstance().currentUser.get());
            isLogged = true;
            document.getElementById('generate').removeAttribute('disabled');
            authorizeButton.style.display = 'none';
            signoutButton.style.display = 'block';
            connexion_status.style.display = 'none'
        } else {
            authorizeButton.style.display = 'block';
            signoutButton.style.display = 'none';
            connexion_status.style.display = 'none'
        }
    }

    /**
     *  Sign in the user upon button click.
     */
    function handleAuthClick(event) {
        gapi.auth2.getAuthInstance().signIn();
    }

    /**
     *  Sign out the user upon button click.
     */
    function handleSignoutClick(event) {
        gapi.auth2.getAuthInstance().signOut();
    }
