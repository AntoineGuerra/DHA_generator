

// Array of API discovery doc URLs for APIs used by the quickstart
// var DISCOVERY_DOCS = ["https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest"];
var DISCOVERY_DOCS = ["https://www.googleapis.com/calendar/v3/calendars/primary/events"];
// var DISCOVERY_DOCS = ["https://www.googleapis.com/calendar/v3/users/me/calendarList"];

// Authorization scopes required by the API; multiple scopes can be
// included, separated by spaces.
var SCOPES = [
    "https://www.googleapis.com/auth/plus.login",
    "https://www.googleapis.com/auth/calendar",
    // "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/calendar.events",
    // "https://www.googleapis.com/auth/calendar.events.readonly",
];
SCOPES = "https://www.googleapis.com/auth/calendar.events";
    autoDl = false;
defaultFamily = 'MEP';





function authenticate() {
    // SEPARATOR -> SPACE ->  + " " +
    return gapi.auth2.getAuthInstance()
        .signIn({
            scope: "https://www.googleapis.com/auth/calendar" + " " +
                "https://www.googleapis.com/auth/calendar.events" + " " +
                "https://www.googleapis.com/auth/calendar.events.readonly" + " " +
                "https://www.googleapis.com/auth/calendar.readonly"
        }).then(function(response) {
            console.log("Sign-in successful"); 
//             console.log('response auth', response);
            if (response.w3 !== undefined) {
                let user = response.w3;
                if (user.ofa && user.wea) {
                    document.getElementById('acronyme').value = (user.ofa.substr(0, 1) + user.wea.substr(0, 2)).toUpperCase();
                }
//                 console.log('response.w3', user.ofa.substr(0, 1), user.wea.substr(0, 2));
            }
            isLogged = true;
            let connectBtn = document.getElementById('Gauth');
            connectBtn.addEventListener('click', function (click) {
                click.preventDefault();
            });
            connectBtn.classList.add('disabled');
            connectBtn.setAttribute('disabled', true);
            connectBtn.innerText = 'Connecté'
            },
            function(err) { console.error("Error signing in", err); });
}
function loadClient() {
    // handleClientLoad();
    // isLogged = true;
    return gapi.client.load("https://content.googleapis.com/discovery/v1/apis/calendar/v3/rest")
        .then(function() {
//             console.log("GAPI client loaded for API", gapi.client);
                handleClientLoad()
            },
            function(err) { console.error("Error loading GAPI client for API", err); });
}






// var authorizeButton = document.getElementById('authorize_button');
// var signoutButton = document.getElementById('signout_button');
document.addEventListener("DOMContentLoaded", function () {
    document.getElementById('generate').addEventListener('click', function (click) {
        click.preventDefault();
    });
    let acroCookie = getCookie('acronyme');
    if (acroCookie !== undefined && acroCookie.length === 3) {
        document.getElementById('acronyme').value = acroCookie;
    }
    let autoDlBtn = document.getElementById('autoDl');
    let autoDlCookie = getCookie('autoDl');
    if (autoDlCookie !== undefined) {
        autoDl = JSON.parse(autoDlCookie);
//         console.log('cookie', getCookie('autoDl'));

    }
    let defaultFamilyCookie = getCookie('defaultFamily');
//     console.log('condition cookie', (EventFilter.saveFamily(defaultFamilyCookie) ? true : false));
    if (defaultFamilyCookie !== undefined && EventFilter.saveFamily(defaultFamilyCookie)) {
        defaultFamily = defaultFamilyCookie;
        document.getElementById('defaultFamily' + defaultFamilyCookie).setAttribute('selected', 'selected');
    }

    let defaultFamilySelect = document.getElementById('defaultFamily');
    defaultFamilySelect.addEventListener('change', function () {
//         console.log('change', this.value);
        defaultFamily = EventFilter.saveFamily(this.value)
        EventFilter.saveFamilyCookie(this.value);
    });
//     console.log('autodl', autoDl, (autoDl == true));
    if (autoDl) {
        autoDlBtn.classList.add('btn-outline-danger');
        autoDlBtn.innerHTML = 'Désactiver Téléchargement automatique'
        // this.classList.remove('btn-success');
    } else {
        autoDlBtn.classList.add('btn-outline-success');
        autoDlBtn.innerHTML = 'Activer Téléchargement automatique'
//         console.log('cookie', getCookie('autoDl'));
        // this.classList.remove('btn-danger');
    }





    gapi.load("client:auth2", function() {
        gapi.auth2.init({client_id: "889910005804-tq4btqe01v6sapk7fk65telp3kcdle3p.apps.googleusercontent.com"});
        //gapi.client.calendar.events.list({
        //             "calendarId": "primary",
        //             "alwaysIncludeEmail": true
        //         })
        //             .then(function(response) {
        //                     // Handle the results here (response.result has the parsed body).
        //                     console.log("Response", response);
        //                 },
        //                 function(err) { console.error("Execute error", err); });
        // onSignIn(gapi.client)
//         console.log('gapi', gapi);
        // authenticate().then(loadClient);
    });
    document.getElementById('Gauth').addEventListener('click', function () {
        authenticate().then(loadClient)

    });




    autoDlBtn.addEventListener('click', function () {
        if (autoDl) {
            this.classList.add('btn-outline-success');
            this.classList.remove('btn-outline-danger');
            autoDl = false;
            document.cookie = 'autoDl=false; expires=Fri, 31 Dec 2030 23:59:59 GMT';
            this.innerHTML = 'Activer Téléchargement automatique'
//             console.log('cookie', getCookie('autoDl'));
        } else {
            this.classList.add('btn-outline-danger');
            this.classList.remove('btn-outline-success');
            autoDl = true;
            document.cookie = 'autoDl=true; expires=Fri, 31 Dec 2030 23:59:59 GMT';
            this.innerHTML = 'Désactiver Téléchargement automatique'
//             console.log('cookie', getCookie('autoDl'));
        }
    })
});

var endTable = '';
isLogged = false;
var getCookie = function (name) {
    var value = "; " + document.cookie;
    var parts = value.split("; " + name + "=");
    if (parts.length == 2) return parts.pop().split(";").shift();
};

function onSignIn(googleUser) {
    // Useful data for your client-side scripts:
    var profile = googleUser.getBasicProfile();
    // console.log("ID: " + profile.getId()); // Don't send this directly to your server!
    // console.log('Full Name: ' + profile.getName());
    // console.log('Given Name: ' + profile.getGivenName());
    // console.log('Family Name: ' + profile.getFamilyName());
//     console.log("Profile: " + profile);
    // console.log("Email: " + profile.getEmail());
    let acronyme = profile.getGivenName().substr(0, 1).toUpperCase() + profile.getFamilyName().substr(0, 2).toUpperCase();
    document.getElementById('acronyme').value = acronyme;
    document.cookie = 'acronyme=' + acronyme + '; expires=Fri, 31 Dec 2030 23:59:59 GMT';
    document.getElementById('connectGoogle').classList.add('d-none');
    let userDiv = document.getElementById('UserDiv');
    userDiv.classList.remove('d-none');
    let userName = document.getElementById('user-name');
    userName.innerHTML = 'Hello ' + profile.getGivenName();
    let userImage = document.getElementById('user-image');
    userImage.setAttribute('src', profile.getImageUrl());
    userImage.style.maxWidth = '50px';
    userImage.style.maxHeight = '50px';
    let signOut = document.getElementById('signOut');
    signOut.addEventListener('click', signOut);
    signOut.classList.remove('d-none');

    // The ID token you need to pass to your backend:
    // var id_token = googleUser.getAuthResponse().id_token;
    // console.log("ID Token: " + id_token);
    isLogged = true;
}

function signOut() {
    var auth2 = gapi.auth2.getAuthInstance();
    auth2.signOut().then(function () {
//         console.log('User signed out.');
    });
}

/**
 *  On load, called to load the auth2 library and API client library.
 */
function handleClientLoad() {

    // gapi.load('client:auth2', initClient);
    // console.log('client init ??????', gapi.load());
    let weekInput = document.getElementById('week');
    // let generateBtn = document.getElementById('generate');
    let current_week = getWeekNumber((new Date()))[1];
    weekInput.value = current_week;




    let generateBtn = document.getElementById('generate');

    generateBtn.addEventListener('click', function (event) {
        event.preventDefault();
        if (!isLogged) {
            return alert('Vous devez être connecté !')
        }
        let val = weekInput.value;
//         console.log('val', val);
        let isoDate;
        if (val !== '' && parseInt(val) > 0) {
            isoDate = getDateOfISOWeek(val, (new Date()).getFullYear());
        } else {
            let time = getWeekNumber((new Date()));
            isoDate = getDateOfISOWeek(time[1], time[0])
        }
        listUpcomingEvents(isoDate);
    });
//


    console.log('test btn', document.getElementById('test'));
    // document.getElementById('test').click(function () {
    //     console.log('tessttt');
    //     var event = {
    //         'summary': 'Google I/O 2015',
    //         'location': '800 Howard St., San Francisco, CA 94103',
    //         'description': 'A chance to hear more about Google\'s developer products.',
    //         'start': {
    //             'dateTime': '2019-02-28T09:00:00-07:00',
    //             'timeZone': 'America/Los_Angeles'
    //         },
    //         'end': {
    //             'dateTime': '2019-02-28T17:00:00-07:00',
    //             'timeZone': 'America/Los_Angeles'
    //         },
    //         'recurrence': [
    //             'RRULE:FREQ=DAILY;COUNT=2'
    //         ],
    //         'attendees': [
    //             {'email': 'lpage@example.com'},
    //             {'email': 'sbrin@example.com'}
    //         ],
    //         'reminders': {
    //             'useDefault': false,
    //             'overrides': [
    //                 {'method': 'email', 'minutes': 24 * 60},
    //                 {'method': 'popup', 'minutes': 10}
    //             ]
    //         }
    //     };
    //
    //     var request = gapi.client.calendar.events.insert({
    //         'calendarId': 'primary',
    //         'resource': event
    //     });
    //
    //     request.execute(function(event) {
    //         appendPre('Event created: ' + event.htmlLink);
    //     });
    // })
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
    }).then(function (response) {
//         console.log('client init', response);
    }, function (error) {
//         console.log(JSON.stringify(error, null, 2));
    });
}

/**
 *  Called when the signed in status changes, to update the UI
 *  appropriately. After a sign-in, the API is called.
 */
function updateSigninStatus(isSignedIn) {
    if (isSignedIn) {
        authorizeButton.style.display = 'none';
        signoutButton.style.display = 'block';
        // listUpcomingEvents();
    } else {
        authorizeButton.style.display = 'block';
        signoutButton.style.display = 'none';
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

/**
 * Append a pre element to the body containing the given message
 * as its text node. Used to display the results of the API call.
 *
 * @param {string} message Text to be placed in pre element.
 */
// function appendPre(message) {
//     var pre = document.getElementById('content');
//     var textContent = document.createTextNode(message + '\n');
//     pre.appendChild(textContent);
// }


function replaceForObjectName(name) {
    name = name.replace(/[^\w]/gm, 'x');
    return name;
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
        // console.log('end', end);
    }
    week = getWeekNumber((new Date(week)))[1];
    let projects = {};
    let projectDontProccess = [];
    // console.log('gapi clien2', gapi.client);
//     console.log('gapi', gapi.client.calendar.events);

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
        // appendPre('Upcoming events:');
//         console.log('events', events);
        if (events.length > 0) {

            let acronymeInput = document.getElementById('acronyme');
            let acronyme = acronymeInput.value.trim().toUpperCase();
            if (!acronyme.match(/[A-Z]{3}/)) {
                return alert('Votre acronyme doit contenir UNIQUEMENT 3 Lettres')
            }

            // let div_err = '';
            // let text_content = '';
            // let text_full = '';
            let dha = new DhaBuilder(week, acronyme, events, defaultFamily);
            div_err = document.getElementById('err');
            //
            // if (projectDontProccess.length > 0) {
            //     document.getElementById('div_err').classList.remove('d-none');
            //     text_content = '';
            //     for (let i = 0; i < projectDontProccess.length; i++) {
            //         let obj = projectDontProccess[i];
            //         let declined = (obj.declined) ? 'Refusé !' : '';
            //         text_content += '<div class="col-12 text-danger ">Tâche : <a href="' + obj.link + '" target="_blank">' + obj.name + '</a> ' + declined + ' Durée : ' + obj.duration + 'H</div>';
            //     }
            //     console.log('text content', text_content);
            //     div_err.innerHTML = '' + text_content + '';
            // } else {
            //     document.getElementById('div_err').classList.add('d-none');
            // }
            let div_full = document.getElementById('full');
            empty(div_full);
            document.getElementById('div_DHA').classList.remove('d-none');


        } else {
            // appendPre('No upcoming events found.');
            return alert('Vous n\'avez rien pour cette semaine dans votre agenda !')
        }
    });

}

function empty(div) {
    for (let i = 0; i < div.children.length; i++) {
        div.removeChild(div.children[i]);
    }
}

function parse_family(family) {
    family = family.toUpperCase();
    if (family.match(/MEP|Mises?\s?en\s?Productions?/i)) {
        return 'Mise en production';
    } else if (family.match(/Conceptions?/i)) {
        return 'Conception';
    } else if (family.match(/Graphismes?/i)) {
        return 'Graphisme';
    } else if (family.match(/Pilotages?\s?de\s?projets?/i)) {
        return 'Pilotage de projet';
    } else if (family.match(/Directions?|Conseils?|[ée]ditos?/i)) {
        return 'Direction conseil, technique et éditoriale';
    } else if (family.match(/Formation/i)) {
        return 'Formation';
    } else if (family.match(/Contenus|CM/i)) {
        return 'Contenus et CM';
    } else if (family.match(/Maintenance|interventions/i)) {
        return 'Maintenance et interventions post-projet';
    } else {
        return false;
    }
}



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
//     console.log('year start', yearStart);
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