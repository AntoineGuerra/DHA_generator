// Client ID and API key from the Developer Console
var CLIENT_ID = '889910005804-dsb08b9as3lg1ca8vvn726t8cc6946hs.apps.googleusercontent.com';
var API_KEY = 'AIzaSyBVXv2KFsPEzRtUCHrKOJdaN1avCsvppnk';

// Array of API discovery doc URLs for APIs used by the quickstart
var DISCOVERY_DOCS = ["https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest"];

// Authorization scopes required by the API; multiple scopes can be
// included, separated by spaces.
var SCOPES = "https://www.googleapis.com/auth/calendar.readonly";

var authorizeButton = document.getElementById('authorize_button');
var signoutButton = document.getElementById('signout_button');

/**
 *  On load, called to load the auth2 library and API client library.
 */
function handleClientLoad() {
    gapi.load('client:auth2', initClient);
    let weekInput = document.getElementById('week');
    weekInput.addEventListener('input', function () {
        let val = this.value;
        console.log('val', val);
        let isoDate;
        if (val !== '' && parseInt(val) > 0) {
            isoDate = getDateOfISOWeek(val, (new Date()).getFullYear());
            console.log('iso date', isoDate);
            console.log('iso date', isoDate.toISOString());
            // listUpcomingEvents(val);

        } else {
            let time = getWeekNumber((new Date()));
            isoDate = getDateOfISOWeek(time[1], time[0])
            // console.log('date by week', getDateOfISOWeek(time[1], time[0]));
        }

        listUpcomingEvents(isoDate);
    });
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
        appendPre(JSON.stringify(error, null, 2));
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
        listUpcomingEvents();
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
function appendPre(message) {
    var pre = document.getElementById('content');
    var textContent = document.createTextNode(message + '\n');
    pre.appendChild(textContent);
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
    gapi.client.calendar.events.list({
        'calendarId': 'primary',
        'timeMin': first,
        'timeMax': end,
        'showDeleted': false,
        'singleEvents': true,
        'maxResults': 20,
        'orderBy': 'startTime'
    }).then(function(response) {
        var events = response.result.items;
        appendPre('Upcoming events:');

        if (events.length > 0) {
            let text = 'Déclaration Hebdomadaire d\'Activité (DHA);;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
                ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
                'Nombre d\'heures de congés de la semaine;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
                ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
                ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
                ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
                'Temps passé sur des projets vendus (IMPORTANT POUR LES NOUVEAUX PROJETS ! Pour aider les Chefs de Projet, merci de reprendre l\'intitulé exact de la tâche tel que défini dans le CIP !);;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
                ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
                'QUI ?;N° SEM;CLIENT;PROJET;FAMILLE ACTIVITÉ;DETAIL DE LA TÂCHE ;Heures réellement réalisées;Commentaire sur l\'écart;;;;;;;;;;;;;;;;;;;;;;;';
            let text_AV = '';
            let text_V = '';
            let text_M = '';
            let text_I = '';
            let text_content = '';
            for (i = 0; i < events.length; i++) {
                var event = events[i];
                var when = event.start.dateTime;
                var to = event.end.dateTime;
                if (!when) {
                    when = event.start.date;
                }
                if (!to) {
                    to = event.end.date
                }
                appendPre(event.summary + ' By (' + when + ')');
                let explodedText = event.summary.split('-');
                let category;
                let client;
                let project;
                let family;
                let objName;
                let duration = (((new Date(to)).getTime() - (new Date(when)).getTime() ) / 3600000);
                if (explodedText[0] !== undefined && explodedText[1] !== undefined && explodedText[2] !== undefined) {
                    category = explodedText[0].trim();
                    let client_arr = explodedText[1].split('_');
                    objName = explodedText[1].trim();
                    family = explodedText[2].trim();
                    if (client_arr[0] !== undefined && client_arr[1] !== undefined) {
                        client = client_arr[0].trim();
                        project = client_arr[1].trim();
                    } else {
                        continue;
                    }
                } else {
                    continue;
                }
                objName = category + '-' + client + '_' + project + '-' + family;
                console.log('projobjname', projects[objName]);
                if (projects[objName] !== undefined) {
                    projects[objName].duration += duration;
                    console.log('add duration');
                } else {
                    projects[objName] = {
                            category: category,
                            client: client,
                            project: project,
                            family: family,
                            duration: duration
                    };
                    console.log('add proj', objName);
                }

                // text_content += 'AGU;S' + week + ';' + client + ';' + project + ';' + family + ';detail;' + duration.toString().replace('.', ',') + ';commentaire;;;;;;;;;;;;;;;;;;;;;;;\n';
                // console.log('text content', text_content);
                // eval('text_' + category + ' += ' + text_content);

                // console.log('explodedText', explodedText[1].trim());
                console.log('duration', new Date(to).getTime());
                appendPre(event.summary + ' To (' + to + ')' + ' duration (' + duration + ')')
            }
            console.log('project', projects);
            for (var key in projects) {
                // skip loop if the property is from prototype
                if (!projects.hasOwnProperty(key)) continue;

                var obj = projects[key];
                // eval('text_' + obj.category + ' += "' + text_content + '"');
                text_content = 'AGU;S' + week + ';' + obj.client + ';' + obj.project + ';' + obj.family + ';detail;' + obj.duration.toString().replace('.', ',') + ';commentaire;;;;;;;;;;;;;;;;;;;;;;;\n';
                console.log('obj ', key, obj.category);
                switch (obj.category) {
                    case 'V':
                    case 'v':
                        text_V += text_content;
                        console.log('add v');
                        break;
                    case 'AV':
                    case 'Av':
                    case 'av':
                    case 'aV':
                        text_AV += text_content;
                        console.log('add av');
                        break;
                    case 'I':
                    case 'i':
                        text_I += text_content;
                        console.log('add i');
                        break;
                    case 'M':
                    case 'm':
                        text_M += text_content;
                        console.log('add m');
                        break;

                }
                // for (var prop in obj) {
                //     // skip loop if the property is from prototype
                //     if(!obj.hasOwnProperty(prop)) continue;
                //     // your code
                //     alert(prop + " = " + obj[prop]);
                // }
            }
            console.log('text_v', text_V);
            console.log('text_AV', text_AV);
            console.log('text_I', text_I);
            console.log('text_M', text_M);
            appendPre('\n\n\n');
        } else {
            appendPre('No upcoming events found.');
        }
    });
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
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
    // Get first day of year
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
    console.log('year start', yearStart);
    // Calculate full weeks to nearest Thursday
    var weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7);
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

function download(filename, text) {
    var element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
    element.setAttribute('download', filename);

    element.style.display = 'none';
    document.body.appendChild(element);

    element.click();

    document.body.removeChild(element);
}
// var result = getWeekNumber(new Date());
// document.write('It\'s currently week ' + result[1] + ' of ' + result[0]);