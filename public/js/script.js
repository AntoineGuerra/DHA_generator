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
    let generateBtn = document.getElementById('generate');
    let current_week = getWeekNumber((new Date()))[1];
    weekInput.value = current_week;
    let div_full = document.getElementById('full');
    let div_V = document.getElementById('V');
    let div_AV = document.getElementById('AV');
    let div_M = document.getElementById('M');
    let div_I = document.getElementById('I');
    let text_err = document.getElementById('err');
    div_full.parentElement.style.display = 'none';
    div_V.parentElement.style.display = 'none';
    div_AV.parentElement.style.display = 'none';
    div_M.parentElement.style.display = 'none';
    div_I.parentElement.style.display = 'none';
    text_err.parentElement.style.display = 'none';
    generateBtn.addEventListener('click', function () {
        let acronyme = document.getElementById('acronyme').value;
        if (acronyme === '') {
            return alert('l\'acronyme est requis');
        }
        let val = weekInput.value;
        console.log('val', val);
        let isoDate;
        if (val !== '' && parseInt(val) > 0) {
            isoDate = getDateOfISOWeek(val, (new Date()).getFullYear());
            // console.log('iso date', isoDate);
            // console.log('iso date', isoDate.toISOString());
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
        // appendPre('Upcoming events:');

        if (events.length > 0) {
            // let text = 'Déclaration Hebdomadaire d\'Activité (DHA);;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
            //     ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
            //     'Nombre d\'heures de congés de la semaine;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
            //     ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
            //     ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
            //     ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
            //     'Temps passé sur des projets vendus (IMPORTANT POUR LES NOUVEAUX PROJETS ! Pour aider les Chefs de Projet, merci de reprendre l\'intitulé exact de la tâche tel que défini dans le CIP !);;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
            //     ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n' +
            //     'QUI ?;N° SEM;CLIENT;PROJET;FAMILLE ACTIVITÉ;DETAIL DE LA TÂCHE ;Heures réellement réalisées;Commentaire sur l\'écart;;;;;;;;;;;;;;;;;;;;;;;';
            let acronyme = document.getElementById('acronyme').value;
            let part_AV = {
                text: '',
                count: 0,
            };
            let part_V = {
                text: '<Row ss:AutoFitHeight="0" ss:Height="45.75">\n' +
                    '<Cell ss:StyleID="s29"><Data ss:Type="String">Temps passé sur des projets vendus (IMPORTANT POUR LES NOUVEAUX PROJETS ! Pour aider les Chefs de Projet, merci de reprendre l\'intitulé exact de la tâche tel que défini dans le CIP !)</Data></Cell>\n' +
                    '<Cell ss:StyleID="s30"/>\n' +
                    '<Cell ss:StyleID="s30"/>\n' +
                    '<Cell ss:StyleID="s31"/>\n' +
                    '<Cell ss:StyleID="s31"/>\n' +
                    '<Cell ss:StyleID="s31"/>\n' +
                    '<Cell ss:StyleID="s31"/>\n' +
                    '<Cell ss:StyleID="s31"/>\n' +
                    '</Row>\n' +
                    '<Row ss:AutoFitHeight="0" ss:Height="15.75"/>\n' +
                    '<Row ss:AutoFitHeight="0" ss:Height="18">\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462093501864"><Data ss:Type="String">QUI ?</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462093501884"><Data ss:Type="String">N° SEM</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462093501904"><Data ss:Type="String">CLIENT</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462093501924"><Data ss:Type="String">PROJET</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462093501944"><Data ss:Type="String">FAMILLE ACTIVITÉ</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462093501844"><Data ss:Type="String">DETAIL DE LA TÂCHE </Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462093501804"><Data ss:Type="String">Heures réellement réalisées</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462093501824"><Data ss:Type="String">Commentaire sur l\'écart</Data></Cell>\n' +
                    '<Cell ss:StyleID="s18"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '<Cell ss:StyleID="s35"/>\n' +
                    '</Row>',
                count: 0,
            };
            let part_M = {
                text: '<Row ss:AutoFitHeight="0" ss:Height="24"/>\n' +
                    '<Row ss:AutoFitHeight="0" ss:Height="45.75">\n' +
                    '<Cell ss:StyleID="s29"><Data ss:Type="String">Temps passé en maintenance (temps passé = temps déduit des contrats de maintenance)</Data></Cell>\n' +
                    '<Cell ss:StyleID="s29"/>\n' +
                    '<Cell ss:StyleID="s29"/>\n' +
                    '<Cell ss:StyleID="s44"/>\n' +
                    '<Cell ss:StyleID="s44"/>\n' +
                    '<Cell ss:StyleID="s44"/>\n' +
                    '<Cell ss:StyleID="s44"/>\n' +
                    '<Cell ss:StyleID="s44"/>\n' +
                    '</Row>\n' +
                    '<Row ss:AutoFitHeight="0" ss:Height="15.75" ss:Span="1"/>\n' +
                    '<Row ss:Index="27" ss:AutoFitHeight="0" ss:Height="18">\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462106055548"><Data ss:Type="String">QUI ?</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462106055568"><Data ss:Type="String">N° SEM</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462106055428"><Data ss:Type="String">CLIENT</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462106055408"><Data ss:Type="String">PROJET</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462106055588"><Data ss:Type="String">FAMILLE ACTIVITÉ</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462106055608"><Data ss:Type="String">DETAIL DE LA TÂCHE (facultatif)</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462106024352"><Data ss:Type="String">Temps passés en maintenance</Data></Cell>\n' +
                    '<Cell ss:MergeDown="1" ss:StyleID="m140462106055628"><Data ss:Type="String">Tâche programmée ? (non = urgence)</Data></Cell>\n' +
                    '<Cell ss:StyleID="s18"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '<Cell ss:StyleID="s33"/>\n' +
                    '</Row>',
                count: 0,
            };
            let part_I = {
                text: '',
                count: 0,
            };
            let text_err = '';
            let text_content = '';
            let text_full = '';
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
                // appendPre(event.summary + ' By (' + when + ')');
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
                        projectDontProccess.push({
                            name: event.summary,
                            duration: duration
                        });
                        continue;
                    }
                } else {
                    projectDontProccess.push({
                        name: event.summary,
                        duration: duration
                    });
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
                // console.log('duration', new Date(to).getTime());
                // appendPre(event.summary + ' To (' + to + ')' + ' duration (' + duration + ')')
            }
            // console.log('project', projects);
            for (var key in projects) {
                // skip loop if the property is from prototype
                if (!projects.hasOwnProperty(key)) continue;

                var obj = projects[key];
                // eval('text_' + obj.category + ' += "' + text_content + '"');

                // text_content = acronyme.trim().toUpperCase() + ';S' + week + ';' + obj.client.toUpperCase() + ';' +
                //     obj.project.toUpperCase() + ';' + capitalizeFirstLetter(parse_family(obj.family).toLowerCase()) +
                //     ';detail;' + obj.duration.toString().replace('.', ',') +
                //     ';commentaire\n';
                // console.log('obj ', key, obj.category);
                switch (obj.category) {
                    case 'V':
                    case 'v':
                        console.log('add v');

                        text_content = '<Row ss:AutoFitHeight="0" ss:Height="15.75">\n';
                        text_content += '<Cell ss:StyleID="s37"><Data ss:Type="String">' + acronyme.trim() + '</Data></Cell>\n' +
                            '<Cell ss:StyleID="s37"><Data ss:Type="String">S' + week + '</Data></Cell>\n' +
                            '<Cell ss:StyleID="s37"><Data ss:Type="String">' + obj.client.toUpperCase() + '</Data></Cell>\n' +
                            '<Cell ss:StyleID="s37"><Data ss:Type="String">' + obj.project.toUpperCase() + '</Data></Cell>\n' +
                            '<Cell ss:StyleID="s128"><Data ss:Type="String">' + capitalizeFirstLetter(parse_family(obj.family).toLowerCase()) + '</Data></Cell>\n' +
                            '<Cell ss:StyleID="s132"><Data ss:Type="String">detail</Data></Cell>\n' +
                            '<Cell ss:StyleID="s130"><Data ss:Type="Number">' + obj.duration.toString().replace('.', ',') + '</Data></Cell>\n' +
                            '<Cell ss:StyleID="s132"><Data ss:Type="String">commentaire</Data></Cell>\n';
                        for (let i = 0; i <= 22; i++) {
                            text_content += '<Cell ss:StyleID="s35"/>\n';
                        }
                        text_content += '</Row>\n';
                        // text_content = acronyme.trim().toUpperCase() + ';S' + week + ';' + obj.client.toUpperCase() + ';' +
                        //     obj.project.toUpperCase() + ';' + capitalizeFirstLetter(parse_family(obj.family).toLowerCase()) +
                        //     ';detail;' + obj.duration.toString().replace('.', ',') +
                        //     ';commentaire\n';
                        part_V.text += text_content;
                        part_V.count++;
                        break;
                    case 'M':
                    case 'm':
                        console.log('add m');
                        text_content = acronyme.trim().toUpperCase() + ';S' + week + ';' + obj.client.toUpperCase() + ';' +
                            obj.project.toUpperCase() + ';' + capitalizeFirstLetter(parse_family(obj.family).toLowerCase()) +
                            ';detail;' + obj.duration.toString().replace('.', ',') +
                            ';commentaire\n';
                        part_M.text += text_content;
                        part_M.count++;
                        break;
                    case 'AV':
                    case 'Av':
                    case 'av':
                    case 'aV':
                        text_content = acronyme.trim().toUpperCase() + ';S' + week + ';' + obj.client.toUpperCase() +
                            ';' + obj.project.toUpperCase() + ';detail;' +
                            obj.duration.toString().replace('.', ',') + ';commentaire\n';
                        part_AV.text += text_content;
                        part_AV.count++;
                        console.log('add av');
                        break;
                    case 'I':
                    case 'i':
                        text_content = acronyme.trim().toUpperCase() + ';S' + week + ';' + obj.client.toUpperCase() +
                            ';' + obj.project.toUpperCase() + ';detail;' +
                            obj.duration.toString().replace('.', ',') + ';commentaire\n';
                        part_I.text += text_content;
                        part_I.count++;
                        console.log('add i');
                        break;
                }
            }
            text_err = document.getElementById('err');

            if (projectDontProccess.length > 0) {
                text_err.parentElement.style.display = 'block';
                text_content = '';
                for (let i = 0; i < projectDontProccess.length; i++) {
                    let obj = projectDontProccess[i];
                    text_content += 'Nom : ' + obj.name + ' Durée : ' + obj.duration + '\n';
                }
                // text_err = text_content;
                console.log('text content', text_content);
                text_err.innerHTML = '<span style="color: red">' + text_content + '</span>';
            } else {
                text_err.parentElement.style.display = 'none';
            }
            console.log('text_v', part_V);
            console.log('part_AV', part_AV);
            console.log('part_I', part_I);
            console.log('part_M', part_M);
            let div_full = document.getElementById('full');
            empty(div_full);

            let div_V = document.getElementById('V');
            empty(div_V);
            // div_V.innerText = part_V;
            if (part_V.text !== '') {
                // text_full += 'VENDU\n' + 'QUI ?;N SEM;CLIENT;PROJET;FAMILLE ACTIVITE;DETAIL;Temps;Commentaire\n' +
                //     part_V + '\n';
                let txt = '';
                div_V.parentElement.style.display = 'block';
                txt += '<Row ss:AutoFitHeight="0" ss:Height="27">\n' +
                    '<Cell ss:StyleID="s44"/>\n' +
                    '<Cell ss:StyleID="s44"/>\n' +
                    '<Cell ss:StyleID="s44"/>\n' +
                    '<Cell ss:StyleID="s44"/>\n' +
                    '<Cell ss:StyleID="s45"/>\n' +
                    '<Cell ss:StyleID="s45"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                    '<Cell ss:StyleID="s46" ss:Formula="=SUM(R[-' + part_V.count + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                    '<Cell ss:StyleID="s46"/>\n';
                for (let i = 0; i <= 22; i++) {
                    txt += '<Cell ss:StyleID="s35"/>\n';
                }
                    txt += '</Row>\n' + '';
                part_V.text += txt;
                createDownloader('DHA-VENTES-S' + week + '.xml', part_V.text, div_V);
            } else {
                div_V.parentElement.style.display = 'none';

            }

            let div_AV = document.getElementById('AV');
            empty(div_AV);
            // div_AV.innerText = part_AV;
            if (part_AV.text !== '') {
                div_AV.parentElement.style.display = 'block';
                // text_full += 'AVANT VENTE\n' +
                //     'QUI ?;N SEM;CLIENT;PROJET;DETAIL;Temps;Commentaire\n' +
                //     part_AV;
                createDownloader('DHA-AVANT_VENTES-S' + week + '.csv', part_AV, div_AV);
            } else {
                div_AV.parentElement.style.display = 'none';

            }
            let div_M = document.getElementById('M');
            empty(div_M);
            // div_M.innerText = part_M;
            let txt = '';
            if (part_M.text !== '') {
                // text_full += 'MAINTENANCE\n' +
                //     'QUI ?;N SEM;CLIENT;PROJET;FAMILLE ACTIVITE;DETAIL;Temps;Tache programmee ? (non = urgence)\n' +
                //     part_M.text;

                div_M.parentElement.style.display = 'block';
                part_M.text += txt;
                createDownloader('DHA-MAINTENANCES-S' + week + '.csv', part_M.text, div_M);
            } else {
                div_M.parentElement.style.display = 'none';
            }
            let div_I = document.getElementById('I');
            empty(div_I);
            // div_I.innerText = part_I;
            if (part_I.text !== '') {
                div_I.parentElement.style.display = 'block';
                createDownloader('DHA-INTERNES-S' + week + '.csv', part_I, div_I);
                text_full += 'INTERNE\n' +
                    'QUI ?;N SEM;CLIENT;PROJET;DETAIL;Temps;Commentaire\n' +
                    part_I;
            } else {
                div_I.parentElement.style.display = 'none';

            }
            if (part_V !== '' || part_M !== '' || part_I !== '' || part_AV !== '') {
                div_full.parentElement.style.display = 'block';
                createDownloader('DHA-FULL-S' + week + '.csv', text_full, div_full);
            } else {
                div_full.parentElement.style.display = 'none';

            }
            // let div_err = document.getElementById('err');
            // div_err.innerText = text_err;
            // appendPre('\n\n\n');
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
    if (family.match(/MEP|Mise\s?en\s?Production/i)) {
        return 'Mise en production';
    } else if (family.match(/Conception/i)) {
        return 'Conception';
    } else if (family.match(/Graphisme/i)) {
        return 'Graphisme';
    } else if (family.match(/Pilotage\s?de\s?projet|PDP/i)) {
        return 'Pilotage de projet';
    } else if (family.match(/Direction|Conseil|[ée]dito/i)) {
        return 'Direction conseil, technique et éditoriale';
    } else if (family.match(/Formation/i)) {
        return 'Formation';
    } else if (family.match(/Contenus|CM/i)) {
        return 'Contenus et CM';
    } else if (family.match(/Maintenance|interventions|PP/i)) {
        return 'Maintenance et interventions post-projet';
    }
}
function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
}
// function parseValue(value) {
//
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

function createDownloader(filename, text, div) {
    var element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
    element.setAttribute('download', filename);
    element.innerText = filename;
    // element.style.display = 'none';
    div.appendChild(element);

    // element.click();

    // document.body.removeChild(element);
}
// var result = getWeekNumber(new Date());
// document.write('It\'s currently week ' + result[1] + ' of ' + result[0]);