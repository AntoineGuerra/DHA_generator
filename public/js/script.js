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

var endTable = '';




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
    // div_full.parentElement.style.display = 'none';
    // div_V.parentElement.style.display = 'none';
    // div_AV.parentElement.style.display = 'none';
    // div_M.parentElement.style.display = 'none';
    // div_I.parentElement.style.display = 'none';
    // text_err.parentElement.style.display = 'none';
    let acronyme = document.getElementById('acronyme').value;
    // document.getElementById('acronyme').value = 'test';
    // console.log('gapi clien1', gapi);
    checkClient = window.setInterval(function () {
        if (gapi.client !== undefined && gapi.client.calendar !== undefined && gapi.client.calendar.events !== undefined) {
            gapi.client.calendar.events.list({
                'calendarId': 'primary',
                'showDeleted': false,
                'singleEvents': true,
                'maxResults': 1,
                // 'orderBy': 'startTime'
            }).then(function(response) {
                var events = response.result.items;
                if (events[0] !== undefined && events[0].creator !== undefined && events[0].creator.email !== undefined) {
                    let email = events[0].creator.email;
                    let matchs = email.match(/^([a-z]{1})[^\.]*\.([a-z]{2}).*/i)
                    if (matchs) {
                        console.log('match', matchs);
                    }
                    // console.log('ever', events[0].creator.email);
                    if (matchs.length >= 3) {

                        document.getElementById('acronyme').value = (matchs[1] + matchs[2]).toUpperCase();
                    }
                }
            });
            clearInterval(checkClient);
        }
    }, 500);
    // gapi.client.calendar.events.list({
    //     'calendarId': 'primary',
    //     'showDeleted': false,
    //     'singleEvents': true,
    //     'maxResults': 20,
    //     'orderBy': 'startTime'
    // }).then(function(response) {
    //     var events = response.result.items;
    //
    //     console.log('ever', events.creator.email);
    // });
    //
    
    generateBtn.addEventListener('click', function (event) {
        event.preventDefault();
        // if (acronyme === '') {
        //     return alert('l\'acronyme est requis');
        // }
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
    console.log('gapi clien2', gapi.client);
    gapi.client.calendar.events.list({
        'calendarId': 'primary',
        'timeMin': first,
        'timeMax': end,
        'showDeleted': false,
        'singleEvents': true,
        // 'maxResults': 20,
        'orderBy': 'startTime'
    }).then(function(response) {
        var events = response.result.items;
        // appendPre('Upcoming events:');
        console.log('events', events);
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

            let acronymeInput = document.getElementById('acronyme');
            if (acronymeInput.value === '') {

                acronymeInput.value = events.creator.email;
            }
            let acronyme = acronymeInput.value.trim().toUpperCase();
            if (!acronyme.match(/[A-Z]{3}/)) {
                return alert('Votre acronyme doit contenir UNIQUEMENT 3 Lettres')
            }

            let part_V = [];
            let part_V_duration = 0;
            let part_AV = [];
            let part_AV_duration = 0;
            let part_M = [];
            let part_M_duration = 0;
            let part_I = [];
            let part_I_duration = 0;
            let div_err = '';
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
                console.log('exploded text', explodedText);
                let duration = (((new Date(to)).getTime() - (new Date(when)).getTime() ) / 3600000);
                if (explodedText[0] !== undefined && explodedText[1] !== undefined && explodedText[2] !== undefined) {
                    category = explodedText[0].trim().toUpperCase();
                    let client_arr = explodedText[1].split('_');
                    // objName = explodedText[1].trim().toUpperCase();
                    family = explodedText[2].trim().toUpperCase();
                    if ((client_arr[0] !== undefined && client_arr[1] !== undefined)) {
                        client = client_arr[0].trim().toUpperCase();
                        project = client_arr[1].trim().toUpperCase();
                    } else {
                        projectDontProccess.push({
                            name: event.summary,
                            duration: duration,
                        });
                        console.log('non trier ', explodedText, client_arr);
                        continue;
                    }
                } else if (explodedText[0] !== undefined && explodedText[1] !== undefined &&
                    (explodedText[0].trim().toUpperCase() === 'AV' || explodedText[0].trim().toUpperCase() === 'I')) {
                    category = explodedText[0].trim().toUpperCase();
                    let client_arr = explodedText[1].split('_');
                    // objName = explodedText[1].trim().toUpperCase();
                    // family = explodedText[2].trim().toUpperCase();
                    family = '';
                    if ((client_arr[0] !== undefined && client_arr[1] !== undefined)) {
                        client = client_arr[0].trim();
                        project = client_arr[1].trim();
                    } else {
                        projectDontProccess.push({
                            name: event.summary,
                            duration: duration
                        });
                        console.log('non trier 2', explodedText, client_arr);
                        continue;
                    }
                } else {
                    console.log('non trier 3', explodedText);
                    projectDontProccess.push({
                        name: event.summary,
                        duration: duration
                    });
                    continue;
                }
                // if (family === '' || category === 'AV' || category === 'I') {
                //     objName = category.toUpperCase() + '-' + client.toUpperCase() + '_' + project.toUpperCase();
                // } else {
                objName = replaceForObjectName(category) + '-' + replaceForObjectName(client) + '_' + replaceForObjectName(project) + '-' + replaceForObjectName(family);
                // }
                let description = event.description;
                let detail = 'SANS';
                let comment = 'Sans Commentaire';
                if (description !== undefined) {
                    description = description.replace(new RegExp("&nbsp;", 'g'), ' ');
                    let tacheMatches = description.match(/(.*)tache\s?\:\s?<?b?r?>?\s?\<b\>([^<]*)\<\/b\>(.*)/i);
                    if (tacheMatches) {
                        console.log('tache match', tacheMatches);
                        description = tacheMatches[1] + tacheMatches[3];
                        detail = tacheMatches[2];
                    } else {
                        detail = 'SANS';
                    }
                    let commentMatches = description.match(/(.*)Commentaire\s*:\s*<?b?r?>?\s?\<b\>([^<]*)\<\/b\>(.*)/i);
                    console.log('objname', objName);
                    if (category === 'M') {
                        // description = 'detail';
                        let urgentMatches = description.match(/(.*)\<b\>URGENT\<\/b\>(.*)/i);
                        if (urgentMatches) {
                            comment = 'NON';
                            description = urgentMatches[1] + urgentMatches[3];
                            console.log('urgent match');
                        } else {
                            console.log('urgent dont match', );
                            comment = 'OUI';
                        }
                    } else if (commentMatches) {
                        console.log('comment match', commentMatches);
                        description = commentMatches[1] + commentMatches[3];
                        comment = commentMatches[2];
                    } else {
                        comment = 'Sans Commentaire';
                    }
                } else if (category === 'M') {
                    comment = 'OUI';
                }
                console.log('description', description, events);
                console.log('projobjname', projects[objName]);
                if (projects[objName] !== undefined) {
                    if (projects[objName].detail !== detail && projects[objName].detail !== 'SANS' && detail !== 'SANS' && projects[objName + replaceForObjectName(detail)] === undefined) {
                        let message = 'Votre intervention : "' + projects[objName].detail + '" est en conflit avec : "' + detail + '" sur le projet ' + parse_category(category) + ' : ' + toTitleCase(project.toLowerCase()) + ' de ' + toTitleCase(client.toLowerCase()) + ' \n';
                        message += 'Créer deux Tâches Séparé ?';
                        if (confirm(message)) {
                            let randStr = Math.random().toString(36).replace(/[^a-z]+/g, '').substr(0, 3);
                            projects[objName + replaceForObjectName(detail)] = {
                                category: category.toUpperCase(),
                                client: capitalizeFirstLetter(client),
                                project: capitalizeFirstLetter(project),
                                family: family.toUpperCase(),
                                detail: detail,
                                commentaire: comment,
                                duration: duration
                            };
                        } else {
                            if (confirm('Remplacer : "' + projects[objName].detail + '" Par : "' + detail + '"')) {
                                projects[objName].detail = detail;
                            }
                            projects[objName].duration += duration;
                        }
                    } else if (projects[objName + replaceForObjectName(detail)] !== undefined) {
                        projects[objName + replaceForObjectName(detail)].duration += duration;
                    } else if (projects[objName].detail === 'SANS') {
                        projects[objName].detail = detail;
                        projects[objName].duration += duration;
                    } else {
                        projects[objName].duration += duration;
                    }
                    console.log('add duration');
                } else {
                    projects[objName] = {
                            category: category.toUpperCase(),
                            client: capitalizeFirstLetter(client),
                            project: capitalizeFirstLetter(project),
                            family: family.toUpperCase(),
                            detail: detail,
                            commentaire: comment,
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
                    // case 'v':
                        console.log('add v');

                        // text_content = '<Row ss:AutoFitHeight="0" ss:Height="15.75">\n';
                        // text_content += '<Cell ss:StyleID="s37"><Data ss:Type="String">' + acronyme.trim() + '</Data></Cell>\n' +
                        //     '<Cell ss:StyleID="s37"><Data ss:Type="String">S' + week + '</Data></Cell>\n' +
                        //     '<Cell ss:StyleID="s37"><Data ss:Type="String">' + obj.client.toUpperCase() + '</Data></Cell>\n' +
                        //     '<Cell ss:StyleID="s37"><Data ss:Type="String">' + obj.project.toUpperCase() + '</Data></Cell>\n' +
                        //     '<Cell ss:StyleID="s128"><Data ss:Type="String">' + capitalizeFirstLetter(parse_family(obj.family)) + '</Data></Cell>\n' +
                        //     '<Cell ss:StyleID="s132"><Data ss:Type="String">' + removeBalise(obj.detail) + '</Data></Cell>\n' +
                        //     '<Cell ss:StyleID="s130"><Data ss:Type="Number">' + obj.duration.toString().replace('.', ',') + '</Data></Cell>\n' +
                        //     '<Cell ss:StyleID="s132"><Data ss:Type="String">' + removeBalise(obj.commentaire) + '</Data></Cell>\n';
                        // for (let i = 0; i <= 22; i++) {
                        //     text_content += '<Cell ss:StyleID="s35"/>\n';
                        // }
                        // text_content += '</Row>\n';
                        // text_content = acronyme.trim().toUpperCase() + ';S' + week + ';' + obj.client.toUpperCase() + ';' +
                        //     obj.project.toUpperCase() + ';' + capitalizeFirstLetter(parse_family(obj.family).toLowerCase()) +
                        //     ';detail;' + obj.duration.toString().replace('.', ',') +
                        //     ';commentaire\n';
                        part_V.push(obj);
                        part_V_duration += obj.duration;
                        // part_V.text += text_content;
                        // part_V.count++;
                        break;
                    case 'M':
                    // case 'm':
                        console.log('add m');
                        // text_content = acronyme.trim().toUpperCase() + ';S' + week + ';' + obj.client.toUpperCase() + ';' +
                        //     obj.project.toUpperCase() + ';' + capitalizeFirstLetter(parse_family(obj.family).toLowerCase()) +
                        //     ';detail;' + obj.duration.toString().replace('.', ',') +
                        //     ';commentaire\n';
                        // part_M.text += text_content;
                        // part_M.count++;
                        part_M.push(obj);
                        part_M_duration += obj.duration;
                        break;
                    case 'AV':
                    // case 'Av':
                    // case 'av':
                    // case 'aV':
                        // text_content = acronyme.trim().toUpperCase() + ';S' + week + ';' + obj.client.toUpperCase() +
                        //     ';' + obj.project.toUpperCase() + ';detail;' +
                        //     obj.duration.toString().replace('.', ',') + ';commentaire\n';
                        // part_AV.text += text_content;
                        // part_AV.count++;
                        part_AV.push(obj);
                        part_AV_duration += obj.duration;
                        console.log('add av');
                        break;
                    case 'I':
                    // case 'i':
                        // text_content = acronyme.trim().toUpperCase() + ';S' + week + ';' + obj.client.toUpperCase() +
                        //     ';' + obj.project.toUpperCase() + ';detail;' +
                        //     obj.duration.toString().replace('.', ',') + ';commentaire\n';
                        // part_I.text += text_content;
                        // part_I.count++;
                        part_I.push(obj);
                        part_I_duration += obj.duration;
                        console.log('add i');
                        break;
                }
            }
            div_err = document.getElementById('err');

            if (projectDontProccess.length > 0) {
                // div_err.parentElement.style.display = 'block';
                document.getElementById('div_err').classList.remove('d-none');
                text_content = '';
                for (let i = 0; i < projectDontProccess.length; i++) {
                    let obj = projectDontProccess[i];
                    text_content += '<div class="col-6 text-danger text-center">Nom : ' + obj.name + ' Durée : ' + obj.duration + 'H</div>';
                }
                // text_err = text_content;
                console.log('text content', text_content);
                let errMsg = '<span class="text-warning">Votre Projet doit être sous cette forme :</span><br><span class="text-success">V - Client_Projet - MEP (optionnel Pour AV et I)</span>';
                div_err.innerHTML = '<div class="col-6 text-center">' + errMsg + '</div><br>' + text_content + '';
            } else {
                document.getElementById('div_err').classList.add('d-none');
                // div_err.parentElement.style.display = 'none';
            }
            console.log('text_v', part_V);
            console.log('part_AV', part_AV);
            console.log('part_I', part_I);
            console.log('part_M', part_M);
            let div_full = document.getElementById('full');
            empty(div_full);

            // let div_V = document.getElementById('V');
            // empty(div_V);
            // div_V.innerText = part_V;
            // if (part_V !== undefined) {
                // text_full += 'VENDU\n' + 'QUI ?;N SEM;CLIENT;PROJET;FAMILLE ACTIVITE;DETAIL;Temps;Commentaire\n' +
                //     part_V + '\n';
                // let txt = '';
                // div_V.parentElement.style.display = 'block';
                // txt += '<Row ss:AutoFitHeight="0" ss:Height="27">\n' +
                //     '<Cell ss:StyleID="s44"/>\n' +
                //     '<Cell ss:StyleID="s44"/>\n' +
                //     '<Cell ss:StyleID="s44"/>\n' +
                //     '<Cell ss:StyleID="s44"/>\n' +
                //     '<Cell ss:StyleID="s45"/>\n' +
                //     '<Cell ss:StyleID="s45"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                //     // '<Cell ss:StyleID="s46" ss:Formula="=SUM(R[-' + part_V.count + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                //     '<Cell ss:StyleID="s46"/>\n';
                // for (let i = 0; i <= 22; i++) {
                //     txt += '<Cell ss:StyleID="s35"/>\n';
                // }
                //     txt += '</Row>\n' + '';
                // part_V.text += txt;
                // createDownloader('DHA-VENTES-S' + week + '.xml', (part_V.text + endTable), div_V);
            // } else {
            //     div_V.parentElement.style.display = 'none';
            //
            // }

            // let div_AV = document.getElementById('AV');
            // empty(div_AV);
            // div_AV.innerText = part_AV;
            // if (part_AV !== undefined) {
            //     div_AV.parentElement.style.display = 'block';
                // text_full += 'AVANT VENTE\n' +
                //     'QUI ?;N SEM;CLIENT;PROJET;DETAIL;Temps;Commentaire\n' +
                //     part_AV;
                // createDownloader('DHA-AVANT_VENTES-S' + week + '.csv', part_AV, div_AV);
            // } else {
            //     div_AV.parentElement.style.display = 'none';
            //
            // }
            // let div_M = document.getElementById('M');
            // empty(div_M);
            // div_M.innerText = part_M;
            // let txt = '';
            // if (part_M !== undefined) {
            //     // text_full += 'MAINTENANCE\n' +
            //     //     'QUI ?;N SEM;CLIENT;PROJET;FAMILLE ACTIVITE;DETAIL;Temps;Tache programmee ? (non = urgence)\n' +
            //     //     part_M.text;
            //
            //     div_M.parentElement.style.display = 'block';
            //     // part_M.text += txt;
            //     // createDownloader('DHA-MAINTENANCES-S' + week + '.csv', part_M.text, div_M);
            // } else {
            //     div_M.parentElement.style.display = 'none';
            // }
            // let div_I = document.getElementById('I');
            // empty(div_I);
            // // div_I.innerText = part_I;
            // if (part_I !== undefined) {
            //     div_I.parentElement.style.display = 'block';
            //     // createDownloader('DHA-INTERNES-S' + week + '.csv', part_I, div_I);
            //     // text_full += 'INTERNE\n' +
            //     //     'QUI ?;N SEM;CLIENT;PROJET;DETAIL;Temps;Commentaire\n' +
            //     //     part_I;
            // } else {
            //     div_I.parentElement.style.display = 'none';
            //
            // }
            // if (part_V !== '' || part_M !== '' || part_I !== '' || part_AV !== '') {
            //     div_full.parentElement.style.display = 'block';
                // createDownloader('DHA-FULL-S' + week + '.csv', text_full, div_full);
            // } else {
            //     div_full.parentElement.style.display = 'none';
            //
            // }
            let text_test = '<?xml version="1.0"?>\n' +
                '<?mso-application progid="Excel.Sheet"?>\n' +
                '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n' +
                ' xmlns:o="urn:schemas-microsoft-com:office:office"\n' +
                ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n' +
                ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"\n' +
                ' xmlns:html="http://www.w3.org/TR/REC-html40">\n' +
                ' <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">\n' +
                '  <LastAuthor>Antoine Guerra</LastAuthor>\n' +
                '  <Created>2019-01-25T15:49:35Z</Created>\n' +
                '  <LastSaved>2019-01-25T15:49:35Z</LastSaved>\n' +
                '  <Version>16.00</Version>\n' +
                ' </DocumentProperties>\n' +
                ' <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">\n' +
                '  <AllowPNG/>\n' +
                ' </OfficeDocumentSettings>\n' +
                ' <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">\n' +
                '  <WindowHeight>15540</WindowHeight>\n' +
                '  <WindowWidth>25600</WindowWidth>\n' +
                '  <WindowTopX>32767</WindowTopX>\n' +
                '  <WindowTopY>460</WindowTopY>\n' +
                '  <ProtectStructure>False</ProtectStructure>\n' +
                '  <ProtectWindows>False</ProtectWindows>\n' +
                ' </ExcelWorkbook>\n' +
                ' <Styles>\n' +
                '  <Style ss:ID="Default" ss:Name="Normal">\n' +
                '   <Alignment ss:Vertical="Bottom"/>\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <Interior/>\n' +
                '   <NumberFormat/>\n' +
                '   <Protection/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056512">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056532">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056552">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056572">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056592">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056612">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056632">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056192">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056212">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056232">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056252">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056272">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056292">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106056312">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106024352">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106024372">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106024392">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106024412">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106024432">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106024452">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106024472">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106024492">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106024512">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106024532">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055328">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055348">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055368">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055388">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055408">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055428">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055448">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055468">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055488">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055508">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055528">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055548">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055568">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055588">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055608">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="m140462106055628">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s15">\n' +
                '   <Alignment ss:Vertical="Bottom"/>\n' +
                '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s16">\n' +
                '   <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000" ss:Bold="1"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s17">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="22" ss:Color="#000000" ss:Bold="1"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s18">\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s19">\n' +
                '   <Alignment ss:Vertical="Center" ss:WrapText="1"/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s20">\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="11" ss:Color="#000000"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s21">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s22">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="14" ss:Color="#000000" ss:Bold="1"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s23">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="8" ss:Color="#000000" ss:Bold="1"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s24">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#FF9300" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s25">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#000000" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#FF9300" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s26">\n' +
                '   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s27">\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <Interior ss:Color="#FF9300" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s28">\n' +
                '   <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s29">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#9BBB59" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s30">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#000000" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#9BBB59" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s31">\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <Interior ss:Color="#9BBB59" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s32">\n' +
                '   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s33">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000" ss:Bold="1"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s34">\n' +
                '   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s35">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s36">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#3F3F3F"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s37">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s38">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
                '   <Interior ss:Color="#C2D69B" ss:Pattern="Solid"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s39">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s40">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s41">\n' +
                '   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s42">\n' +
                '   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s43">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s44">\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
                '   <Interior ss:Color="#9BBB59" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s45">\n' +
                '   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s46">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s47">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s48">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s49">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#B2A1C7" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s50">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#000000" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#B2A1C7" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s51">\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <Interior ss:Color="#B2A1C7" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s52">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
                '   <Interior ss:Color="#CCC0D9" ss:Pattern="Solid"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s53">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#0080FF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s54">\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
                '   <Interior ss:Color="#0080FF" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s55">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders>\n' +
                '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
                '     ss:Color="#000000"/>\n' +
                '   </Borders>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
                '   <Interior ss:Color="#8DB3E2" ss:Pattern="Solid"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s56">\n' +
                '   <Alignment ss:Vertical="Center"/>\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#000000" ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s57">\n' +
                '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="16" ss:Color="#000000"\n' +
                '    ss:Bold="1"/>\n' +
                '   <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>\n' +
                '   <NumberFormat ss:Format="0.0"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s58">\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
                '   <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>\n' +
                '  </Style>\n' +
                '  <Style ss:ID="s70">\n' +
                '   <Alignment ss:Vertical="Center" ss:WrapText="1"/>\n' +
                '   <Borders/>\n' +
                '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
                '  </Style>\n' +
                ' </Styles>\n' +
                ' <Worksheet ss:Name="DHA">\n' +
                '  <Table ss:ExpandedColumnCount="31" ss:ExpandedRowCount="1000" x:FullColumns="1"\n' +
                '   x:FullRows="1" ss:StyleID="s15" ss:DefaultColumnWidth="67"\n' +
                '   ss:DefaultRowHeight="15">\n' +
                '   <Column ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="63" ss:Span="1"/>\n' +
                '   <Column ss:Index="3" ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="189"\n' +
                '    ss:Span="1"/>\n' +
                '   <Column ss:Index="5" ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="258"/>\n' +
                '   <Column ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="355"/>\n' +
                '   <Column ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="118"/>\n' +
                '   <Column ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="346"/>\n' +
                '   <Column ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="63" ss:Span="22"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="42">\n' +
                '    <Cell ss:StyleID="s17"><Data ss:Type="String">Déclaration Hebdomadaire d\'Activité (DHA)</Data></Cell>\n' +
                emptyCell(17, 26) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                '    <Cell ss:StyleID="s20"/>\n' +
                '    <Cell ss:StyleID="s20"/>\n' +
                emptyCell(22, 25) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="45.75">\n' +
                '    <Cell ss:StyleID="s24"><Data ss:Type="String">Nombre d\'heures de congés de la semaine</Data></Cell>\n' +
                '    <Cell ss:StyleID="s25"/>\n' +
                '    <Cell ss:StyleID="s25"/>\n' +
                '    <Cell ss:StyleID="s27"/>\n' +
                '    <Cell ss:StyleID="s27"/>\n' +
                '    <Cell ss:StyleID="s27"/>\n' +
                '    <Cell ss:StyleID="s27"/>\n' +
                '    <Cell ss:StyleID="s27"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                '    <Cell ss:StyleID="s20"/>\n' +
                '    <Cell ss:StyleID="s20"/>\n' +
                emptyCell(22, 25) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                '    <Cell ss:StyleID="s20"/>\n' +
                '    <Cell ss:StyleID="s20"/>\n' +
                emptyCell(22, 5) +
                '    <Cell ss:StyleID="s23"/>\n' +
                emptyCell(22, 18) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                '    <Cell ss:StyleID="s20"/>\n' +
                '    <Cell ss:StyleID="s20"/>\n' +
                emptyCell(22, 25) +
                '   </Row>\n';
            if (part_V.length > 0) {
                text_test += '   <Row ss:AutoFitHeight="0" ss:Height="25.5"/>\n' +
                    '   <Row ss:AutoFitHeight="0" ss:Height="48">\n' +
                    '    <Cell ss:StyleID="s29"><Data ss:Type="String">Temps passé sur des projets vendus (IMPORTANT POUR LES NOUVEAUX PROJETS ! Pour aider les Chefs de Projet, merci de reprendre l\'intitulé exact de la tâche tel que défini dans le CIP !)</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s30"/>\n' +
                    '    <Cell ss:StyleID="s30"/>\n' +
                    '    <Cell ss:StyleID="s31"/>\n' +
                    '    <Cell ss:StyleID="s31"/>\n' +
                    '    <Cell ss:StyleID="s31"/>\n' +
                    '    <Cell ss:StyleID="s31"/>\n' +
                    '    <Cell ss:StyleID="s31"/>\n' +
                    '   </Row>\n';
                text_test += '   <Row ss:AutoFitHeight="0" ss:Height="18">\n' +
                    '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056232"><Data ss:Type="String">QUI ?</Data></Cell>\n' +
                    '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056252"><Data ss:Type="String">N° SEM</Data></Cell>\n' +
                    '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056272"><Data ss:Type="String">CLIENT</Data></Cell>\n' +
                    '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056292"><Data ss:Type="String">PROJET</Data></Cell>\n' +
                    '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056312"><Data ss:Type="String">FAMILLE ACTIVITÉ</Data></Cell>\n' +
                    '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056212"><Data ss:Type="String">DETAIL DE LA TÂCHE </Data></Cell>\n' +
                    '    <Cell ss:MergeDown="1" ss:StyleID="m140462106024432"><Data ss:Type="String">Heures réellement réalisées</Data></Cell>\n' +
                    '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056192"><Data ss:Type="String">Commentaire sur l\'écart</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s18"/>\n' +
                    emptyCell(33, 18) +
                    '   </Row>\n' +
                    '   <Row ss:AutoFitHeight="0" ss:Height="36.75">\n' +
                    '    <Cell ss:Index="9" ss:StyleID="s18"/>\n' +
                    emptyCell(33, 18) +
                    '    <Cell ss:StyleID="s18"/>\n' +
                    '    <Cell ss:StyleID="s18"/>\n' +
                    '    <Cell ss:StyleID="s18"/>\n' +
                    '   </Row>\n';
                for (let i = 0; i < part_V.length; i++) {
                    let object = part_V[i];
                    text_test += '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                        '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + acronyme.trim() + '</Data></Cell>\n' +
                        '    <Cell ss:StyleID="s34"><Data ss:Type="String">S' + week + '</Data></Cell>\n' +
                        '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + object.client + '</Data></Cell>\n' +
                        '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + object.project + '</Data></Cell>\n' +
                        '    <Cell ss:StyleID="s36"><Data ss:Type="String">' + parse_family(object.family) + '</Data></Cell>\n' +
                        '    <Cell ss:StyleID="s39"><Data ss:Type="String">' + removeBalise(object.detail) + '</Data></Cell>\n' +
                        '    <Cell ss:StyleID="s38"><Data ss:Type="Number">' + object.duration + '</Data></Cell>\n' +
                        '    <Cell ss:StyleID="s39"><Data ss:Type="String">' + removeBalise(object.commentaire) + '</Data></Cell>\n' +
                        '    <Cell ss:StyleID="s23"/>\n' +
                        emptyCell(33, 21) +
                        '   </Row>\n';
                }
                text_test += '   <Row ss:AutoFitHeight="0" ss:Height="27">\n' +
                    '    <Cell ss:StyleID="s41"/>\n' +
                    '    <Cell ss:StyleID="s41"/>\n' +
                    '    <Cell ss:StyleID="s41"/>\n' +
                    '    <Cell ss:StyleID="s41"/>\n' +
                    '    <Cell ss:StyleID="s42"/>\n' +
                    '    <Cell ss:StyleID="s42"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s43" ss:Formula="=SUM(R[-' + part_V.length + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                    '    <Cell ss:StyleID="s43"/>\n' +
                    '    <Cell ss:StyleID="s23"/>\n' +
                    emptyCell(33, 21) +
                    '   </Row>\n';
            }
            // if (part_M.length > 0) {
            text_test += '   <Row ss:AutoFitHeight="0" ss:Height="24"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="45.75">\n' +
                '    <Cell ss:StyleID="s29"><Data ss:Type="String">Temps passé en maintenance (temps passé = temps déduit des contrats de maintenance)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s29"/>\n' +
                '    <Cell ss:StyleID="s29"/>\n' +
                '    <Cell ss:StyleID="s44"/>\n' +
                '    <Cell ss:StyleID="s44"/>\n' +
                '    <Cell ss:StyleID="s44"/>\n' +
                '    <Cell ss:StyleID="s44"/>\n' +
                '    <Cell ss:StyleID="s44"/>\n' +
                '   </Row>\n' +
                // '   <Row ss:AutoFitHeight="0" ss:Height="15.75" ss:Span="1"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18">\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106055548"><Data ss:Type="String">QUI ?</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106055568"><Data ss:Type="String">N° SEM</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106055428"><Data ss:Type="String">CLIENT</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106055408"><Data ss:Type="String">PROJET</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106055588"><Data ss:Type="String">FAMILLE ACTIVITÉ</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106055608"><Data ss:Type="String">DETAIL DE LA TÂCHE (facultatif)</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106024352"><Data ss:Type="String">Temps passés en maintenance</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106055628"><Data ss:Type="String">Tâche programmée ? (non = urgence)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                emptyCell(33, 21) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="36.75">\n' +
                '    <Cell ss:Index="9" ss:StyleID="s18"/>\n' +
                emptyCell(33, 21) +
                '   </Row>\n';
            for (let i = 0; i < part_M.length; i++) {
                let object = part_M[i];
                text_test += '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                    '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + acronyme.trim() + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s34"><Data ss:Type="String">S' + week + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + object.client + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + object.project + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s36"><Data ss:Type="String">' + parse_family(object.family) + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s39"><Data ss:Type="String">' + removeBalise(object.detail) + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s38"><Data ss:Type="Number">' + object.duration + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="m140462106055528"><Data ss:Type="String">' + removeBalise(object.commentaire) + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s23"/>\n' +
                    emptyCell(33, 21) +
                    '   </Row>\n';
            }
            text_test += '   <Row ss:AutoFitHeight="0" ss:Height="27">\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s42"/>\n' +
                '    <Cell ss:StyleID="s42"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s43" ss:Formula="=SUM(R[-' + part_M.length + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                '    <Cell ss:StyleID="s48"/>\n' +
                '    <Cell ss:StyleID="s23"/>\n' +
                emptyCell(33, 21) +
                '   </Row>\n';
            // }
            text_test += '   <Row ss:AutoFitHeight="0" ss:Height="25.5"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="48">\n' +
                '    <Cell ss:StyleID="s49"><Data ss:Type="String">Temps passé en avant-vente</Data></Cell>\n' +
                '    <Cell ss:StyleID="s50"/>\n' +
                '    <Cell ss:StyleID="s50"/>\n' +
                '    <Cell ss:StyleID="s51"/>\n' +
                '    <Cell ss:StyleID="s51"/>\n' +
                '    <Cell ss:StyleID="s51"/>\n' +
                '    <Cell ss:StyleID="s51"/>\n' +
                '    <Cell ss:StyleID="s51"/>\n' +
                '   </Row>\n';
            text_test += '   <Row ss:AutoFitHeight="0" ss:Height="18">\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106024372"><Data ss:Type="String">QUI ?</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106024392"><Data ss:Type="String">N° SEM</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106024412"><Data ss:Type="String">CLIENT</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056592"><Data ss:Type="String">PROJET</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056572"><Data ss:Type="String">DETAIL DE LA TÂCHE </Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056552"><Data ss:Type="String">Temps passé</Data></Cell>\n' +
                '    <Cell ss:MergeAcross="1" ss:MergeDown="1" ss:StyleID="m140462106056612"><Data ss:Type="String">Commentaire </Data></Cell>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                emptyCell(33, 21) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                '    <Cell ss:Index="9" ss:StyleID="s18"/>\n' +
                emptyCell(33, 21) +
                '   </Row>\n';
            for (let i = 0; i < part_AV.length; i++) {
                let object = part_AV[i];
                text_test += '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                    '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + acronyme.trim().toUpperCase() + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s34"><Data ss:Type="String">S' + week + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + object.client + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + object.project + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s39"><Data ss:Type="String">' + removeBalise(object.detail) + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s52"><Data ss:Type="Number">' + object.duration + '</Data></Cell>\n' +
                    '    <Cell ss:MergeAcross="1" ss:StyleID="m140462106055528"><Data ss:Type="String">' + removeBalise(object.commentaire) + '</Data></Cell>\n' +
                    emptyCell(33, 21) +
                    '   </Row>\n';
            }
                text_test += '   <Row ss:AutoFitHeight="0" ss:Height="24.75">\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s42"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s43" ss:Formula="=SUM(R[-' + part_AV.length + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                '    <Cell ss:MergeAcross="1" ss:StyleID="m140462106055448"/>\n' +
                '   </Row>\n';

            text_test += '   <Row ss:AutoFitHeight="0" ss:Height="25.5"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="48">\n' +
                '    <Cell ss:StyleID="s53"><Data ss:Type="String">Temps passé sur des projets Mayflower (site internet, marque…)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s53"/>\n' +
                '    <Cell ss:StyleID="s53"/>\n' +
                '    <Cell ss:StyleID="s54"/>\n' +
                '    <Cell ss:StyleID="s54"/>\n' +
                '    <Cell ss:StyleID="s54"/>\n' +
                '    <Cell ss:StyleID="s54"/>\n' +
                '    <Cell ss:StyleID="s54"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18">\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106024372"><Data ss:Type="String">QUI ?</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106024392"><Data ss:Type="String">N° SEM</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106024412"><Data ss:Type="String">CLIENT</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056592"><Data ss:Type="String">PROJET</Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056572"><Data ss:Type="String">DETAIL DE LA TÂCHE </Data></Cell>\n' +
                '    <Cell ss:MergeDown="1" ss:StyleID="m140462106056552"><Data ss:Type="String">Temps passé</Data></Cell>\n' +
                '    <Cell ss:MergeAcross="1" ss:MergeDown="1" ss:StyleID="m140462106056612"><Data ss:Type="String">Commentaire </Data></Cell>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                emptyCell(33, 21) +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                '    <Cell ss:Index="9" ss:StyleID="s18"/>\n' +
                emptyCell(33, 21) +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                // '    <Cell ss:StyleID="s33"/>\n' +
                '   </Row>\n';
            for (let i = 0; i < part_I.length; i++) {
                let object = part_I[i];
                text_test += '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                    '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + acronyme + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s34"><Data ss:Type="String">S' + week + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s45"><Data ss:Type="String">' + object.client + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + object.project + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s39"><Data ss:Type="String">' + object.detail + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s55"><Data ss:Type="Number">' + object.duration + '</Data></Cell>\n' +
                    '    <Cell ss:MergeAcross="1" ss:StyleID="m140462106056532"><Data ss:Type="String">' + removeBalise(object.commentaire) + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s23"/>\n' +
                    emptyCell(33, 21) +
                    '   </Row>\n';
            }
            text_test += '   <Row ss:AutoFitHeight="0" ss:Height="24.75">\n' +
                emptyCell(41, 4) +
                '    <Cell ss:StyleID="s42"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s43" ss:Formula="=SUM(R[-' + part_I.length + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                '    <Cell ss:MergeAcross="1" ss:StyleID="m140462106055348"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="25.5"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="48">\n' +
                '    <Cell ss:StyleID="s56"><Data ss:Type="String">TOTAL HEURES VENDUES</Data></Cell>\n' +
                     emptyCell(56, 2) +
                '    <Cell ss:StyleID="s57" ss:Formula="=SUM(R' + (11 + part_V.length) + 'C7+R' + (16 + part_V.length + part_M.length) + 'C7)"><Data\n' +
                '      ss:Type="Number">' + (part_V_duration + part_M_duration) + '</Data></Cell>\n' +
                     emptyCell(58, 4) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="30.75">\n' +
                '    <Cell ss:StyleID="s56"><Data ss:Type="String">TOTAL SEMAINE</Data></Cell>\n' +
                '    <Cell ss:StyleID="s56"/>\n' +
                '    <Cell ss:StyleID="s56"/>\n' +
                '    <Cell ss:StyleID="s57" ss:Formula="=SUM(R' + (11 + part_V.length) + 'C7+R' +
                    (16 + part_V.length + part_M.length) + 'C7+R' + (21 + part_V.length + part_M.length + part_AV.length) +
                    'C6+R' + (26 + part_V.length + part_M.length + part_AV.length + part_I.length) + 'C6)"><Data\n' +
                '      ss:Type="Number">' + (part_V_duration + part_M_duration + part_I_duration + part_AV_duration) + '</Data></Cell>\n' +
                '    <Cell ss:StyleID="s58"/>\n' +
                '    <Cell ss:StyleID="s58"/>\n' +
                '    <Cell ss:StyleID="s58"/>\n' +
                '    <Cell ss:StyleID="s58"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75" ss:Span="941"/>\n' +
                '  </Table>\n' +
                '  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">\n' +
                '   <PageSetup>\n' +
                '    <Header x:Margin="0"/>\n' +
                '    <Footer x:Margin="0"/>\n' +
                '   </PageSetup>\n' +
                '   <Unsynced/>\n' +
                '   <Print>\n' +
                '    <ValidPrinterInfo/>\n' +
                '    <PaperSizeIndex>9</PaperSizeIndex>\n' +
                '    <HorizontalResolution>600</HorizontalResolution>\n' +
                '    <VerticalResolution>600</VerticalResolution>\n' +
                '   </Print>\n' +
                '   <Selected/>\n' +
                '   <TopRowVisible>34</TopRowVisible>\n' +
                '   <Panes>\n' +
                '    <Pane>\n' +
                '     <Number>3</Number>\n' +
                '     <ActiveRow>17</ActiveRow>\n' +
                '     <ActiveCol>5</ActiveCol>\n' +
                '    </Pane>\n' +
                '   </Panes>\n' +
                '   <ProtectObjects>False</ProtectObjects>\n' +
                '   <ProtectScenarios>False</ProtectScenarios>\n' +
                '  </WorksheetOptions>\n' +
                '  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">\n' +
                '   <Range>R29C8:R32C8,R13C8</Range>\n' +
                '   <Type>List</Type>\n' +
                '   <Value>R38C11:R39C11</Value>\n' +
                '   <InputHide/>\n' +
                '  </DataValidation>\n' +
                '  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">\n' +
                '   <Range>R39C10:R40C10,R50C10:R51C10</Range>\n' +
                '   <Type>List</Type>\n' +
                '   <Value>R39C10:R40C10</Value>\n' +
                '   <InputHide/>\n' +
                '  </DataValidation>\n' +
                '  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">\n' +
                '   <Range>R11C5:R12C5,R15C5:R21C5</Range>\n' +
                '   <Type>List</Type>\n' +
                '   <Value>\'Nomemclature activités\'!R2C1:R9C1</Value>\n' +
                '   <InputHide/>\n' +
                '  </DataValidation>\n' +
                ' </Worksheet>\n' +
                ' <Worksheet ss:Name="Nomemclature activités">\n' +
                '  <Table ss:ExpandedColumnCount="22" ss:ExpandedRowCount="1000" x:FullColumns="1"\n' +
                '   x:FullRows="1" ss:StyleID="s15" ss:DefaultColumnWidth="67"\n' +
                '   ss:DefaultRowHeight="15">\n' +
                '   <Column ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="287"/>\n' +
                '   <Column ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="412"/>\n' +
                '   <Column ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="65" ss:Span="3"/>\n' +
                '   <Column ss:Index="7" ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="63"\n' +
                '    ss:Span="15"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s16"><Data ss:Type="String">Famille d\'activités</Data></Cell>\n' +
                '    <Cell ss:StyleID="s16"><Data ss:Type="String">Exemples</Data></Cell>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Direction conseil, technique et éditoriale</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes les tâches d\'audit, benchmarmark, réalisation de recommandations, définition de plans de com, de stratégies éditoriales, de choix techniques…</Data></Cell>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Pilotage de projet</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes les réunions projets, les tâches liées au suivi, les briefs, le pilotage de l\'équipe interne ou des prestataires, les points téléphoniques client, les reportings…et les livrables associés (compte-rendus, cadrage…)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Conception</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Tous les brainstormings, zonings, le développement technique du projet, les recherches de modules techniques, le SEO, les  chemins de fer,  storyboards… et les livrables associés (specs, prés des zonings…)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Graphisme</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes la veille graphique, la conception graphique, les chartes graphiques, création de logos, retouche d\'image, mise en page, création de pictos, infographies…et les livrables associés (maquettes, moodboards…)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Mise en production</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes les tâches liées à la mise en production comme les recettes, changements de DNS, déploiements, déclaration Cnil, test d\'email, repo GIT….</Data></Cell>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="61.5">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Formation</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Tout ce qui est production de document de formation, formation en présentiel, assistance téléphonique…</Data></Cell>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '    <Cell ss:StyleID="s21"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s26"><Data ss:Type="String">Contenus et CM</Data></Cell>\n' +
                '    <Cell ss:StyleID="s28"><Data ss:Type="String">Tout ce qui est naming, travail sur du contenu (accroches, textes…), community management, insertion de contenu dans le back-office, recherche de visuel pour des articles… et les livrables associés</Data></Cell>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '    <Cell ss:StyleID="s18"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="54.75">\n' +
                '    <Cell ss:StyleID="s32"><Data ss:Type="String">Maintenance et interventions post-projet</Data></Cell>\n' +
                '    <Cell ss:StyleID="s26"><Data ss:Type="String">Tout ce qui est maintenance applicative, maintenance évolutive, gestion des urgences mais aussi tout ce qui concerne les interventions hors période de garantie</Data></Cell>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '    <Cell ss:StyleID="s32"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75" ss:Span="779"/>\n' +
                '  </Table>\n' +
                '  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">\n' +
                '   <PageSetup>\n' +
                '    <Layout x:Orientation="Landscape"/>\n' +
                '    <Header x:Margin="0"/>\n' +
                '    <Footer x:Margin="0"/>\n' +
                '   </PageSetup>\n' +
                '   <Unsynced/>\n' +
                '   <Print>\n' +
                '    <ValidPrinterInfo/>\n' +
                '    <HorizontalResolution>600</HorizontalResolution>\n' +
                '    <VerticalResolution>600</VerticalResolution>\n' +
                '   </Print>\n' +
                '   <Panes>\n' +
                '    <Pane>\n' +
                '     <Number>3</Number>\n' +
                '     <ActiveRow>5</ActiveRow>\n' +
                '    </Pane>\n' +
                '   </Panes>\n' +
                '   <ProtectObjects>False</ProtectObjects>\n' +
                '   <ProtectScenarios>False</ProtectScenarios>\n' +
                '  </WorksheetOptions>\n' +
                ' </Worksheet>\n' +
                '</Workbook>\n';

            createDownloader('DHA-' + acronyme + '-S' + week + '.xml', text_test, div_full);
            document.getElementById('div_DHA').classList.remove('d-none');


            //END TEST







            // let div_err = document.getElementById('err');
            // div_err.innerText = text_err;
            // appendPre('\n\n\n');
        } else {
            // appendPre('No upcoming events found.');
            return alert('Vous n\'avez rien pour cette semaine dans votre agenda !')
        }
    });
}

function emptyCell(style, nbr) {
    let text = '';
    for (let i = 0; i < nbr; i++) {
        text += '    <Cell ss:StyleID="s' + style + '"/>\n';
    }
    return text;
}


function empty(div) {
    for (let i = 0; i < div.children.length; i++) {
        div.removeChild(div.children[i]);
    }
}

function parse_category(category) {
    switch (category) {
        case 'V':
            return 'vendu';
        case 'AV':
            return 'd\'avant vente';
        case 'M':
            return 'de maintenance';
        case 'I':
            return 'interne';
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
    string = string.toLowerCase();
    return string.charAt(0).toUpperCase() + string.slice(1);
}

function toTitleCase(str) {
    str = str.toLowerCase();
    return str.replace(/\w\S*/g, function(txt){
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
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
    // var element = document.createElement('a');
    let element = document.getElementById('DHA');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
    // element.style.display = 'block';
    element.setAttribute('download', filename);
    // element.innerText = filename;
    element.querySelector('#full').innerHTML = filename;
    // element.style.display = 'none';
    // div.appendChild(element);

    // element.click();

    // document.body.removeChild(element);
}

function removeBalise(text) {
    // let matches = text.match(/\<.*\>(.*)\<\/.*\>/);
    // while (matches) {
    //     text = matches[1];
    //     console.log('match text', text, matches);
    //     matches = text.match(/\<.*\>(.*)\<\/.*\>/);
    text = text.replace(/\&nbsp/gm, ' ');
    text = text.replace(/<b>URGENT<\/b>/gmi, ' ');
    // }
    // console.log('text lala', text.replace(/<(?:.|\n)*?>/gm, ''));
    return text.replace(/<(?:.|\n)*?>/gm, '');;
}
// var result = getWeekNumber(new Date());
// document.write('It\'s currently week ' + result[1] + ' of ' + result[0]);