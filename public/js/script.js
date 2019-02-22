

// Array of API discovery doc URLs for APIs used by the quickstart
var DISCOVERY_DOCS = ["https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest"];

// Authorization scopes required by the API; multiple scopes can be
// included, separated by spaces.
var SCOPES = "https://www.googleapis.com/auth/calendar.readonly";
autoDl = false;
defaultFamily = 'MEP';

// var authorizeButton = document.getElementById('authorize_button');
// var signoutButton = document.getElementById('signout_button');
document.addEventListener("DOMContentLoaded", function () {
    let acroCookie = getCookie('acronyme');
    if (acroCookie !== undefined && acroCookie.length === 3) {
        document.getElementById('acronyme').value = acroCookie;
    }
    let autoDlBtn = document.getElementById('autoDl');
    let autoDlCookie = getCookie('autoDl');
    if (autoDlCookie !== undefined) {
        autoDl = JSON.parse(autoDlCookie);
        console.log('cookie', getCookie('autoDl'));

    }
    let defaultFamilyCookie = getCookie('defaultFamily');
    if (defaultFamilyCookie !== undefined) {
        defaultFamily = defaultFamilyCookie;
        document.getElementById('defaultFamily' + defaultFamilyCookie).setAttribute('selected', 'selected');
    }

    let defaultFamilySelect = document.getElementById('defaultFamily');
    defaultFamilySelect.addEventListener('change', function () {
        console.log('change', this.value);
        document.cookie = 'defaultFamily=' + this.value + '; expires=Fri, 31 Dec 2030 23:59:59 GMT';
    });
    console.log('autodl', autoDl, (autoDl == true));
    if (autoDl) {
        autoDlBtn.classList.add('btn-outline-danger');
        autoDlBtn.innerHTML = 'Désactiver Téléchargement automatique'
        // this.classList.remove('btn-success');
    } else {
        autoDlBtn.classList.add('btn-outline-success');
        autoDlBtn.innerHTML = 'Activer Téléchargement automatique'
        console.log('cookie', getCookie('autoDl'));
        // this.classList.remove('btn-danger');
    }


    autoDlBtn.addEventListener('click', function () {
        if (autoDl) {
            this.classList.add('btn-outline-success');
            this.classList.remove('btn-outline-danger');
            autoDl = false;
            document.cookie = 'autoDl=false; expires=Fri, 31 Dec 2030 23:59:59 GMT';
            this.innerHTML = 'Activer Téléchargement automatique'
            console.log('cookie', getCookie('autoDl'));
        } else {
            this.classList.add('btn-outline-danger');
            this.classList.remove('btn-outline-success');
            autoDl = true;
            document.cookie = 'autoDl=true; expires=Fri, 31 Dec 2030 23:59:59 GMT';
            this.innerHTML = 'Désactiver Téléchargement automatique'
            console.log('cookie', getCookie('autoDl'));
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
    console.log("Profile: " + profile);
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
        console.log('User signed out.');
    });
}

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
    let acronyme = document.getElementById('acronyme').value;

    generateBtn.addEventListener('click', function (event) {
        event.preventDefault();
        if (!isLogged) {
            return alert('Vous devez être connecté !')
        }
        let val = weekInput.value;
        console.log('val', val);
        let isoDate;
        if (val !== '' && parseInt(val) > 0) {
            isoDate = getDateOfISOWeek(val, (new Date()).getFullYear());
        } else {
            let time = getWeekNumber((new Date()));
            isoDate = getDateOfISOWeek(time[1], time[0])
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

    }, function (error) {
        console.log(JSON.stringify(error, null, 2));
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
        console.log('events', events);
        if (events.length > 0) {

            let acronymeInput = document.getElementById('acronyme');
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
                // console.log('exploded text', explodedText);
                // console.log('test', (((new Date(to)).getTime() - (new Date(when)).getTime() ) / 3600000));

                let duration = (((new Date(to)).getTime() - (new Date(when)).getTime()) / 3600000);

                let accepted = true;
                if (event.attendees !== undefined) {
                    let attendees = event.attendees;
                    // console.log('testt', event.attendees.length);
                    for (let j = 0; j < attendees.length; j++) {

                        if (attendees[j].self !== undefined && attendees[j].self) {

                            /** Its my response */
                            switch (attendees[j].responseStatus) {
                                case 'needsAction':
                                    // let accept = confirm('La tâche : <a href="' + event.htmlLink + '" target="_blank">' + event.summary + '</a> n\'a pas été accepté<br>Souhaitez-vous l\'enregistrer ?' );
                                    let accept = confirm('La tâche : ' + event.summary + ' (durée : ' + duration + ') n\'a pas été accepté\nSouhaitez-vous l\'enregistrer ?');
                                    if (!accept) {
                                        accepted = false;
                                        continue;
                                    }
                                    break;
                                case 'declined':
                                    accepted = false;
                                    break;
                            }
                            if (attendees[j].responseStatus === 'needsAction') {
                                projectDontProccess.push({
                                    name: event.summary,
                                    duration: duration,
                                    link: event.htmlLink,
                                });
                            }
                        }

                    }
                }

                if (!accepted) {
                    projectDontProccess.push({
                        name: event.summary,
                        duration: duration,
                        link: event.htmlLink,
                        declined: true,
                    });
                    continue;
                }
                let detail = 'SANS';
                /** Check if data are defined */
                console.log('exploded', explodedText, explodedText[0].trim().toUpperCase());
                if (explodedText[0] !== undefined && explodedText[1] !== undefined && explodedText[2] !== undefined && explodedText[3] !== undefined) {

                    category = saveCategory(explodedText[0].trim().toUpperCase());
                    client = saveClient(explodedText[1]);//@todo here change for int -> project
                    project = saveProject(explodedText[2]);
                    family = saveFamily(explodedText[3]);
                    console.log('family', family);

                } else if (explodedText[0] !== undefined && explodedText[1] !== undefined &&
                    ((explodedText[0].trim().toUpperCase() === 'AV') || (explodedText[0].trim().toUpperCase() === 'I' ||
                        explodedText[0].trim().toUpperCase() === 'INT'))) {

                    /** IF is AV || INT = family can be not defined */

                    category = explodedText[0].trim().toUpperCase();
                    family = '';
                    console.log('category', category);
                    if ((category === 'I' || category === 'INT')) {
                        console.log('INTERNE', explodedText);
                        /** IF is INT client can be not defined */
                        if (explodedText[2] === undefined || explodedText[3] === undefined) {
                            console.log('IS MAY');
                            project = explodedText[1].trim();
                            client = 'Mayflower';
                            if (explodedText[2] !== undefined) {
                                detail = removeBalise(explodedText[2]);
                            }
                        } else {
                            console.log('put in vendu');
                            category = 'V';
                            client = saveClient(explodedText[1]);
                            project = saveProject(explodedText[2]);
                            detail = saveTache(explodedText[3]);
                            // family = saveProject(explodedText[2]);
                        }

                    } else if (explodedText[2] === undefined) {

                        /** CANNOT process */

                        projectDontProccess.push({
                            name: event.summary,
                            duration: duration,
                            link: event.htmlLink,
                        });
                        continue;
                    } else {
                        client = explodedText[1].trim();
                        project = explodedText[2].trim();
                    }
                } else {

                    /** CANNOT process */

                    projectDontProccess.push({
                        name: event.summary,
                        duration: duration,
                        link: event.htmlLink,
                    });
                    continue;
                }
                let description = event.description;
                let comment = 'Sans Commentaire';

                /** FILTER description */

                if (description !== undefined) {
                    description = description.replace(new RegExp("&nbsp;", 'g'), ' ');

                    /** FILTER tache */
                    let tacheMatches = description.match(/(.*)t[a|â]ches?\s?\:?\s?<?b?r?>?\s?\<b\>([^<]*)\<\/b\>(.*)/i);
                    if (tacheMatches) {
                        console.log('tache match', tacheMatches);
                        detail = tacheMatches[2];
                    } else {
                        detail = 'SANS';
                    }

                    /** FILTER comment */
                    let commentMatches = description.match(/(.*)Commentaires?\s*:?\s*<?b?r?>?\s?\<b\>([^<]*)\<\/b\>(.*)/i);
                    // console.log('objname', objName);
                    if (category === 'M') {
                        // description = 'detail';
                        let urgentMatches = description.match(/(.*)\<b\>URGENT\<\/b\>(.*)/i);
                        if (urgentMatches) {
                            comment = 'NON';
                            // description = urgentMatches[1] + urgentMatches[3];
                            console.log('urgent match');
                        } else {
                            console.log('urgent dont match',);
                            comment = 'OUI';
                        }
                    } else if (commentMatches) {
                        console.log('comment match', commentMatches);
                        // description = commentMatches[1] + commentMatches[3];
                        comment = commentMatches[2];
                    } else {
                        comment = 'Sans Commentaire';
                    }
                } else if (category === 'M') {

                    /** IF MAINT = URGENT ? */
                    comment = 'OUI';
                }
                let parsed = parse_family(family);
                console.log('parsed family', parsed);
                if (!parsed) {

                    /** FAMILY is not valid = Probably is tache */

                    detail = family;
                    family = parse_family(defaultFamily);
                    if (description !== undefined) {

                        /** FILTER FAMILY */
                        let matchFamily = description.match(/.*famil[l|y]{1}e?\s?:?\s?<b>([^<]*)<\/b>.*/i);
                        if (matchFamily) {
                            family = parse_family(matchFamily[1]);
                            console.log('match family', family, description);
                        }
                    }
                } else {
                    family = parsed;
                }

                console.log('description', description, events);

                /** Projects GROUP BY (category, client, project, tache) */
                objName = replaceForObjectName(saveCategory(category)) + '_' + replaceForObjectName(client) + '_' +
                    replaceForObjectName(project) + '_' + replaceForObjectName(detail);

                if (family !== '' && category !== 'AV' && category !== 'I') {

                    /** IF is V || M GROUP BY family too */
                    objName += '_' + replaceForObjectName(family);

                }

                if (projects[objName] !== undefined) {

                    /** SAME Project EXIST */
                    projects[objName].duration += duration;
                    console.log('proj duration', projects[objName].duration);
                    console.log(' duration', duration);
                } else {

                    /** SAVE Project */
                    projects[objName] = {
                        category: saveCategory(category.toUpperCase()),
                        client: capitalizeFirstLetter(client),
                        project: capitalizeFirstLetter(project),
                        family: family.toUpperCase(),
                        detail: detail,
                        commentaire: comment,
                        duration: duration
                    };
                    console.log('add proj', objName);
                    console.log('duration', duration);
                }
            }

            for (var key in projects) {
                // skip loop if the property is from prototype
                if (!projects.hasOwnProperty(key)) continue;

                var obj = projects[key];
                switch (obj.category) {
                    case 'V':
                        console.log('add v');
                        part_V.push(obj);
                        part_V_duration += obj.duration;
                        break;
                    case 'M':
                        console.log('add m');
                        part_M.push(obj);
                        part_M_duration += obj.duration;
                        break;
                    case 'AV':
                        part_AV.push(obj);
                        part_AV_duration += obj.duration;
                        console.log('add av');
                        break;
                    case 'I':
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
                    let declined = (obj.declined) ? 'Refusé !' : '';
                    text_content += '<div class="col-12 text-danger ">Tâche : <a href="' + obj.link + '" target="_blank">' + obj.name + '</a> ' + declined + ' Durée : ' + obj.duration + 'H</div>';
                }
                // text_err = text_content;
                console.log('text content', text_content);
                div_err.innerHTML = '' + text_content + '';
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

            let xmlContent = xmlWorkBook + xmlStyle +
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
                '   <Row ss:AutoFitHeight="0" ss:Height="42" ss:StyleID="s72">\n' +
                '    <Cell ss:MergeAcross="7" ss:StyleID="s71"><ss:Data ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Color="#404040">DHA S</Font><Font html:Color="#16CABD">' + week + ' ' + acronyme + '</Font></B></ss:Data></Cell>\n' +
                '   </Row>\n';
            xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="25.5"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="48">\n' +
                '    <Cell ss:StyleID="s73"><Data ss:Type="String">Temps passé sur des projets (temps vendu)</Data></Cell>\n' +
                emptyCell(73, 2) +
                emptyCell(73, 5) +
                '   </Row>\n';
            xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="18">\n' +
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
                xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                    acronymeCell(acronyme.trim()) +
                    weekCell(week) +
                    clientCell(object.client) +
                    projectCell(object.project) +
                    familyCell(object.family) +
                    detailCell(removeBalise(object.detail)) +
                    durationCell(object.duration) +
                    commentCell(removeBalise(object.commentaire)) +
                    emptyCell(23, 1) +
                    emptyCell(33, 21) +
                    '   </Row>\n';
            }
            xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="27">\n' +
                emptyCell(41, 4) +
                '    <Cell ss:StyleID="s42"/>\n' +
                '    <Cell ss:StyleID="s42"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s43" ss:Formula="=SUM(R[-' + part_V.length + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                '    <Cell ss:StyleID="s43"/>\n' +
                '    <Cell ss:StyleID="s23"/>\n' +
                emptyCell(33, 21) +
                '   </Row>\n';

            /** MAINT HEADER */
            xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="24"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="45.75">\n' +
                '    <Cell ss:StyleID="s73"><Data ss:Type="String">Temps passé en maintenance (temps vendu)</Data></Cell>\n' +
                emptyCell(73, 3) +
                emptyCell(73, 4) +
                '   </Row>\n' +
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
                xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                    acronymeCell(acronyme.trim()) +
                    weekCell(week) +
                    clientCell(object.client) +
                    projectCell(object.project) +
                    familyCell(object.family) +
                    detailCell(removeBalise(object.detail)) +
                    durationCell(object.duration) +
                    '    <Cell ss:StyleID="m140462106055528"><Data ss:Type="String">' + removeBalise(object.commentaire) + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s23"/>\n' +
                    emptyCell(33, 21) +
                    '   </Row>\n';
            }
            xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="27">\n' +
                emptyCell(41, 4) +
                '    <Cell ss:StyleID="s42"/>\n' +
                '    <Cell ss:StyleID="s42"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s43" ss:Formula="=SUM(R[-' + part_M.length + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                '    <Cell ss:StyleID="s48"/>\n' +
                '    <Cell ss:StyleID="s23"/>\n' +
                emptyCell(33, 21) +
                '   </Row>\n';
            xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="25.5"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="48">\n' +
                '    <Cell ss:StyleID="s74"><Data ss:Type="String">Temps passé en avant-vente</Data></Cell>\n' +
                '    <Cell ss:StyleID="s74"/>\n' +
                '    <Cell ss:StyleID="s74"/>\n' +
                emptyCell(74, 5) +
                '   </Row>\n';
            xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="18">\n' +
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
                xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                    acronymeCell(acronyme.trim()) +
                    weekCell(week) +
                    clientCell(object.client) +
                    projectCell(object.project) +
                    detailCell(removeBalise(object.detail)) +
                    durationCell(object.duration) +
                    '    <Cell ss:MergeAcross="1" ss:StyleID="m140462106055528"><Data ss:Type="String">' + removeBalise(object.commentaire) + '</Data></Cell>\n' +
                    emptyCell(33, 21) +
                    '   </Row>\n';
            }
            xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="24.75">\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s41"/>\n' +
                '    <Cell ss:StyleID="s42"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s43" ss:Formula="=SUM(R[-' + part_AV.length + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                '    <Cell ss:MergeAcross="1" ss:StyleID="m140462106055448"/>\n' +
                '   </Row>\n';

            xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="25.5"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="48">\n' +
                '    <Cell ss:StyleID="s75"><Data ss:Type="String">Temps passé sur des projets Mayflower (site internet, marque…)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s75"/>\n' +
                '    <Cell ss:StyleID="s75"/>\n' +
                emptyCell(75, 5) +
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
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                '    <Cell ss:Index="9" ss:StyleID="s18"/>\n' +
                emptyCell(33, 21) +
                '   </Row>\n';
            for (let i = 0; i < part_I.length; i++) {
                let object = part_I[i];
                xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                    acronymeCell(acronyme.trim()) +
                    weekCell(week) +
                    clientCell(object.client) +
                    projectCell(object.project) +
                    detailCell(removeBalise(object.detail)) +
                    durationCell(object.duration) +
                    '    <Cell ss:MergeAcross="1" ss:StyleID="m140462106056532"><Data ss:Type="String">' + removeBalise(object.commentaire) + '</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s23"/>\n' +
                    emptyCell(33, 21) +
                    '   </Row>\n';
            }
            xmlContent += '   <Row ss:AutoFitHeight="0" ss:Height="24.75">\n' +
                emptyCell(41, 4) +
                '    <Cell ss:StyleID="s42"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s43" ss:Formula="=SUM(R[-' + part_I.length + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                '    <Cell ss:MergeAcross="1" ss:StyleID="m140462106055348"/>\n' +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="25.5"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="48">\n' +
                '    <Cell ss:StyleID="s56"><Data ss:Type="String">TOTAL HEURES VENDUES</Data></Cell>\n' +
                emptyCell(56, 2) +
                '    <Cell ss:StyleID="s57" ss:Formula="=SUM(R' + (8 + part_V.length) + 'C7+R' + (13 + part_V.length + part_M.length) + 'C7)"><Data\n' +
                '      ss:Type="Number">' + (part_V_duration + part_M_duration) + '</Data></Cell>\n' +
                emptyCell(58, 4) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="30.75">\n' +
                '    <Cell ss:StyleID="s56"><Data ss:Type="String">TOTAL SEMAINE</Data></Cell>\n' +
                '    <Cell ss:StyleID="s56"/>\n' +
                '    <Cell ss:StyleID="s56"/>\n' +
                '    <Cell ss:StyleID="s57" ss:Formula="=SUM(R' + (8 + part_V.length) + 'C7+R' +
                (13 + part_V.length + part_M.length) + 'C7+R' + (18 + part_V.length + part_M.length + part_AV.length) +
                'C6+R' + (23 + part_V.length + part_M.length + part_AV.length + part_I.length) + 'C6)"><Data\n' +
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
                emptyCell(18, 19) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Direction conseil, technique et éditoriale</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes les tâches d\'audit, benchmarmark, réalisation de recommandations, définition de plans de com, de stratégies éditoriales, de choix techniques…</Data></Cell>\n' +
                emptyCell(21, 19) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Pilotage de projet</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes les réunions projets, les tâches liées au suivi, les briefs, le pilotage de l\'équipe interne ou des prestataires, les points téléphoniques client, les reportings…et les livrables associés (compte-rendus, cadrage…)</Data></Cell>\n' +
                emptyCell(21, 19) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Conception</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Tous les brainstormings, zonings, le développement technique du projet, les recherches de modules techniques, le SEO, les  chemins de fer,  storyboards… et les livrables associés (specs, prés des zonings…)</Data></Cell>\n' +
                emptyCell(21, 19) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Graphisme</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes la veille graphique, la conception graphique, les chartes graphiques, création de logos, retouche d\'image, mise en page, création de pictos, infographies…et les livrables associés (maquettes, moodboards…)</Data></Cell>\n' +
                emptyCell(21, 19) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Mise en production</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes les tâches liées à la mise en production comme les recettes, changements de DNS, déploiements, déclaration Cnil, test d\'email, repo GIT….</Data></Cell>\n' +
                emptyCell(21, 19) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="61.5">\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Formation</Data></Cell>\n' +
                '    <Cell ss:StyleID="s19"><Data ss:Type="String">Tout ce qui est production de document de formation, formation en présentiel, assistance téléphonique…</Data></Cell>\n' +
                emptyCell(21, 19) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
                '    <Cell ss:StyleID="s26"><Data ss:Type="String">Contenus et CM</Data></Cell>\n' +
                '    <Cell ss:StyleID="s28"><Data ss:Type="String">Tout ce qui est naming, travail sur du contenu (accroches, textes…), community management, insertion de contenu dans le back-office, recherche de visuel pour des articles… et les livrables associés</Data></Cell>\n' +
                emptyCell(18, 19) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="54.75">\n' +
                '    <Cell ss:StyleID="s32"><Data ss:Type="String">Maintenance et interventions post-projet</Data></Cell>\n' +
                '    <Cell ss:StyleID="s26"><Data ss:Type="String">Tout ce qui est maintenance applicative, maintenance évolutive, gestion des urgences mais aussi tout ce qui concerne les interventions hors période de garantie</Data></Cell>\n' +
                emptyCell(32, 19) +
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

            createDownloader('DHA-' + acronyme + '-S' + week + '.xml', xmlContent);
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


function capitalizeFirstLetter(string) {
    string = string.toLowerCase();
    return string.charAt(0).toUpperCase() + string.slice(1);
}

function toTitleCase(str) {
    str = str.toLowerCase();
    return str.replace(/\w\S*/g, function (txt) {
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
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    // Get first day of year
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    console.log('year start', yearStart);
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

function createDownloader(filename, text) {
    // var element = document.createElement('a');
    let element = document.getElementById('DHA');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
    // element.style.display = 'block';
    element.setAttribute('download', filename);
    // element.innerText = filename;
    element.querySelector('#full').innerHTML = filename;
    if (autoDl) {
        element.click();
    }
}

function removeBalise(text, description = true) {
    if (description) {
        text = text.replace(/<b>URGENT<\/b>/gmi, ' ');
    }
    text = text.replace(/\&nbsp/gm, ' ');
    return text.replace(/<(?:.|\n)*?>/gm, '');
}

function saveCategory(data) {
    data = data.toUpperCase().trim();
    switch (data) {
        case 'PRO':
        case 'VENDU':
        case 'VENDUS':
        case 'PROJET':
        case 'PROJETS':
        case 'P':
            return 'V';

        case 'INT':
        case 'INTERNE':
            return 'I';

        case 'AVANT VENTE':
        case 'AVANT VENTES':
        case 'AVANTS VENTES':
        case 'AVANTS VENTE':
            return 'AV';

        case 'MAINT':
        case 'MAINTENANCE':
            return 'M';

        default:
            return data;
    }
}


function saveClient(client) {
    return client.trim().toUpperCase();
}
function saveProject(project) {
    return project.trim().toUpperCase();
}
function saveTache(tache) {
    //@todo complete saveClient
}
function saveFamily(family) {
    return family.trim();
}


function acronymeCell(acronyme) {
    return '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + acronyme.trim() + '</Data></Cell>\n';
}

function weekCell(week) {
    return '    <Cell ss:StyleID="s34"><Data ss:Type="String">S' + week + '</Data></Cell>\n';
}

function clientCell(client) {
    return '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + client + '</Data></Cell>\n';
}

function projectCell(project) {
    return '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + project + '</Data></Cell>\n';
}

function familyCell(family) {
    return '    <Cell ss:StyleID="s36"><Data ss:Type="String">' + parse_family(family) + '</Data></Cell>\n';
}

function detailCell(detail) {
    return '    <Cell ss:StyleID="s39"><Data ss:Type="String">' + removeBalise(detail) + '</Data></Cell>\n';
}

function durationCell(duration) {
    return '    <Cell ss:StyleID="s76"><Data ss:Type="Number">' + duration + '</Data></Cell>\n';
}

function commentCell(comment) {
    return '    <Cell ss:StyleID="s39"><Data ss:Type="String">' + removeBalise(comment) + '</Data></Cell>\n';
}



// var result = getWeekNumber(new Date());
// document.write('It\'s currently week ' + result[1] + ' of ' + result[0]);