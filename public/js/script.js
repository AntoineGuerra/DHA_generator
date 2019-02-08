// Client ID and API key from the Developer Console
var CLIENT_ID = '889910005804-93lmvk75fa0un1dpd7ju8usqqp1orf2g.apps.googleusercontent.com';
var API_KEY = 'AIzaSyB46fEogWzK2xwYpV4EPN44B1OmuixoWJ4';

// Array of API discovery doc URLs for APIs used by the quickstart
var DISCOVERY_DOCS = ["https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest"];

// Authorization scopes required by the API; multiple scopes can be
// included, separated by spaces.
var SCOPES = "https://www.googleapis.com/auth/calendar.readonly";
autoDl = false;
defaultFamily = 'MEP';

// var authorizeButton = document.getElementById('authorize_button');
// var signoutButton = document.getElementById('signout_button');
document.addEventListener("DOMContentLoaded", function() {
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
    // div_full.parentElement.style.display = 'none';
    // div_V.parentElement.style.display = 'none';
    // div_AV.parentElement.style.display = 'none';
    // div_M.parentElement.style.display = 'none';
    // div_I.parentElement.style.display = 'none';
    // text_err.parentElement.style.display = 'none';
    let acronyme = document.getElementById('acronyme').value;
    // document.getElementById('acronyme').value = 'test';
    // console.log('gapi clien1', gapi);

    // checkClient = window.setInterval(function () {
    //     if (gapi !== undefined && gapi.auth2 !== undefined) {
    //         // gapi.client.calendar.events.list({
    //         //     'calendarId': 'primary',
    //         //     'showDeleted': false,
    //         //     'singleEvents': true,
    //         //     'maxResults': 1,
    //         //     // 'orderBy': 'startTime'
    //         // }).then(function(response) {
    //         //     var events = response.result.items;
    //         //     if (events[0] !== undefined && events[0].creator !== undefined && events[0].creator.email !== undefined) {
    //         //
    //         //     }
    //         // });
    //         auth2 = gapi.auth2.getAuthInstance();
    //         if (auth2.isSignedIn.get()) {
    //             var profile = auth2.currentUser.get().getBasicProfile();
    //             // console.log('ID: ' + profile.getId());
    //             // console.log('Full Name: ' + profile.getName());
    //             // console.log('Given Name: ' + profile.getGivenName());
    //             // console.log('Family Name: ' + profile.getFamilyName());
    //             // console.log('Image URL: ' + profile.getImageUrl());
    //             // console.log('Email: ' + profile.getEmail());
    //             let email = profile.getEmail();
    //             let matchs = email.match(/^([a-z]{1})[^\.]*\.([a-z]{2}).*/i);
    //             if (matchs) {
    //                  console.log('match', matchs);
    //             }
    //              console.log('ever', email);
    //             if (matchs.length >= 3) {
    //
    //                 document.getElementById('acronyme').value = (matchs[1] + matchs[2]).toUpperCase();
    //             }
    //         }
    //         clearInterval(checkClient);
    //     }
    // }, 500);
    
    generateBtn.addEventListener('click', function (event) {
        event.preventDefault();
        if (!isLogged) {
            return alert('Vous devez être connecté !')
        }
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
        // gapi.auth2.getAuthInstance().isSignedIn.listen(updateSigninStatus);

        // Handle the initial sign-in state.
        // updateSigninStatus(gapi.auth2.getAuthInstance().isSignedIn.get());
        // authorizeButton.onclick = handleAuthClick;
        // signoutButton.onclick = handleSignoutClick;
    }, function(error) {
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
    }).then(function(response) {
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
                 console.log('exploded text', explodedText);
                 console.log('test', (((new Date(to)).getTime() - (new Date(when)).getTime() ) / 3600000));
                let duration = (((new Date(to)).getTime() - (new Date(when)).getTime() ) / 3600000);
                if (explodedText[0] !== undefined && explodedText[1] !== undefined && explodedText[2] !== undefined && explodedText[3] !== undefined) {
                    category = saveData(explodedText[0].trim().toUpperCase());
                    // let client_arr = explodedText[1].split('_');
                    // if ((client_arr[0] !== undefined && client_arr[1] !== undefined)) {
                    client = explodedText[1].trim().toUpperCase();
                    project = explodedText[2].trim().toUpperCase();
                    // objName = explodedText[1].trim().toUpperCase();
                    family = explodedText[3].trim();

                    // } else {
                    //     projectDontProccess.push({
                    //         name: event.summary,
                    //         duration: duration,
                    //     });
                    //     console.log('non trier ', explodedText, client_arr);
                    //     continue;
                    // }
                } else if (explodedText[0] !== undefined && explodedText[1] !== undefined &&
                    ((explodedText[0].trim().toUpperCase() === 'AV') || (explodedText[0].trim().toUpperCase() === 'I' || explodedText[0].trim().toUpperCase() === 'INT') )) {
                    category = explodedText[0].trim().toUpperCase();
                    // let client_arr = explodedText[1].split('_');
                    // objName = explodedText[1].trim().toUpperCase();
                    // family = explodedText[2].trim().toUpperCase();
                    family = '';
                    // if ((client_arr[1] !== undefined && client_arr[1] !== undefined)) {

                    if ((category === 'I' || category === 'INT') && explodedText[2] === undefined) {
                        project = explodedText[1].trim();
                        client = 'Mayflower'
                    } else if (explodedText[2] === undefined) {
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
                    // } else {
                    //     projectDontProccess.push({
                    //         name: event.summary,
                    //         duration: duration
                    //     });
                    //     console.log('non trier 2', explodedText, client_arr);
                    //     continue;
                    // }
                } else {
                     console.log('non trier 3', explodedText);
                    projectDontProccess.push({
                        name: event.summary,
                        duration: duration,
                        link: event.htmlLink,
                    });
                    continue;
                }
                // objName = replaceForObjectName(category) + '_' + replaceForObjectName(client) + '_' + replaceForObjectName(project) + '_' + replaceForObjectName(detail);
                // if (family !== '' && category !== 'AV' && category !== 'I') {
                //      objName += '_' + replaceForObjectName(family);
                // }
                // if () else {
                //     objName = replaceForObjectName(category) + '-' + replaceForObjectName(client) + '_' + replaceForObjectName(project) + ;
                // }
                let description = event.description;
                let detail = 'SANS';
                let comment = 'Sans Commentaire';
                if (description !== undefined) {
                    description = description.replace(new RegExp("&nbsp;", 'g'), ' ');
                    let tacheMatches = description.match(/(.*)t[a|â]ches?\s?\:?\s?<?b?r?>?\s?\<b\>([^<]*)\<\/b\>(.*)/i);
                    if (tacheMatches) {
                         console.log('tache match', tacheMatches);
                        // description = tacheMatches[1] + tacheMatches[3];
                        detail = tacheMatches[2];
                    } else {
                        detail = 'SANS';
                    }
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
                             console.log('urgent dont match', );
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
                    comment = 'OUI';
                }
                let parsed = parse_family(family);
                if (!parsed) {
                    detail = family;
                    family = parse_family(defaultFamily);
                    if (description !== undefined) {
                        let matchFamily = description.match(/.*famil[l|y]{1}e?\s?:?\s?<b>([^<]*)<\/b>.*/i);
                        if (matchFamily) {
                            family = parse_family(matchFamily[1]);
                        }
                    }
                } else {
                    family = parsed;
                }

                 console.log('description', description, events);
                objName = replaceForObjectName(saveData(category)) + '_' + replaceForObjectName(client) + '_' + replaceForObjectName(project) + '_' + replaceForObjectName(detail);
                if (family !== '' && category !== 'AV' && category !== 'I') {
                    objName += '_' + replaceForObjectName(family);
                }
                if (projects[objName] !== undefined) {
                    projects[objName].duration += duration;
                     console.log('proj duration', projects[objName].duration);
                     console.log(' duration', duration);
                } else {
                    projects[objName] = {
                            category: saveData(category.toUpperCase()),
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
                    // case 'PRO':
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
                    // case 'INT':
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
                    text_content += '<div class="col-12 text-danger ">Tâche : <a href="' + obj.link + '">' + obj.name + '</a> Durée : ' + obj.duration + 'H</div>';
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

            let text_test = xmlWorkBook + xmlStyle +
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
                emptyCell(25, 2) +
                emptyCell(27, 5) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                emptyCell(20, 2) +
                emptyCell(22, 25) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                // '    <Cell ss:StyleID="s20"/>\n' +
                // '    <Cell ss:StyleID="s20"/>\n' +
                emptyCell(20, 2) +
                emptyCell(22, 5) +
                '    <Cell ss:StyleID="s23"/>\n' +
                emptyCell(22, 18) +
                '   </Row>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                // '    <Cell ss:StyleID="s20"/>\n' +
                // '    <Cell ss:StyleID="s20"/>\n' +
                emptyCell(20, 2) +
                emptyCell(22, 25) +
                '   </Row>\n';
                text_test += '   <Row ss:AutoFitHeight="0" ss:Height="25.5"/>\n' +
                    '   <Row ss:AutoFitHeight="0" ss:Height="48">\n' +
                    '    <Cell ss:StyleID="s29"><Data ss:Type="String">Temps passé sur des projets vendus (IMPORTANT POUR LES NOUVEAUX PROJETS ! Pour aider les Chefs de Projet, merci de reprendre l\'intitulé exact de la tâche tel que défini dans le CIP !)</Data></Cell>\n' +
                    emptyCell(30, 2) +
                    emptyCell(31, 5) +
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
                text_test += '   <Row ss:AutoFitHeight="0" ss:Height="27">\n' +
                    emptyCell(41, 4) +
                    '    <Cell ss:StyleID="s42"/>\n' +
                    '    <Cell ss:StyleID="s42"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                    '    <Cell ss:StyleID="s43" ss:Formula="=SUM(R[-' + part_V.length + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                    '    <Cell ss:StyleID="s43"/>\n' +
                    '    <Cell ss:StyleID="s23"/>\n' +
                    emptyCell(33, 21) +
                    '   </Row>\n';
            text_test += '   <Row ss:AutoFitHeight="0" ss:Height="24"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="45.75">\n' +
                '    <Cell ss:StyleID="s29"><Data ss:Type="String">Temps passé en maintenance (temps passé = temps déduit des contrats de maintenance)</Data></Cell>\n' +
                emptyCell(29, 3) +
                emptyCell(44, 4) +

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
                text_test += '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
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
            text_test += '   <Row ss:AutoFitHeight="0" ss:Height="27">\n' +
                    emptyCell(41, 4) +
                '    <Cell ss:StyleID="s42"/>\n' +
                '    <Cell ss:StyleID="s42"><Data ss:Type="String">Total (heures)</Data></Cell>\n' +
                '    <Cell ss:StyleID="s43" ss:Formula="=SUM(R[-' + part_M.length + ']C:R[-1]C)"><Data ss:Type="Number"></Data></Cell>\n' +
                '    <Cell ss:StyleID="s48"/>\n' +
                '    <Cell ss:StyleID="s23"/>\n' +
                emptyCell(33, 21) +
                '   </Row>\n';
            text_test += '   <Row ss:AutoFitHeight="0" ss:Height="25.5"/>\n' +
                '   <Row ss:AutoFitHeight="0" ss:Height="48">\n' +
                '    <Cell ss:StyleID="s49"><Data ss:Type="String">Temps passé en avant-vente</Data></Cell>\n' +
                '    <Cell ss:StyleID="s50"/>\n' +
                '    <Cell ss:StyleID="s50"/>\n' +
                    emptyCell(51, 5) +
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
                    // '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + acronyme.trim().toUpperCase() + '</Data></Cell>\n' +
                    // '    <Cell ss:StyleID="s34"><Data ss:Type="String">S' + week + '</Data></Cell>\n' +
                    // '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + object.client + '</Data></Cell>\n' +
                    // '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + object.project + '</Data></Cell>\n' +
                    // '    <Cell ss:StyleID="s39"><Data ss:Type="String">' + removeBalise(object.detail) + '</Data></Cell>\n' +
                    // '    <Cell ss:StyleID="s52"><Data ss:Type="Number">' + object.duration + '</Data></Cell>\n' +
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
                    emptyCell(54, 5) +
                // '    <Cell ss:StyleID="s54"/>\n' +
                // '    <Cell ss:StyleID="s54"/>\n' +
                // '    <Cell ss:StyleID="s54"/>\n' +
                // '    <Cell ss:StyleID="s54"/>\n' +
                // '    <Cell ss:StyleID="s54"/>\n' +
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
                text_test += '   <Row ss:AutoFitHeight="0" ss:Height="15.75">\n' +
                    // '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + acronyme + '</Data></Cell>\n' +
                    // '    <Cell ss:StyleID="s34"><Data ss:Type="String">S' + week + '</Data></Cell>\n' +
                    // '    <Cell ss:StyleID="s45"><Data ss:Type="String">' + object.client + '</Data></Cell>\n' +
                    // '    <Cell ss:StyleID="s34"><Data ss:Type="String">' + object.project + '</Data></Cell>\n' +
                    // '    <Cell ss:StyleID="s39"><Data ss:Type="String">' + object.detail + '</Data></Cell>\n' +
                    // '    <Cell ss:StyleID="s55"><Data ss:Type="Number">' + object.duration + '</Data></Cell>\n' +
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

            createDownloader('DHA-' + acronyme + '-S' + week + '.xml', text_test);
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

function saveData(data) {
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
    return '    <Cell ss:StyleID="s38"><Data ss:Type="Number">' + duration + '</Data></Cell>\n';
}
function commentCell(comment) {
    return '    <Cell ss:StyleID="s39"><Data ss:Type="String">' + removeBalise(comment) + '</Data></Cell>\n';
}


// var result = getWeekNumber(new Date());
// document.write('It\'s currently week ' + result[1] + ' of ' + result[0]);