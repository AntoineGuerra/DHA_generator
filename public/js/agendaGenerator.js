
/** Example
 // var event = {
        //     'summary': 'Google I/O 2015',
        //     'location': '800 Howard St., San Francisco, CA 94103',
        //     'description': 'A chance to hear more about Google\'s developer products.',
        //     'start': {
        //         'dateTime': '2019-05-18T09:10:00+01:00',
        //         'timeZone': 'Europe/Paris'
        //     },
        //     'end': {
        //         'dateTime': '2019-05-18T19:12:00+01:00',
        //         'timeZone': 'Europe/Paris'
        //     },
        //     'recurrence': [
        //         'RRULE:FREQ=DAILY;COUNT=2'
        //     ],
        //     'attendees': [
        //         {'email': 'lpage@example.com'},
        //         {'email': 'sbrin@example.com'}
        //     ],
        //     'reminders': {
        //         'useDefault': false,
        //         'overrides': [
        //             {'method': 'email', 'minutes': 24 * 60},
        //             {'method': 'popup', 'minutes': 10}
        //         ]
        //     }
        // };
 */
class Event {

    /**
     * CONSTANT START_DATE
     * @returns {{current_week: string, next_week: string}}
     * @constructor
     */
    get START_DATE() {
        return {
            current_week: 'CURRENT_WEEK',
            next_week: 'NEXT_WEEK',
            selected_week: 'SELECTED_WEEK',
        }
    }

    colors = {
        maintenance: 6,
        vendus: 10,
        avantVente: 8,
        interne: 2,
    };

    summary = '';
    start = {
        dateTime: '',
        timeZone: '',
    };
    end = {
        dateTime: '',
        timeZone: '',
    };
    // colorId = ;
    colorId;
    set duration(value) {
        this.addDate(value);
    }
    constructor(name, duration) {
        this.settings = {
            start_date: document.getElementById('agenda_settings_day').value,
        };
        this.summary = name.trim();
        if (this.summary.length > 70) {
            let tache = this.summary.split('-');

            this.description = tache.pop().trim();
        }
        this.addDate(parseFloat(duration));
        this.setColor();
    }

    setColor() {
        if (this.summary.startsWith('M') || this.summary.startsWith('m')) {
            this.colorId = this.colors.maintenance;
        } else if (this.summary.startsWith('P') || this.summary.startsWith('p')) {
            this.colorId = this.colors.vendus;
        } else if (this.summary.startsWith('i') || this.summary.startsWith('I')) {
            this.colorId = this.colors.interne;
        } else if (this.summary.startsWith('av') || this.summary.startsWith('AV')) {
            this.colorId = this.colors.avantVente;
        }
    }

    addDate(duration) {
        const HOURS = 10;
        const DAY = 1;

        let date = new Date();

        this.start.timeZone = 'Europe/Paris';
        this.end.timeZone = 'Europe/Paris';
        let start_date;
        switch (this.settings.start_date) {
            case this.START_DATE.current_week:
                let time = getWeekNumber((new Date()));
                console.log('time ', time);
                let tmp = time[1];
                console.log('tmp date', tmp);
                start_date = getDateOfISOWeek(tmp, (new Date()).getFullYear());
                console.log('start date ', start_date);
                break;
            case this.START_DATE.next_week:
                start_date = Event.nextWeekdayDate(date, DAY).toISOString();
                break;
            case this.START_DATE.selected_week:
                let week = document.getElementById('week').value;
                if (parseInt(week) > 0) {
                    start_date = getDateOfISOWeek(week, (new Date()).getFullYear());
                } else {
                    let time = getWeekNumber((new Date()));
                    start_date = getDateOfISOWeek(time[1], time[0]);
                }
                break;
        }
        start_date.setHours(10);
        console.log('start date', start_date);
        this.start.dateTime = start_date.toISOString();
        console.log('this start date', start_date);
        let end = start_date;
        end.setHours(HOURS + duration);
        end.setMinutes((parseFloat(duration) % 1) * 60);
        console.log('this start date after', this.start.dateTime);
        this.end.dateTime = end.toISOString();
    }
    /**
     * params
     * date [JS Date()]
     * day_in_week [int] 1 (Mon) - 7 (Sun)
     */
    static nextWeekdayDate(date, day_in_week) {
        var ret = new Date(date|| new Date());
        ret.setDate(ret.getDate() + (day_in_week - 1 - ret.getDay() + 7) % 7 + 1);
        ret.setHours(10);
        ret.setMinutes(0);
        ret.setSeconds(0);
        return ret;
    }
}

/**
 *
 */
class AgendaGenerator {

    result_array = [];
    wait = {
        array: false,
    };
    constructor() {
        let that = this;
        this.tippys = {};
        this.week = parseInt(document.getElementById('week').value);
        this.acronyme = document.getElementById('acronyme').value;
        if (this.acronyme === "" || this.acronyme === undefined ||
            Number.isNaN(this.week)
        ) {
            return alert('l\'acronyme ne peut pas être vide !');
        }
        this.acronyme = this.acronyme.toUpperCase();
        this.events = {};
        this.agenda_api_events = [];
        
    }

    /**
     *
     * @param callback (buildTable)
     */
    getMacro(callback) {
        let that = this;
        gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: document.querySelector('#agenda_settings_sheet_id').value,
            range: 'SYNTHESE PROJETS (Magic tab)',
        }).then(function(response) {
            if (callback.name === 'parseMacro') {
                let result_macro = callback(response, that, 'array');
                if (result_macro.error !== undefined) {
                    return result_macro.error;
                }
                that.result_array = result_macro;
                that.wait.array = true;
            } else {
                // that.buildTable(response);
                callback(response, that);
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

    /**
     *
     *  @param data : {
     * infos: {
     *          range: {
     *              start: Int,
     *              end: Int,
     *          },
     *          week_column: Int,
     *      }
     * values: Array[ Array<row>, Array<row> ]
     *   }
     *   @param output : string = html, array, event
     */
    parseMacro(response, that, output = 'array') {
        const OUTPUTS = {
            html: 'html',
            array: 'array',
            event: 'event',
        };
        let result = response.result;

        let valuesInfo = that.parseValuesInfos(result.values);
        let data = {
            infos: valuesInfo,
            values: result.values,
        };

        let total_duration = 0.0;

        let count = 0;
        let btn_count = 0;
        let html = '';
        let result_arr = [];
        let event_obj = {};
        if (output === OUTPUTS.html) {
            document.getElementById('macro_file').innerHTML =
                '<button data-tippy-content="Voir le macro" class="btn btn-outline-secondary white-hover">' +
                '<a target="_blank" href="https://docs.google.com/spreadsheets/d/' + document.querySelector('#agenda_settings_sheet_id').value + '/edit' +
                '#gid=1532188122&range=' + (data.infos.range.start + 1) + ':' + (data.infos.range.end + 1) + '">' +
                '<i class="far fa-file-excel fa-3x"></i>' +
                '</a></button>';
        }
        // let i = 0;
        for (let row_number = data.infos.range.start; row_number < data.infos.range.end; row_number++) {
            let row = data.values[row_number];
            console.log('row_number', row[data.infos.week_column]);
            if (row[data.infos.week_column] === undefined) {
                return {
                    error: {
                        status: 'ERROR',
                        message: 'Semaine introuvable dans le document',
                        code: '#',
                    }
                }
            }
            let duration = parseFloat(row[data.infos.week_column].replace(',', '.'));

            let ids = {
                name: 'agenda_name_' + row_number,
                duration: 'agenda_duration_' + row_number,
                status: 'agenda_status_' + row_number,
                actions: 'agenda_actions_' + row_number,
            };

            if (duration !== undefined && !Number.isNaN(duration) && duration > 0) {
                count++;

                total_duration += duration;

                let name = {
                    project: row[1].trim(),
                    tache: row[2].trim(),
                };
                if (output === OUTPUTS.html) {
                    /**
                     * PARSE TACHES
                     */

                    name.tache = EventFilter.parseExceptions(name.tache);
                    let taches = name.tache.split('-');

                    for (let i = 0; i < taches.length; i++) {
                        // for (var key in ids) {
                        //     if (ids.hasOwnProperty(key)) {
                        //         ids[key] *= 2;
                        //     }
                        console.log('ids before', ids);
                        // }
                        let tache_ids = JSON.parse(JSON.stringify(ids));
                        Object.keys(tache_ids).map(function(key, index) {
                            tache_ids[key] += '_' + i;
                        });
                        console.log('ids after', ids);
                        let splitted_tache = {
                            name: taches[i].trim(),
                            duration: taches[i].trim().match(/^\s?([0-9]+\.?\,?[0-9]*)h\s+/i),
                        };



                        if (this.isSpecificExceptions(name.project)) {
                            console.log('parseInt(duration / 7)', parseInt(duration / 7));
                            console.log('(duration % 7)', (duration % 7));
                            splitted_tache.duration = [];
                            // splitted_tache.duration = AgendaGenerator.renderDay(duration);
                            // splitted_tache.duration[1] = (parseInt(duration / 7) + (duration % 7)).toString();
                            // console.log('name project _', name.project.substr(1, name.project.length));
                            splitted_tache.duration[1] = duration.toString();
                            /**
                             * Added Because Acronyme is removed after for all case where project name unmatch _ as first letter
                             * @type {string}
                             */
                            name.project = name.project.substr(1) + 'AGU';
                            console.log('name project _after', name.project);
                        }
                        console.log('splitted tache', splitted_tache);
                        if (splitted_tache.duration !== null) {
                            name.tache = name.tache.replace(splitted_tache.name, '');
                            splitted_tache.name = splitted_tache.name.replace(/^\s?([0-9]+\.?\,?[0-9]*)h\s+/i, '');
                            if (i < taches.length - 1 && !taches[i + 1].match(/^\s?([0-9]+\.?\,?[0-9]*)h\s+/i)) {
                                console.log('yeahhhoo', splitted_tache.name, taches[i + 1]);
                                console.log('taches', taches, taches.length);
                                console.log('tach i', i);
                                taches.splice(0, i + 1);
                                console.log('new taches', taches);
                                splitted_tache.name += ' ' + taches.join(' ');
                            }
                            let parsed_name = {
                                project: name.project,
                                tache: splitted_tache.name,
                            };
                            btn_count++;
                            console.log('duration unique tache', splitted_tache.duration[1]);
                            let splitted_duration = parseFloat(splitted_tache.duration[1].replace(',', '.'));
                            duration -= splitted_duration;
                            console.log('create table', parsed_name);
                            html += this.createTableRow(tache_ids, {
                                duration: splitted_duration,
                                name: parsed_name,
                                count: count,
                                row_number: row_number,
                                btn_count: btn_count,
                                range: data.infos.range,
                                target: row_number + '_' + i,
                            });
                            if (duration > 0) {
                                count++;
                            }
                        }
                    }


                    console.log('taches', taches);
                    // let result_name = data.name.project.substr(0, data.name.project.trim().length - 3).trim() + ' ' ;
                    console.log('duration ', duration);
                    if (duration > 0) {
                        console.log('create table 2', name);

                        html += this.createTableRow(ids, {
                            duration: duration,
                            name: name,
                            count: count,
                            row_number: row_number,
                            btn_count: btn_count,
                            range: data.infos.range,
                            // agenda_api_events_name: agenda_api_events_name,
                            // button: button,
                        });
                    }
                } else if (output === OUTPUTS.array) {
                    result_arr.push({
                        duration: duration,
                        name: name,
                        row: row_number,
                    })
                } else if (output === OUTPUTS.event) {

                }
            } else {
                // document.getElementById('table-error-AGENDA-content').innerHTML += '<tr>' +
                // '<td>' +  + '</td>' +
                // '';
            }
        }
        if (output === OUTPUTS.html) {
            html += AgendaGenerator.createTotalRow(count, data.infos, total_duration);
            return html;
        } else if (output === OUTPUTS.array) {
            return result_arr;
        } else if (output === OUTPUTS.event) {
            return event_obj;
        }

    }

    isSpecificExceptions(name) {
        return (name === '_Congés et absences' || name === '_Imprévus')
    }

    createTableRow(ids, data) {
        let button = '';
        let result_name = '';
        console.log('datttta target', data.target);
        if (data.target === undefined) {
            data.target = data.row_number;
        }
        let cells = {
            tache: '',
            duration: '',
            status: '',
            action: '',
            row: '<td data-tippy-content="Voir la ligne" class="tippy table_file_row">' +
                '<button class="btn btn-outline-secondary white-hover">' +
                    '<a target="_blank" href="https://docs.google.com/spreadsheets/d/' + document.querySelector('#agenda_settings_sheet_id').value + '/edit' +
                    '#gid=1532188122&range=' + (data.row_number + 1) + ':' + (data.row_number + 1) + '">' +
                        (data.row_number + 1) +
                    '</a>' +
                    '</button>' +
                '</td>',
        };
        if (data.name.project[0] !== '_') {
            cells.action = '<td id="' + ids.actions + '">' +
                '<button data-target="' + data.target + '" ' +
                'class="btn btn-outline-secondary white-hover agenda_generate_btn" ' +
                'data-tippy-content="Ajouter à l\'agenda" id="generate_btn_' + data.btn_count++  + '">' +
                    '<i class="far fa-calendar-plus fa-2x"></i>' +
                '</button>' + '</td>';

            result_name = data.name.project.substr(0, data.name.project.trim().length - 3).trim() + ' ' + data.name.tache.replace(/\-/g, ' ').trim();

            cells.tache = '<td><textarea id="' + ids.name + '" class="w-100 btn btn-light">' + result_name + '</textarea></td>';
            cells.duration = '<td>' + // Durée
                    '<input id="' + ids.duration + '" type="number" step="0.25" min="0" value="' + data.duration + '" class="text-center btn btn-light">' +
                '</td>';
            // cells.action =


        } else {
            result_name = data.name.project + ' ' + data.name.tache.replace(/\-/g, ' ').trim();
            cells.tache = '<td>' + result_name + '</td>';
            console.log('data duration 2', data.duration);
            cells.duration = '<td>' + data.duration + '</td>';
            cells.action = '<td></td>';
        }

        let api_datas = this.checkIfExist(result_name);
        let agenda_duration = api_datas.duration;
        let link = api_datas.link;

        console.log('result name', data.name);
        console.log('result duration', agenda_duration);
        console.log('macro duration', data.duration);
        let status_cell = {
            tippy_content: '',
            class_icon: '',
            class_text: '',
        };
        if (data.duration === agenda_duration) {
            status_cell.class_icon = 'far fa-calendar-check';
            status_cell.tippy_content = data.duration + 'H Demandées (' + agenda_duration + 'H dans l\'agenda)';
            status_cell.class_text = 'text-success';

            cells.action = '<td id="' + ids.actions + '">' + '<button class="btn btn-outline-secondary white-hover">' +
                '<a target="_blank" href="' + link + '">' +
                '<i data-tippy-content="Ouvrir dans l\'agenda" class="tippy far fa-eye fa-2x"></i>' +
                '</a>' +
                '</button>' +
                '</td>';
        } else if (data.duration < agenda_duration) {
            status_cell.class_icon = 'far fa-calendar-check';
            status_cell.tippy_content = data.duration + 'H Demandées (' + agenda_duration + 'H dans l\'agenda)';
            status_cell.class_text = 'text-warning';

            cells.action = '<td id="' + ids.actions + '">' + '<button class="btn btn-outline-secondary white-hover">' +
                '<a target="_blank" href="' + link + '">' +
                '<i data-tippy-content="Ouvrir dans l\'agenda" class="tippy far fa-eye fa-2x"></i>' +
                '</a>' +
                '</button>' +
                '</td>';
        } else if (agenda_duration === 0) {
            status_cell.class_icon = 'far fa-calendar-times';
            status_cell.tippy_content = 'N\'existe pas dans l\'agenda';
            status_cell.class_text = 'text-danger';

            // cells.status = '<td id="' + ids.status + '" class="text-danger" ' + '>' + // Status
            //     '<i data-tippy-content="N\'existe pas dans l\'agenda" class=" agenda_status far fa-calendar-times fa-2x"></i>' +
            //     '</td>';
        } else {
            status_cell.class_icon = 'far fa-calendar-times';
            status_cell.tippy_content = data.duration + 'H Demandées dans le macro (' + agenda_duration + 'H dans l\'agenda)';
            status_cell.class_text = 'text-danger';
            data.duration -= agenda_duration;

            console.log('yeahhh inferieur');
            // cells.status = '<td id="' + ids.status + '" class="text-danger" ' + '>' + // Status
            //     '<i data-tippy-content="' + '" class=" agenda_status far fa-calendar-times fa-2x"></i>' +
            //     '</td>';
        }
        cells.status = '<td id="' + ids.status + '" class="' + status_cell.class_text + '" ' + '>' + // Status
            '<i data-tippy-content="' + status_cell.tippy_content + '" class="agenda_status ' + status_cell.class_icon + ' fa-2x"></i>' +
            '</td>';
        if (data.name[0] !== '_') {
            this.events[data.target] = {
                data_ids: ids,
                event: new Event(result_name, data.duration),
            };
        }
        // if (data.duration > agenda_duration) {
        //     console.log('yeahhh inferieur');
        //     cells.status = '<td id="' + ids.status + '" class="text-danger" ' + '>' + // Status
        //         '<i data-tippy-content="' + data.duration + 'H Demandées dans le macro (' + agenda_duration + 'H dans l\'agenda)" class=" agenda_status far fa-calendar-times fa-2x"></i>' +
        //         '</td>';
        // } else {
        //     cells.status = '<td id="' + ids.status + '" class="text-danger" ' + '>' + // Status
        //         '<i data-tippy-content="N\'existe pas dans l\'agenda" class=" agenda_status far fa-calendar-times fa-2x"></i>' +
        //         '</td>';
        // }



        return  '<tr class="text-center">' +
                    '<td>' + data.count + '</td>' + // #
                    cells.row + // Ligne
                    cells.tache + // Tâche
                    cells.duration +
                    cells.status +
                    cells.action +// actions
                '</tr>';
    }

    static createTotalRow(count, valuesInfo, total_duration) {
        return  '<tr class="text-center">' +
                    '<td>' + (count + 1) + '</td>' +
            '<td data-tippy-content="Voir la ligne" class="tippy table_file_row">' +
            '<button class="btn btn-outline-secondary white-hover">' +
            '<a target="_blank" href="https://docs.google.com/spreadsheets/d/' + document.querySelector('#agenda_settings_sheet_id').value + '/edit' +
            '#gid=1532188122&range=' + (valuesInfo.range.end + 1) + ':' + (valuesInfo.range.end + 1) + '">' +
            (valuesInfo.range.end + 1) +
            '</a>' +
            '</button>' +
            '</td>' +
                    '<td>' + 'Total' + '</td>' +
                    '<td colspan="2">' + total_duration + '</td>' +
                    '<td><button id="generate_agenda_all" class="btn btn-outline-secondary white-hover" ' +
                                'data-tippy-content="Tout générer">' +
                            '<i class="far fa-calendar-plus mr-2 fa-2x"></i><i class="far fa-calendar-plus fa-2x"></i>' +
                        '</button>' +
                    '</td>' +
                '</tr>';
    }

    createArray() {
        return this.getMacro(this.parseMacro)
    }

    createTable() {
        this.getMacro(this.buildTable);

    }

    parseValuesInfos(values) {
        let that = this;

        let range = {
            start: Number,
            end: Number
        };

        let week_column = values[0].indexOf("S" + this.week);

        range.start = values.indexOf(values.find(function(element) {
            return element[0] === that.acronyme;
        }));
        range.end = values.indexOf(values.find(function(element) {
            return element[0] === that.acronyme + " Total";
        }));

        return {
            range: range,
            week_column: week_column,
        };

    }

    buildTable(response, that) {
        console.log('calll build table ', that);
        // let that = this;
        let result = response.result;
        let values = result.values;

        if (values.length > 0) {


            let table = document.getElementById('table-AGENDA-content');
            // Add HTML table content
            let result_macro = that.parseMacro(response, that, 'html');
            if (result_macro.error !== undefined) {
                alert(checkError(result_macro.error));
                emptyTables();
                removeLoader();
                return result_macro;
            }
            table.innerHTML = result_macro;
            tippy('.tippy');
            tippy('#macro_file button');
            that.tippys['generate_agenda_all'] = tippy('#generate_agenda_all');
            document.getElementById('table-AGENDA').classList.remove('d-none');
            let agenda_generate_btns = document.getElementsByClassName('agenda_generate_btn');
            tippy('.agenda_status');
            removeLoader();
            console.log('agenda_generate_btns', agenda_generate_btns);
            for (let i = 0; i < agenda_generate_btns.length; i++) {
                let btn = agenda_generate_btns[i];
                let target = btn.dataset.target;
                let current_event = that.events[target];
                console.log('events', that.events);
                console.log('current evennt', current_event);
                console.log('target', target);
                that.tippys['generate_btn_' + i] = tippy('#generate_btn_' + i);
                btn.addEventListener('click', function () {
                    that.generateAgendaEvent(current_event)
                });
                document.getElementById(current_event.data_ids.name).addEventListener('change', function () {
                    current_event.event.summary = this.value;
                });
                document.getElementById(current_event.data_ids.duration).addEventListener('change', function () {
                    current_event.event.duration = this.value;
                });
                // if ()
            }
            document.getElementById('generate_agenda_all').addEventListener('click', function () {
                that.generateAllEvent();
            })
        } else {
            alert('Il y eu une erreur la requête n\'a renvoyé aucun résultat \nResponse : ' + response.toString());
        }
    }

    generateAllEvent() {
        let that = this;
        let agenda_generate_btns = document.getElementsByClassName('agenda_generate_btn');
        for (let i = 0; i < agenda_generate_btns.length; i++) {
            let btn = agenda_generate_btns[i];
            let target = btn.dataset.target;
            let current_event = that.events[target];
            that.generateAgendaEvent(current_event)
        }
    }

    generateAgendaEvent(eventDatas) {
        let status = document.getElementById(eventDatas.data_ids.status);
        let actions = document.getElementById(eventDatas.data_ids.actions);

        status.classList.remove('text-danger');
        status.innerHTML = '<i data-tippy-content="En cours d\'envoi" class="agenda_status fas fa-spinner fa-spin fa-2x"></i>';
        tippy('.agenda_status');
        status.classList.add('text-muted');
        var request = gapi.client.calendar.events.insert({
            'calendarId': 'primary',
            'resource': eventDatas.event
        });
        request.execute(function(event) {
            console.log('event created', event);
            status.classList.remove('text-muted');
            if (event !== undefined && event !== null &&
                event.status === 'confirmed') {
                status.classList.add('text-success');
                status.innerHTML = '<i data-tippy-content="Envoyé à l\'agenda avec succès" class="agenda_status far fa-calendar-check fa-2x"></i>';
                actions.innerHTML = '<button class="btn btn-outline-secondary white-hover"><a target="_blank" href="' + event.htmlLink + '"><i data-tippy-content="Ouvrir dans l\'agenda" class="tippy far fa-eye fa-2x"></i></a></button>'
            } else {
                status.classList.add('text-danger');
                status.innerHTML = '<i data-tippy-content="Echec de l\'envoi" class="agenda_status far fa-calendar-times fa-2x"></i>';
            }
            tippy('.agenda_status');
            tippy('.tippy');

        });
    }

    checkIfExist(name) {
        console.log('agenda_events', this.agenda_api_events);

        let exist_api_agenda = this.agenda_api_events.find(function (element) {
            return element.summary === name;
        });

        let agenda_duration = 0;
        let link = '';
        while (exist_api_agenda !== undefined) {

            // remove value
            this.agenda_api_events.splice(this.agenda_api_events.indexOf(exist_api_agenda), 1);
            console.log('agenda_events after splice', this.agenda_api_events);
            console.log('agenda_events val', exist_api_agenda);
            let converted_duration = EventFilter.convert_duration(exist_api_agenda.start.dateTime, exist_api_agenda.end.dateTime);
            console.log('duration event ', converted_duration);
            agenda_duration += converted_duration;
            link = exist_api_agenda.htmlLink;
            // cells.status = '<td id="' + ids.status + '" class="text-success" ' + '>' + // Status
            //                 '<i data-tippy-content="Existe dans l\'agenda" class="agenda_status far fa-calendar-check fa-2x"></i>' +
            //                 '</td>';
            // cells.action = '<td id="' + ids.actions + '">' + '<button class="btn btn-outline-secondary white-hover">' +
            //         '<a target="_blank" href="' + exist_api_agenda.htmlLink + '">' +
            //             '<i data-tippy-content="Ouvrir dans l\'agenda" class="tippy far fa-eye fa-2x"></i>' +
            //         '</a>' +
            //     '</button>' +
            // '</td>';
            exist_api_agenda = this.agenda_api_events.find(function (element) {
                return element.summary === name;
            });
        }
        return {
            duration: agenda_duration,
            link: link,
        };
    }

    static renderDay(duration) {
        if (parseInt(duration / 7) > 1) {
            return (7).toString();
            console.log('seven H');
        } else {
            return duration.toString();
            console.log('duration H', duration);
        }
        // return undefined;
    }
}

