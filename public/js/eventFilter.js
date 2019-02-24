class EventFilter {
    get error() {
        return this._error;
    }

    set error(value) {
        if (this._error.length > 0) {
            this._error += ', ';
        }
        this._error = value;
    }



    constructor(event, defaultFamily, dhaBuilder) {
        console.log('call event filter', event);
        this.event = event;
        this.dha = dhaBuilder;
        console.log('dha', this.dha);
        this.defaultFamily = defaultFamily;
        this.startTime = EventFilter.getStartTime(this.event);
        this.endTime = EventFilter.getEndTime(this.event);

        this._error = '';

        this.allProjects = this.dha.projects;
        this.errorProjects = this.dha.errorProjects;

        this._category = '';
        this._client = '';
        this._project = '';
        this._family = '';
        this._tache = '';
        this.duration = '';
        this._comment = '';
        this.objName = '';

        this.valid = true;

        this.declined = false;

        this.getDuration();
        this.checkEventDeclined(this.event);

        let datas = this.event.summary.split('-');

        this.category = datas[0];
        this.client = datas[1];
        this.project = datas[2];
        this.tache = datas[3];
        this.family = this.defaultFamily;

        /** Check IF it's interne */
        this.checkInterne(datas);

        this.filterDescription(this.event.description);

        if (!this.saveFinalProject() || !this.valid) {

            let startDate = new Date(this.startTime);
            let endDate = new Date(this.endTime);

            let date = startDate.toDateString() + ' de ' + startDate.getHours() + 'H';
            date += (startDate.getMinutes() > 0) ? startDate.getMinutes() : '00';

            date += ' à ' + endDate.getHours() + 'H';
            date += (endDate.getMinutes() > 0) ? endDate.getMinutes() : '00';


            let err = {
                name: event.summary,
                duration: this.duration,
                link: event.htmlLink,
                declined: this.declined,
                error: this.error,
                date: date,
            };
            this.errorProjects.push(err);
        }



    }

    get tache() {
        return this._tache;
    }

    set tache(tache) {
        this._tache = EventFilter.saveTache(tache);
    }

    get category() {
        return this._category;
    }

    set category(value) {
        this._category = this.saveCategory(value);
    }

    get comment() {
        return this._comment;
    }

    set comment(value) {
        this._comment = EventFilter.saveComment(value);
    }

    get client() {
        return this._client;
    }

    set client(value) {
        this._client = EventFilter.saveClient(value);
    }

    get project() {
        return this._project;
    }

    set project(value) {
        this._project = EventFilter.saveProject(value);
    }

    get family() {
        return this._family;
    }

    set family(value) {
        this._family = EventFilter.saveFamily(value);
    }

    /**
     * All accepted case for category
     * @param category
     * @returns {string|boolean}
     *      IF (category's match)
     *          parsed category
     *      ELSE
     *          false
     */
    saveCategory(category) {
        category = category.toUpperCase().trim();
        switch (category) {
            case 'V':
            case 'PRO':
            case 'VENDU':
            case 'VENDUS':
            case 'PROJET':
            case 'PROJETS':
            case 'P':
                return DhaBuilder.categorys.vendu;

            case 'I':
            case 'INT':
            case 'INTERNE':
                return DhaBuilder.categorys.interne;

            case 'AV':
            case 'AVANT VENTE':
            case 'AVANT VENTES':
            case 'AVANTS VENTES':
            case 'AVANTS VENTE':
                return DhaBuilder.categorys.avantVente;

            case 'M':
            case 'MAINT':
            case 'MAINTENANCE':
                return DhaBuilder.categorys.maintenance;

            default:
                // return category;
                return false;
                // break;
        }
    }

    /**
     * Parse project to convenient as XML && JS Object
     * @param client {string}
     * @returns {string|boolean}
     */
    static saveClient(client) {
        return (client !== undefined) ? EventFilter.stripTags(client.trim().toUpperCase()) : false;
    }

    /**
     * Parse project to convenient as XML && JS Object
     * @param project {string}
     * @returns {string|boolean}
     */
    static saveProject(project) {
        return (project !== undefined) ? EventFilter.stripTags(project.trim().toUpperCase()) : false;
    }

    /**
     * Parse tache to convenient as XML && JS Object
     * @param tache {string}
     * @returns {string|boolean}
     */
    static saveTache(tache) {
        console.log('save tache', (tache !== undefined) ? EventFilter.stripTags(tache.trim()) : 'SANS');
        return (tache !== undefined) ? EventFilter.stripTags(tache.trim()) : 'SANS';
    }

    /**
     * Parse comment to convenient as XML && JS Object
     * @param comment {string}
     * @returns {string|boolean}
     */
    static saveComment(comment) {
        return (comment !== undefined) ? EventFilter.stripTags(comment.trim()) : false;
    }




    /**
     * Parse family to convenient as XML && JS Object
     * @param family {string}
     * @returns {string|boolean}
     */
    static saveFamily(family) {
        // return (family !== undefined) ? this.stripTags(family.trim()) : false;
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
        } else if (family.match(/interventions|Maintenance/i)) {
            return 'Maintenance et interventions post-projet';
        } else {
            return false;
        }
    }

    /**
     * Parse Object name to convenient as JS Object
     * @param name
     * @returns {string}
     */
    static saveObjName(name) {
        return (name !== undefined) ? name.replace(/[^\w]/gm, 'x') : false;
    }

    /**
     * Check if event's accepted | declined | needsAction
     * IF not -> Set valid = FALSE
     * @param event {object}
     */
    checkEventDeclined(event) {
        if (event.attendees !== undefined) {
            let attendees = event.attendees;
            // console.log('testt', event.attendees.length);
            for (let j = 0; j < attendees.length; j++) {

                if (attendees[j].self !== undefined && attendees[j].self) {

                    /** Its my response */
                    switch (attendees[j].responseStatus) {
                        case 'needsAction':
                            /** Ask User IF save OR not */
                            let accept = confirm('La tâche : ' + event.summary + ' (durée : ' + duration + ') n\'a pas été accepté\nSouhaitez-vous l\'enregistrer ?');
                            if (!accept) {
                                this.valid = false;
                                this.declined = true;
                                this.error = 'Annulé par l\'utilisateur';
                                continue;
                            }
                            break;
                        case 'declined':
                            this.valid = false;
                            this.declined = true;
                            this.error = 'Refusé par l\'utilisateur dans l\'agenda';
                            break;
                    }
                }
            }
        }
    }

    /**
     * Get the start timestamp of event
     * @param event {object}
     * @returns {string}
     */
    static getStartTime(event) {
        var when = event.start.dateTime;
        if (!when) {
            when = event.start.date;
        }
        return when;
    }

    /**
     * Get the End timestamp of event
     * @param event {object}
     * @returns {string}
     */
    static getEndTime(event) {
        var to = event.end.dateTime;
        if (!to) {
            to = event.end.date
        }
        return to;
    }

    /**
     * Set duration of current event
     */
    getDuration() {
        this.duration = (((new Date(this.endTime)).getTime() - (new Date(this.startTime)).getTime()) / 3600000)
    }

    checkInterne(datas) {
        if (this.category === DhaBuilder.categorys.interne) {
            console.log('it\'s interne', 'category', this.category, 'client', this.client, 'project', this.project);
            /** in Really : Client is undefined */
            if (!this.project) {
                this.project = this.client;
                this.client = 'Mayflower';
                this.tache = datas[2];
            } else if (this.client !== 'MAYFLOWER') {
                console.log('put in vendu case');
                this.category = DhaBuilder.categorys.vendu;
            }
        }
    }

    filterDescription(description) {
        if (description !== undefined) {
            description = description.replace(new RegExp("&nbsp;", 'g'), ' ');

            /** FILTER tache */
            let tacheMatches = description.match(/(.*)t[a|â]ches?\s?:?\s?<?b?r?>?\s?<b>([^<]*)<\/b>(.*)/i);
            if (tacheMatches) {
                console.log('tache match', tacheMatches);
                this.tache = tacheMatches[2];
            }

            /** FILTER comment */
            let commentMatches = description.match(/(.*)Commentaires?\s*:?\s*<?b?r?>?\s?<b>([^<]*)<\/b>(.*)/i);
            if (this.category === DhaBuilder.categorys.maintenance) {
                let urgentMatches = description.match(/(.*)<b>URGENT<\/b>(.*)/i);
                if (urgentMatches) {
                    this.comment = 'NON';
                } else {
                    this.comment = 'OUI';
                }
            } else if (commentMatches) {
                this.comment = commentMatches[2];
            } else {
                this.comment = 'Sans Commentaire';
            }

            /** FILTER family */

                /** FILTER FAMILY */
            let matchFamily = description.match(/.*famil[l|y]{1}e?\s?:?\s?<b>([^<]*)<\/b>.*/i);
            if (matchFamily) {
                this.family = matchFamily[1];
                console.log('match family', this.family, description);
            }
        } else if (this.category === DhaBuilder.categorys.maintenance) {

            /** IF MAINT === URGENT ? */
            this.comment = 'OUI';
        } else {
            this.comment = 'SANS';
        }
    }

    saveFinalProject() {
        /** Projects GROUP BY (category, client, project, tache) */
        if (this.category && this.client && this.project && this.valid) {
            let objName = this.category + '_' + this.client + '_' +
                this.project + '_' + this.tache;

            if (this.family !== false && (this.category === DhaBuilder.categorys.vendu || this.category === DhaBuilder.categorys.maintenance)) {

                /** IF is V || M GROUP BY family too */
                objName += '_' + replaceForObjectName(this.family);

            }

            objName = EventFilter.saveObjName(objName);

            if (this.allProjects[objName] !== undefined) {

                /** SAME Project EXIST */
                this.allProjects[objName].duration += this.duration;
                console.log('proj duration', this.allProjects[objName].duration);
                console.log(' duration', this.duration);
            } else {

                /** SAVE Project */
                this.allProjects[objName] = {
                    category: this.category,
                    client: this.client,
                    project: this.project,
                    family: this.family,
                    tache: this.tache,
                    comment: this.comment,
                    duration: this.duration,
                };
                console.log('add proj', objName);
                console.log('duration', this.duration);
            }
            return true;
        } else {
            if (!this.category) {
                this.error += 'Catégorie Invalide : <pre>' + this.category + '</pre>';
            }
            if (!this.client) {
                this.error += 'Client Invalide : <pre>' + this.client + '</pre>';
            }
            if (!this.project) {
                this.error += 'Projet Invalide : <pre>' + this.project + '</pre>';
            }
            return false;
        }
    }


    /**
     * Remove TAGS to XML
     * @param data {string}
     * @returns {void | string}
     */
    static stripTags(data) {
        return data.replace(/(<([^>]+)>)/ig,"");
    }

}