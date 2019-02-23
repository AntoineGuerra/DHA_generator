class EventFilter {

    constructor(event) {
        console.log('call event filter', event);
        this.event = event;
        this.startTime = this.getStartTime(event);
        this.endTime = this.getEndTime(event);

        this.parts = {
            vendu: [],
            maintenance: [],
            avantVente: [],
            interne: [],
        };


        this.category = '';
        this.client = '';
        this.project = '';
        this.family = '';
        this.tache = '';
        this.duration = '';
        this.objName = '';

        this.valid = true;


        this.getDuration();
        this.checkEventDeclined(event);

        let datas = event.summary.split('-');

        this.category = this.saveCategory(datas[0]);
        this.client = this.saveClient(datas[1]);
        this.project = this.saveProject(datas[2]);
        this.tache = this.saveTache(datas[3]);

        if (this.category === this.categorys.avantVente || this.category === this.categorys.interne) {
            /** family can be undefined */
            if (this.category === this.categorys.interne) {
                console.log('it\'s interne', 'category', this.category, 'client', this.client, 'project', this.project);
                /** Really Client is undefined */
                if (!this.project) {
                    this.project = this.client;
                    this.client = 'Mayflower';
                }
            }
        }


        // this.client = this.saveClient(datas[1]);
        // this.project = this.saveProject(datas[2]);
        // this.tache = this.saveTache(datas[3]);

    }

    get categorys() {
        return {
            vendu: 'V',
            maintenance: 'M',
            avantVente: 'AV',
            interne: 'I',
        }
    }

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
                return this.categorys.vendu;

            case 'I':
            case 'INT':
            case 'INTERNE':
                return this.categorys.interne;

            case 'AV':
            case 'AVANT VENTE':
            case 'AVANT VENTES':
            case 'AVANTS VENTES':
            case 'AVANTS VENTE':
                return this.categorys.avantVente;

            case 'M':
            case 'MAINT':
            case 'MAINTENANCE':
                return this.categorys.maintenance;

            default:
                // return category;
                return false;
                // break;
        }
    }

    saveClient(client) {
        return (client !== undefined) ? this.stripTags(client.trim().toUpperCase()) : false;
    }

    saveProject(project) {
        return (project !== undefined) ? this.stripTags(project.trim().toUpperCase()) : false;
    }

    saveTache(tache) {
        return (tache !== undefined) ? this.stripTags(tache.trim()) : false;
    }

    saveFamily(family) {
        return (family !== undefined) ? this.stripTags(family.trim()) : false;
    }

    checkEventDeclined(event) {
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
                                this.valid = false;
                                continue;
                            }
                            break;
                        case 'declined':
                            this.valid = false;
                            break;
                    }
                }
            }
        }
    }

    getStartTime(event) {
        var when = event.start.dateTime;
        if (!when) {
            when = event.start.date;
        }
        return when;
    }
    getEndTime(event) {
        var to = event.end.dateTime;
        if (!to) {
            to = event.end.date
        }
        return to;
    }

    getDuration() {
        this.duration = (((new Date(this.endTime)).getTime() - (new Date(this.startTime)).getTime()) / 3600000)
    }


    stripTags(data) {
        return data.replace(/(<([^>]+)>)/ig,"");
    }

}