class DhaBuilder {

    constructor(week, acronyme, events, defaultFamily) {
        const dha = this;
        this.projects = {};
        this.errorProjects = [];
        this.week = week;
        this.acronyme = acronyme;
        this.parts = {
            vendu: [],
            maintenance: [],
            avantVente: [],
            interne: [],
        };
        // this.parsedEvents = [];
        for (let i = 0; i < events.length; i++) {

            var event = events[i];
            new EventFilter(event, defaultFamily, dha);
            console.log('new this.projects', this.projects);

        }
        for (let key in this.projects) {
            // skip loop if the property is from prototype
            if (!this.projects.hasOwnProperty(key)) continue;

            let project = this.projects[key];
            switch (project.category) {
                case this.categorys.vendu:
                    console.log('add v');
                    this.parts.vendu.push(project);
                    break;
                case this.categorys.maintenance:
                    console.log('add m');
                    this.parts.maintenance.push(project);
                    break;
                case this.categorys.avantVente:
                    this.parts.avantVente.push(project);
                    console.log('add av');
                    break;
                case this.categorys.interne:
                    this.parts.interne.push(project);
                    console.log('add i');
                    break;
            }
        }

        this.xmlBuilder = new XmlBuilder(this.week, this.acronyme, this.parts);
        this.createDownloader('DHA-' + this.acronyme + '-S' + this.week + '.xml', this.xmlBuilder.content);
    }

    get testParts() {
        return {
            vendu : [
                {
                    client: 'test cli',
                    project: 'test proj',
                    family: 'test fam',
                    tache: 'test tach',
                    duration: 2,
                    comment: 'test com',
                },
                {
                    client: 'test cli',
                    project: 'test proj',
                    family: 'test fam',
                    tache: 'test tach',
                    duration: 23,
                    comment: 'test com',
                },

            ],
            maintenance: [
                {
                    client: 'test cli',
                    project: 'test proj',
                    family: 'test fam',
                    tache: 'test tach',
                    duration: 2,
                    comment: 'test com',
                },
                {
                    client: 'test cli',
                    project: 'test proj',
                    family: 'test fam',
                    tache: 'test tach',
                    duration: 23,
                    comment: 'test com',
                },

            ],
            avantVente: [
                {
                    client: 'test cli',
                    project: 'test proj',
                    // family: 'test fam',
                    tache: 'test tach',
                    duration: 2,
                    comment: 'test com',
                },
                {
                    client: 'test cli',
                    project: 'test proj',
                    // family: 'test fam',
                    tache: 'test tach',
                    duration: 23,
                    comment: 'test com',
                },

            ],
            interne: [
                {
                    client: 'test cli',
                    project: 'test proj',
                    // family: 'test fam',
                    tache: 'test tach',
                    duration: 2,
                    comment: 'test com',
                },
                {
                    client: 'test cli',
                    project: 'test proj',
                    // family: 'test fam',
                    tache: 'test tach',
                    duration: 23,
                    comment: 'test com',
                },

            ],

        };
    }

    /**
     * Really it's CONSTANT
     * @returns {{avantVente: string, interne: string, maintenance: string, vendu: string}}
     */
    get categorys() {
        return {
            vendu: 'V',
            maintenance: 'M',
            avantVente: 'AV',
            interne: 'I',
        }
    }

    createDownloader(filename, text) {
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


}