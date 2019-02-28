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
//             console.log('new this.projects', this.projects);

        }
        for (let key in this.projects) {
            // skip loop if the property is from prototype
            if (!this.projects.hasOwnProperty(key)) continue;

            let project = this.projects[key];
            switch (project.category) {
                case DhaBuilder.categorys.vendu:
//                     console.log('add v');
                    this.parts.vendu.push(project);
                    break;
                case DhaBuilder.categorys.maintenance:
//                     console.log('add m');
                    this.parts.maintenance.push(project);
                    break;
                case DhaBuilder.categorys.avantVente:
                    this.parts.avantVente.push(project);
//                     console.log('add av');
                    break;
                case DhaBuilder.categorys.interne:
                    this.parts.interne.push(project);
//                     console.log('add i');
                    break;
            }
        }

        XmlBase.cellIndent = 4;
        XmlBase.colIndent = 3;

        this.xmlBuilder = new XmlBuilder(this.week, this.acronyme, this.parts);
        DhaBuilder.createDownloader('DHA-' + this.acronyme + '-S' + this.week + '.xml', this.xmlBuilder.content);

        this.displayErrors();

        XmlBase.currentRow = 0;
        XmlBase.currentCol = 0;

    }

    // get testParts() {
    //     return {
    //         vendu : [
    //             {
    //                 client: 'test cli',
    //                 project: 'test proj',
    //                 family: 'test fam',
    //                 tache: 'test tach',
    //                 duration: 2,
    //                 comment: 'test com',
    //             },
    //             {
    //                 client: 'test cli',
    //                 project: 'test proj',
    //                 family: 'test fam',
    //                 tache: 'test tach',
    //                 duration: 23,
    //                 comment: 'test com',
    //             },
    //
    //         ],
    //         maintenance: [
    //             {
    //                 client: 'test cli',
    //                 project: 'test proj',
    //                 family: 'test fam',
    //                 tache: 'test tach',
    //                 duration: 2,
    //                 comment: 'test com',
    //             },
    //             {
    //                 client: 'test cli',
    //                 project: 'test proj',
    //                 family: 'test fam',
    //                 tache: 'test tach',
    //                 duration: 23,
    //                 comment: 'test com',
    //             },
    //
    //         ],
    //         avantVente: [
    //             {
    //                 client: 'test cli',
    //                 project: 'test proj',
    //                 // family: 'test fam',
    //                 tache: 'test tach',
    //                 duration: 2,
    //                 comment: 'test com',
    //             },
    //             {
    //                 client: 'test cli',
    //                 project: 'test proj',
    //                 // family: 'test fam',
    //                 tache: 'test tach',
    //                 duration: 23,
    //                 comment: 'test com',
    //             },
    //
    //         ],
    //         interne: [
    //             {
    //                 client: 'test cli',
    //                 project: 'test proj',
    //                 // family: 'test fam',
    //                 tache: 'test tach',
    //                 duration: 2,
    //                 comment: 'test com',
    //             },
    //             {
    //                 client: 'test cli',
    //                 project: 'test proj',
    //                 // family: 'test fam',
    //                 tache: 'test tach',
    //                 duration: 23,
    //                 comment: 'test com',
    //             },
    //
    //         ],
    //
    //     };
    // }

    /**
     * Really it's CONSTANT
     * @returns {{avantVente: string, interne: string, maintenance: string, vendu: string}}
     */
    static get categorys() {
        return {
            vendu: 'V',
            maintenance: 'M',
            avantVente: 'AV',
            interne: 'I',
        }
    }

    static createDownloader(filename, text) {
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

    displayErrors() {
        let div_err = document.getElementById('table-error-content');

        if (this.errorProjects.length > 0) {
            document.getElementById('table-error').classList.remove('d-none');
            document.getElementById('syntax-error').classList.remove('d-none');
            document.getElementById('title-error').classList.remove('d-none');
            let text_content = '';
            for (let i = 0; i < this.errorProjects.length; i++) {
                let obj = this.errorProjects[i];
                // let declined = (obj.declined) ? 'Refusé !' : '';

                text_content +=  '<tr>' +
                // text_content +=  '<tr onclick="window.open(\'' + obj.link + '\', \'_blank\');"' + ' class="clickable">' +
                        '<th scope="row">' + i + '</th>' +
                        '<td>' + obj.name + '</td>' +
                        '<td>' + obj.date + '</td>' +
                        '<td>' + obj.duration + 'H</td>' +
                        '<td>' + obj.error + '</td>' +
                        '<td>' + obj.actions.edit + obj.actions.show + '</td>' +
                    '</tr>';


                // text_content += '<div class="col-12 text-danger ">Tâche : <a href="' + obj.link + '" target="_blank">' + obj.name + '</a> ' + declined + ' Durée : ' + obj.duration + 'H</div>';
            }
//             console.log('text content', text_content);
            div_err.innerHTML = text_content;
        } else {
            document.getElementById('div_err').classList.add('d-none');
        }
    }

}