class DhaBuilder {

    constructor(week, acronyme) {

        this.xmlBuilder = new XmlBuilder(week, acronyme, {
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

        });
        this.createDownloader('test.xml', this.xmlBuilder.content);
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