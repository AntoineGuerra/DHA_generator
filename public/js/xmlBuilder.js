class XmlBuilder extends XmlBase {


    constructor(week, acronyme, parts) {
        super(week, acronyme);
        this.acronyme = acronyme;
        this.week = week;
        this.parts = parts;

        this.styleSheet = xmlStyle;
        this.workBook = xmlWorkBook;
        this.endWorkBook = endXmlWorkBook;
        this.content = this.workBook + this.styleSheet + this.header() + this.vendu() + this.maintenance() + this.avantVente() + this.interne() + this.total() + endXmlWorkBook;
    }

    header() {
        return ' <Worksheet ss:Name="DHA">\n' +
            '  <Table ss:ExpandedColumnCount="31" ss:ExpandedRowCount="1000" x:FullColumns="1"\n' +
            '   x:FullRows="1" ss:StyleID="s15" ss:DefaultColumnWidth="67"\n' +
            '   ss:DefaultRowHeight="15">\n' +
            this.col('s15', 63, ['ss:Span="1"']) +
            this.col('s15', 189, ['ss:Index="3"', 'ss:Span="1"']) +
            this.col('s15', 258, ['ss:Index="5"']) +
            this.col('s15', 355) +
            this.col('s15', 118) +
            this.col('s15', 346) +
            this.col('s15', 346, ['ss:Span="22"']) +
            this.row(42,
                this.cell('s17', this.dataTypes.string, 'Déclaration Hebdomadaire d\'Activité (DHA)') +
                this.emptyCell(17, 26)
            ) +
            this.row(15.75,
                this.emptyCell(20, 2) +
                this.emptyCell(22, 25)
                ) +
            this.row(42,
                this.cellIndent + '<Cell ss:MergeAcross="7" ss:StyleID="s71"><ss:Data ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Color="#404040">DHA S</Font><Font html:Color="#16CABD">' + this.week + ' ' + this.acronyme + '</Font></B></ss:Data></Cell>\n'
                , ['ss:StyleID="s72"']);
    }


    vendu() {
        let venduContent = '';
        venduContent += this.row(25.5) +
            this.row(48,
                this.cell('s73', this.dataTypes.string, 'Temps passé sur des projets (temps vendu)') +
                this.emptyCell(73, 7)
            );
        venduContent += this.row(18,
            this.acronymeHeader('QUI ?') +
            this.weekHeader('N° SEM') +
            this.clientHeader('CLIENT') +
            this.projectHeader('PROJET') +
            this.familyHeader('FAMILLE ACTIVITÉ') +
            this.tacheHeader('DETAIL DE LA TÂCHE') +
            this.durationHeader('Heures réellement réalisées') +
            this.commentHeader('Commentaire sur l\'écart')
            // this.emptyCell(18, 1) +
            // this.emptyCell(33, 18)
            );
        venduContent += this.row(36.75,
            this.cellIndent + '<Cell ss:Index="9" ss:StyleID="s18"/>\n'
            // this.emptyCell(33, 18) +
            // this.emptyCell(18, 3)
            );

        for (let i = 0; i < this.parts.vendu.length; i++) {
            let element = this.parts.vendu[i];
            venduContent += this.row(15.75,
                this.acronymeCell(this.acronyme) +
                this.weekCell(this.week) +
                this.clientCell(element.client) +
                this.projectCell(element.project) +
                this.familyCell(element.family) +
                this.detailCell(element.tache) +
                this.durationCell(element.duration) +
                this.commentCell(element.comment)
                // this.emptyCell(23, 1) +
                // this.emptyCell(33, 21)
                );
        }

        venduContent += this.row(27,
            this.emptyCell(41, 4) +
            this.emptyCell(42, 1) +
            this.cell('s42', this.dataTypes.string, 'Total (heures)') +
            this.cell('s42', this.dataTypes.number, '', ['ss:Formula="=SUM(R[-' + this.parts.vendu.length + ']C:R[-1]C)"']) +
            this.emptyCell(43, 1)
            // this.emptyCell(23, 1)
            );
        return venduContent;
    }

    maintenance() {
        let maintenanceContent = '';
        maintenanceContent += this.row(25.5) +
            this.row(48,
                this.cell('s73', this.dataTypes.string, 'Temps passé en maintenance (temps vendu)') +
                this.emptyCell(73, 7)
            );
        maintenanceContent += this.row(18,
            this.acronymeHeader('QUI ?') +
            this.weekHeader('N° SEM') +
            this.clientHeader('CLIENT') +
            this.projectHeader('PROJET') +
            this.familyHeader('FAMILLE ACTIVITÉ') +
            this.tacheHeader('DETAIL DE LA TÂCHE  (facultatif)') +
            this.durationHeader('Temps passés en maintenance') +
            this.commentHeader('Tâche programmée ? (non = urgence)')
            // this.emptyCell(18, 1) +
            // this.emptyCell(33, 21)
            );
        maintenanceContent += this.row(36.75,
            this.cellIndent + '<Cell ss:Index="9" ss:StyleID="s18"/>\n'
            // this.emptyCell(33, 18) +
            // this.emptyCell(18, 3)
            );

        for (let i = 0; i < this.parts.vendu.length; i++) {
            let element = this.parts.vendu[i];
            maintenanceContent += this.row(15.75,
                this.acronymeCell(this.acronyme) +
                this.weekCell(this.week) +
                this.clientCell(element.client) +
                this.projectCell(element.project) +
                this.familyCell(element.family) +
                this.detailCell(element.tache) +
                this.durationCell(element.duration) +
                this.commentCell(element.comment)
                // this.emptyCell(23, 1)
                // this.emptyCell(33, 21)
                );
        }

        maintenanceContent += this.row(27,
            this.emptyCell(41, 4) +
            this.emptyCell(42, 1) +
            this.cell('s42', this.dataTypes.string, 'Total (heures)') +
            this.cell('s42', this.dataTypes.number, '', ['ss:Formula="=SUM(R[-' + this.parts.maintenance.length + ']C:R[-1]C)"']) +
            this.emptyCell(43, 1)
            // this.emptyCell(23, 1)
            );
        return maintenanceContent;
    }


    avantVente() {
        let avContent = '';
        avContent += this.row(25.5) +
            this.row(48,
                this.cell('s74', this.dataTypes.string, 'Temps passé en avant-vente') +
                this.emptyCell(74, 7)
            );
        avContent += this.row(18,
            this.acronymeHeader('QUI ?') +
            this.weekHeader('N° SEM') +
            this.clientHeader('CLIENT') +
            this.projectHeader('PROJET') +
            // this.familyHeader('FAMILLE ACTIVITÉ') +
            this.tacheHeader('DETAIL DE LA TÂCHE') +
            this.durationHeader('Temps passé') +
            this.commentHeader('Commentaire', true)
            // this.emptyCell(18, 1) +
            // this.emptyCell(33, 21)
            );
        avContent += this.row(36.75,
            this.cellIndent + '<Cell ss:Index="9" ss:StyleID="s18"/>\n'
            // this.emptyCell(33, 18) +
            // this.emptyCell(18, 3)
            );

        for (let i = 0; i < this.parts.vendu.length; i++) {
            let element = this.parts.vendu[i];
            avContent += this.row(15.75,
                this.acronymeCell(this.acronyme) +
                this.weekCell(this.week) +
                this.clientCell(element.client) +
                this.projectCell(element.project) +
                // this.familyCell(element.family) +
                this.detailCell(element.tache) +
                this.durationCell(element.duration) +
                this.commentCell(element.comment, true)
                // this.emptyCell(23, 1)
                // this.emptyCell(33, 21)
                );
        }

        avContent += this.row(27,
            this.emptyCell(41, 3) +
            this.emptyCell(42, 1) +
            this.cell('s42', this.dataTypes.string, 'Total (heures)') +
            this.cell('s42', this.dataTypes.number, '', ['ss:Formula="=SUM(R[-' + this.parts.avantVente.length + ']C:R[-1]C)"']) +
            '<Cell ss:MergeAcross="1" ss:StyleID="m140462106055348"/>\n'
            // this.emptyCell(43, 1)
            // this.emptyCell(23, 1)
            );
        return avContent;
    }

    interne() {
        let interneContent = '';
        interneContent += this.row(25.5) +
            this.row(48,
                this.cell('s75', this.dataTypes.string, 'Temps passé sur des projets Mayflower (site internet, marque…)') +
                this.emptyCell(75, 7)
            );
        interneContent += this.row(18,
            this.acronymeHeader('QUI ?') +
            this.weekHeader('N° SEM') +
            this.clientHeader('CLIENT') +
            this.projectHeader('PROJET') +
            // this.familyHeader('FAMILLE ACTIVITÉ') +
            this.tacheHeader('DETAIL DE LA TÂCHE') +
            this.durationHeader('Temps passé') +
            this.commentHeader('Commentaire', true)
            // this.emptyCell(18, 1) +
            // this.emptyCell(33, 21)
            );
        interneContent += this.row(36.75,
            this.cellIndent + '<Cell ss:Index="9" ss:StyleID="s18"/>\n'
            // this.emptyCell(33, 18) +
            // this.emptyCell(18, 3)
            );

        for (let i = 0; i < this.parts.vendu.length; i++) {
            let element = this.parts.vendu[i];
            interneContent += this.row(15.75,
                this.acronymeCell(this.acronyme) +
                this.weekCell(this.week) +
                this.clientCell(element.client) +
                this.projectCell(element.project) +
                // this.familyCell(element.family) +
                this.detailCell(element.tache) +
                this.durationCell(element.duration) +
                this.commentCell(element.comment, true)
                // this.emptyCell(23, 1)
                // this.emptyCell(33, 21)
                );
        }

        interneContent += this.row(27,
            this.emptyCell(41, 3) +
            this.emptyCell(42, 1) +
            this.cell('s42', this.dataTypes.string, 'Total (heures)') +
            this.cell('s42', this.dataTypes.number, '', ['ss:Formula="=SUM(R[-' + this.parts.interne.length + ']C:R[-1]C)"']) +
            '<Cell ss:MergeAcross="1" ss:StyleID="m140462106055348"/>\n'
            // this.emptyCell(43, 2)
            // this.emptyCell(23, 1)
            );
        return interneContent;
    }

    total() {
        return this.row(48,
            // this.emptyCell(41, 4) +
            this.cell('s56', this.dataTypes.string, 'TOTAL HEURES VENDUES') +
            this.emptyCell(56, 2) +
            this.cell('s57', this.dataTypes.number, '',
                ['ss:Formula="=SUM(R' + (8 + this.parts.vendu.length) + 'C7+R' + (13 + this.parts.vendu.length + this.parts.maintenance.length) + 'C7)"'])
            ) +
            this.row(30,
            // this.emptyCell(41, 4) +
            this.cell('s56', this.dataTypes.string, 'TOTAL SEMAINE') +
                this.emptyCell(56, 2) +
                this.cell('s57', this.dataTypes.number, '',
                    ['ss:Formula="=SUM(R' + (8 + this.parts.vendu.length) + 'C7' +
                    '+R' + (13 + this.parts.vendu.length + this.parts.maintenance.length) + 'C7' +
                    '+R' + (18 + this.parts.vendu.length + this.parts.maintenance.length + this.parts.avantVente.length) + 'C6' +
                    '+R' + (23 + this.parts.vendu.length + this.parts.maintenance.length + this.parts.avantVente.length + this.parts.interne.length) + 'C6)"'
                    ]
                )
            );

    }




}