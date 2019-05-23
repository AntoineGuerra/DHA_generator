class PartXML extends XmlBase {

    get title() {
        return this._title;
    }

    set title(value) {
        this.part = (this.parts[value] !== undefined) ? this.parts[value] : [];
        switch (value) {
            case 'vendu':
                this.headerStyle = 73;
                this.headerContent = 'Temps passé sur des projets (temps vendu)';
                break;
            case 'maintenance':
                this.headerStyle = 73;
                this.headerContent = 'Temps passé en maintenance (temps vendu)';
                break;
            case 'avantVente':
                this.headerStyle = 74;
                this.headerContent = 'Temps passé en avant-vente';
                this.haveFamily = false;
                break;
            case 'interne':
                this.headerStyle = 75;
                this.haveFamily = false;
                this.headerContent = 'Temps passé sur des projets Mayflower (site internet, marque…)';
                break;

        }

        this._title = value;
    }


    constructor(week, acronyme, parts, title) {
        super(week, acronyme);
        this.week = week;
        this.acronyme = acronyme
        this.parts = parts;
        this.part = [];
        this.rows = {
            header: 0,
            total: 0,
            family: 0,
        };
        this.cols = {
            duration: 0,
            family: 0,
        };
        this.haveFamily = true;
        this.headerStyle = '';
        this.headerContent = '';
        this.title = title;
    }

    header() {
        let header = '';
        header += this.row(25.5).xml +
            this.row(48,
                this.cell('s' + this.headerStyle, this.dataTypes.string, this.headerContent).xml +
                this.emptyCell(this.headerStyle, 7)
            ).xml;

        let acronymeCell =  this.acronymeHeader('QUI ?');
        let weekCell =  this.weekHeader('N° SEM');
        let clientCell =  this.clientHeader('CLIENT');
        let projectCell =  this.projectHeader('PROJET');
        let familyCell = (this.haveFamily) ? this.familyHeader('FAMILLE ACTIVITÉ') : false;
        let tacheCell =  this.tacheHeader('DETAIL DE LA TÂCHE');
        let durationCell =  this.durationHeader('Temps passé');
        let commentCell = (this.title === 'maintenance') ? this.commentHeader('Tâche programmée ? (si urgence mettre NON)', (!this.haveFamily)) : this.commentHeader('Commentaire', (!this.haveFamily));


        let headerRowContent = acronymeCell.xml +
            weekCell.xml +
            clientCell.xml +
            projectCell.xml;

        if (familyCell) {
            headerRowContent += familyCell.xml;
        }

        headerRowContent += tacheCell.xml +
            durationCell.xml +
            commentCell.xml;
        let headerRow = this.row(18,
            headerRowContent
        );
        this.cols.duration = durationCell.number;
        if (this.haveFamily && familyCell) {
            this.cols.family = familyCell.number
        }
        this.rows.header = headerRow.number;

        header += headerRow.xml;

        header += this.row(36.75,
            XmlBase.cellIndent + '<Cell ss:Index="9" ss:StyleID="s18"/>\n'
        ).xml;

        return header;
    }

    content() {
        let content = '';
        this.part.sort(function (a, b) {
            if (a.client < b.client) {
                return -1;
            }
            if (a.client > b.client) {
                return 1;
            }
            if (a.project < b.project) {
                return -1;
            }
            if (a.project > b.project) {
                return 1;
            }
            return 0;
        });
        for (let i = 0; i < this.part.length; i++) {
            let element = this.part[i];
            let rowContent = this.acronymeCell(this.acronyme) +
                this.weekCell(this.week) +
                this.clientCell(element.client) +
                this.projectCell(element.project);

            if (this.haveFamily) {
                rowContent += this.familyCell(element.family);
            }

            rowContent += this.tacheCell(element.tache) +
                this.durationCell(element.duration) +
                this.commentCell(element.comment, (!this.haveFamily));
            let row = this.row(15.75, rowContent);
            content += row.xml;
        }
        let formula = (this.part.length > 0) ? '=SUM(R[-' + this.part.length + ']C:R[-1]C)' : '';
        let totalRowContent = (this.haveFamily) ? this.emptyCell(41, 4) : this.emptyCell(41, 3);

        totalRowContent += this.emptyCell(42, 1) +
            this.cell('s42', this.dataTypes.string, 'Total (heures)').xml +
            this.cell('s42', this.dataTypes.number, '', ['ss:Formula="' + formula + '"']).xml;
        if (this.haveFamily) {
            totalRowContent += this.emptyCell(43, 1);
        } else {
            totalRowContent += '<Cell ss:MergeAcross="1" ss:StyleID="m140462106055348"/>\n';
        }
        let totalRow = this.row(27,
            totalRowContent
        );
        content += totalRow.xml;
        this.rows.total = totalRow.number;

        return content;
    }

}
