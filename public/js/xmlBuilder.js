class XmlBuilder extends XmlBase {


    constructor(week, acronyme, parts) {
        super(week, acronyme);
        this.acronyme = acronyme;
        this.week = week;
        this.parts = parts;

        this.styleSheet = xmlStyle;
        this.workBook = xmlWorkBook;
        this.endWorkBook = endXmlWorkBook;
        this.content = this.workBook + this.styleSheet + xmlWorkBookFirst;
        this.content += this.header();
        //
        // this.vendu = new VenduXML(week, acronyme, parts);
        // this.content += this.vendu.header();
        // this.content += this.vendu.content();
        //
        // this.maintenance = new MaintenanceXML(week, acronyme, parts);
        // this.content += this.maintenance.header();
        // this.content += this.maintenance.content();
        //
        // this.avantVente = new AvantVenteXML(week, acronyme, parts);
        // this.content += this.avantVente.header();
        // this.content += this.avantVente.content();
        //
        // this.interne = new AvantVenteXML(week, acronyme, parts);
        // this.content += this.avantVente.header();
        // this.content += this.avantVente.content();

        this.vendu = new PartXML(week, acronyme, parts, 'vendu');
        this.content += this.vendu.header();
        this.content += this.vendu.content();

        this.maintenance = new PartXML(week, acronyme, parts, 'maintenance');
        this.content += this.maintenance.header();
        this.content += this.maintenance.content();

        this.avantVente = new PartXML(week, acronyme, parts, 'avantVente');
        this.content += this.avantVente.header();
        this.content += this.avantVente.content();

        this.interne = new PartXML(week, acronyme, parts, 'interne');
        this.content += this.interne.header();
        this.content += this.interne.content();



        // this.content += this.maintenance();
        // this.content += this.avantVente();
        // this.content += this.interne();
        this.content += this.total();
        this.content += this.footer();
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
                this.cell('s17', this.dataTypes.string, 'Déclaration Hebdomadaire d\'Activité (DHA)').xml +
                this.emptyCell(17, 26)
            ).xml +
            this.row(15.75,
                this.emptyCell(20, 2) +
                this.emptyCell(22, 25)
                ).xml +
            this.row(42,
                XmlBase.cellIndent + '<Cell ss:MergeAcross="7" ss:StyleID="s71"><ss:Data ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Color="#404040">DHA S</Font><Font html:Color="#16CABD">' + this.week + ' ' + this.acronyme + '</Font></B></ss:Data></Cell>\n'
                , ['ss:StyleID="s72"']
            ).xml;
    }

    // venduHeader() {
    //     let venduHeader = '';
    //     venduHeader += this.row(25.5).xml +
    //         this.row(48,
    //             this.cell('s73', this.dataTypes.string, 'Temps passé sur des projets (temps vendu)').xml +
    //             this.emptyCell(73, 7)
    //         ).xml;
    //     venduHeader += this.row(18,
    //         this.acronymeHeader('QUI ?') +
    //         this.weekHeader('N° SEM') +
    //         this.clientHeader('CLIENT') +
    //         this.projectHeader('PROJET') +
    //         this.familyHeader('FAMILLE ACTIVITÉ') +
    //         this.tacheHeader('DETAIL DE LA TÂCHE') +
    //         this.durationHeader('Heures réellement réalisées') +
    //         this.commentHeader('Commentaire sur l\'écart')
    //     ).xml;
    //     venduHeader += this.row(36.75,
    //         this.cellIndent + '<Cell ss:Index="9" ss:StyleID="s18"/>\n'
    //     ).xml;
    //     return venduHeader;
    // }
    // vendu() {
    //     let venduContent = '';
    //
    //     venduContent += this.venduHeader();
    //
    //     for (let i = 0; i < this.parts.vendu.length; i++) {
    //         let element = this.parts.vendu[i];
    //         venduContent += this.row(15.75,
    //             this.acronymeCell(this.acronyme) +
    //             this.weekCell(this.week) +
    //             this.clientCell(element.client) +
    //             this.projectCell(element.project) +
    //             this.familyCell(element.family) +
    //             this.tacheCell(element.tache) +
    //             this.durationCell(element.duration) +
    //             this.commentCell(element.comment)
    //             ).xml;
    //     }
    //     let formula = (this.parts.vendu.length > 0) ? '=SUM(R[-' + this.parts.vendu.length + ']C:R[-1]C)' : '';
    //
    //     venduContent += this.row(27,
    //         this.emptyCell(41, 4) +
    //         this.emptyCell(42, 1) +
    //         this.cell('s42', this.dataTypes.string, 'Total (heures)').xml +
    //         this.cell('s42', this.dataTypes.number, '', ['ss:Formula="' + formula + '"']).xml +
    //         this.emptyCell(43, 1)
    //         ).xml;
    //     return venduContent;
    // }

    // maintenance() {
    //     let maintenanceContent = '';
    //     maintenanceContent += this.row(25.5) +
    //         this.row(48,
    //             this.cell('s73', this.dataTypes.string, 'Temps passé en maintenance (temps vendu)').xml +
    //             this.emptyCell(73, 7)
    //         ).xml;
    //     maintenanceContent += this.row(18,
    //         this.acronymeHeader('QUI ?') +
    //         this.weekHeader('N° SEM') +
    //         this.clientHeader('CLIENT') +
    //         this.projectHeader('PROJET') +
    //         this.familyHeader('FAMILLE ACTIVITÉ') +
    //         this.tacheHeader('DETAIL DE LA TÂCHE  (facultatif)') +
    //         this.durationHeader('Temps passés en maintenance') +
    //         this.commentHeader('Tâche programmée ? (non = urgence)')
    //         // this.emptyCell(18, 1) +
    //         // this.emptyCell(33, 21)
    //         ).xml;
    //     maintenanceContent += this.row(36.75,
    //         this.cellIndent + '<Cell ss:Index="9" ss:StyleID="s18"/>\n'
    //         // this.emptyCell(33, 18) +
    //         // this.emptyCell(18, 3)
    //         ).xml;
    //
    //     for (let i = 0; i < this.parts.maintenance.length; i++) {
    //         let element = this.parts.maintenance[i];
    //         maintenanceContent += this.row(15.75,
    //             this.acronymeCell(this.acronyme) +
    //             this.weekCell(this.week) +
    //             this.clientCell(element.client) +
    //             this.projectCell(element.project) +
    //             this.familyCell(element.family) +
    //             this.tacheCell(element.tache) +
    //             this.durationCell(element.duration) +
    //             this.commentCell(element.comment)
    //             // this.emptyCell(23, 1)
    //             // this.emptyCell(33, 21)
    //             ).xml;
    //     }
    //     let formula = (this.parts.maintenance.length > 0) ? '=SUM(R[-' + this.parts.maintenance.length + ']C:R[-1]C)' : '';
    //
    //     maintenanceContent += this.row(27,
    //         this.emptyCell(41, 4) +
    //         this.emptyCell(42, 1) +
    //         this.cell('s42', this.dataTypes.string, 'Total (heures)').xml +
    //         this.cell('s42', this.dataTypes.number, '', ['ss:Formula="' + formula + '"']).xml +
    //         this.emptyCell(43, 1)
    //         // this.emptyCell(23, 1)
    //         ).xml;
    //     return maintenanceContent;
    // }


    // avantVente() {
    //     let avContent = '';
    //     avContent += this.row(25.5).xml +
    //         this.row(48,
    //             this.cell('s74', this.dataTypes.string, 'Temps passé en avant-vente').xml +
    //             this.emptyCell(74, 7)
    //         ).xml;
    //     avContent += this.row(18,
    //         this.acronymeHeader('QUI ?') +
    //         this.weekHeader('N° SEM') +
    //         this.clientHeader('CLIENT') +
    //         this.projectHeader('PROJET') +
    //         // this.familyHeader('FAMILLE ACTIVITÉ') +
    //         this.tacheHeader('DETAIL DE LA TÂCHE') +
    //         this.durationHeader('Temps passé') +
    //         this.commentHeader('Commentaire', true)
    //         // this.emptyCell(18, 1) +
    //         // this.emptyCell(33, 21)
    //         ).xml;
    //     avContent += this.row(36.75,
    //         this.cellIndent + '<Cell ss:Index="9" ss:StyleID="s18"/>\n'
    //         // this.emptyCell(33, 18) +
    //         // this.emptyCell(18, 3)
    //         ).xml;
    //
    //     for (let i = 0; i < this.parts.avantVente.length; i++) {
    //         let element = this.parts.avantVente[i];
    //         avContent += this.row(15.75,
    //             this.acronymeCell(this.acronyme) +
    //             this.weekCell(this.week) +
    //             this.clientCell(element.client) +
    //             this.projectCell(element.project) +
    //             // this.familyCell(element.family) +
    //             this.tacheCell(element.tache) +
    //             this.durationCell(element.duration) +
    //             this.commentCell(element.comment, true)
    //             // this.emptyCell(23, 1)
    //             // this.emptyCell(33, 21)
    //             ).xml;
    //     }
    //     let formula = (this.parts.avantVente.length > 0) ? '=SUM(R[-' + this.parts.avantVente.length + ']C:R[-1]C)' : '';
    //
    //     avContent += this.row(27,
    //         this.emptyCell(41, 3) +
    //         this.emptyCell(42, 1) +
    //         this.cell('s42', this.dataTypes.string, 'Total (heures)').xml +
    //         this.cell('s42', this.dataTypes.number, '', ['ss:Formula="' + formula + '"']).xml +
    //         '<Cell ss:MergeAcross="1" ss:StyleID="m140462106055348"/>\n'
    //         // this.emptyCell(43, 1)
    //         // this.emptyCell(23, 1)
    //         ).xml;
    //     return avContent;
    // }

    // interne() {
    //     let interneContent = '';
    //     interneContent += this.row(25.5).xml +
    //         this.row(48,
    //             this.cell('s75', this.dataTypes.string, 'Temps passé sur des projets Mayflower (site internet, marque…)').xml +
    //             this.emptyCell(75, 7)
    //         ).xml;
    //     interneContent += this.row(18,
    //         this.acronymeHeader('QUI ?') +
    //         this.weekHeader('N° SEM') +
    //         this.clientHeader('CLIENT') +
    //         this.projectHeader('PROJET') +
    //         // this.familyHeader('FAMILLE ACTIVITÉ') +
    //         this.tacheHeader('DETAIL DE LA TÂCHE') +
    //         this.durationHeader('Temps passé') +
    //         this.commentHeader('Commentaire', true)
    //         // this.emptyCell(18, 1) +
    //         // this.emptyCell(33, 21)
    //         ).xml;
    //     interneContent += this.row(36.75,
    //         this.cellIndent + '<Cell ss:Index="9" ss:StyleID="s18"/>\n'
    //         // this.emptyCell(33, 18) +
    //         // this.emptyCell(18, 3)
    //         ).xml;
    //
    //     for (let i = 0; i < this.parts.interne.length; i++) {
    //         let element = this.parts.interne[i];
    //         interneContent += this.row(15.75,
    //             this.acronymeCell(this.acronyme) +
    //             this.weekCell(this.week) +
    //             this.clientCell(element.client) +
    //             this.projectCell(element.project) +
    //             // this.familyCell(element.family) +
    //             this.tacheCell(element.tache) +
    //             this.durationCell(element.duration) +
    //             this.commentCell(element.comment, true)
    //             // this.emptyCell(23, 1)
    //             // this.emptyCell(33, 21)
    //             ).xml;
    //     }
    //     let formula = (this.parts.interne.length > 0) ? '=SUM(R[-' + this.parts.interne.length + ']C:R[-1]C)' : '';
    //     interneContent += this.row(27,
    //         this.emptyCell(41, 3) +
    //         this.emptyCell(42, 1) +
    //         this.cell('s42', this.dataTypes.string, 'Total (heures)').xml +
    //         this.cell('s42', this.dataTypes.number, '', ['ss:Formula="' + formula + '"']).xml +
    //         '<Cell ss:MergeAcross="1" ss:StyleID="m140462106055348"/>\n'
    //         // this.emptyCell(43, 2)
    //         // this.emptyCell(23, 1)
    //         ).xml;
    //     return interneContent;
    // }

    total() {
        // let formulaTotal = 'R' + (8 + this.parts.vendu.length) + 'C7' +
        //     '+R' + (13 + this.parts.vendu.length + this.parts.maintenance.length) + 'C7' +
        //     '+R' + (18 + this.parts.vendu.length + this.parts.maintenance.length + this.parts.avantVente.length) + 'C6' +
        //     '+R' + (23 + this.parts.vendu.length + this.parts.maintenance.length + this.parts.avantVente.length + this.parts.interne.length) + 'C6';
        let formulaTotal = 'R' + this.vendu.rows.total + 'C' + this.vendu.cols.duration +
            '+R' + this.maintenance.rows.total + 'C' + this.maintenance.cols.duration +
            '+R' + this.avantVente.rows.total + 'C' + this.avantVente.cols.duration +
            '+R' + this.interne.rows.total + 'C' + this.interne.cols.duration;
//         console.log('vendu cols', this.vendu.cols);
        let formulaVendu = 'R' + this.vendu.rows.total + 'C' + this.vendu.cols.duration + '+' +
            'R' + this.maintenance.rows.total + 'C' + this.maintenance.cols.duration;
        return this.row(48,
            // this.emptyCell(41, 4) +
            this.cell('s56', this.dataTypes.string, 'TOTAL HEURES VENDUES').xml +
            this.emptyCell(56, 2) +
            this.cell('s57', this.dataTypes.number, '',
                ['ss:Formula="=SUM(' + formulaVendu + ')"']).xml
            ).xml +
            this.row(30,
            // this.emptyCell(41, 4) +
            this.cell('s56', this.dataTypes.string, 'TOTAL SEMAINE').xml +
                this.emptyCell(56, 2) +
                this.cell('s57', this.dataTypes.number, '', ['ss:Formula="=SUM(' + formulaTotal + ')"']).xml
            ).xml;
    }

    footer() {
        let firstLineVendu = this.vendu.rows.header + 1;
        let lastLineVendu = firstLineVendu + this.parts.vendu.length;

        let firstLineMaintenance = this.maintenance.rows.header + 1;
        let lastLineMaintenance = (firstLineMaintenance + this.parts.maintenance.length);

        let listRanges = {
            vendu: 'R' + firstLineVendu + 'C' + this.vendu.cols.family + ':' + 
                'R' + lastLineVendu + 'C'+ this.maintenance.cols.family,
            maintenance: 'R' + firstLineMaintenance + 'C' + this.maintenance.cols.family + ':' +
                'R' + lastLineMaintenance + 'C' + this.maintenance.cols.family,
        };
//         console.log('list ranges', listRanges);
        

        return '  </Table>\n' +
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
            '   <Range>' + listRanges.vendu + '</Range>\n' +
            '   <Type>List</Type>\n' +
            '   <Value>\'Nomemclature activités\'!R2C1:R9C1</Value>\n' +
            '   <InputHide/>\n' +
            '  </DataValidation>\n' +
            '  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">\n' +
            '   <Range>' + listRanges.maintenance + '</Range>\n' +
            '   <Type>List</Type>\n' +
            '   <Value>\'Nomemclature activités\'!R2C1:R9C1</Value>\n' +
            '   <InputHide/>\n' +
            '  </DataValidation>\n' +
            ' </Worksheet>\n' +
            '</Workbook>\n';
    }



}