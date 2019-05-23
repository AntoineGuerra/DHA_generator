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

    total() {
        let formulaTotal = 'R' + this.vendu.rows.total + 'C' + this.vendu.cols.duration +
            '+R' + this.maintenance.rows.total + 'C' + this.maintenance.cols.duration +
            '+R' + this.avantVente.rows.total + 'C' + this.avantVente.cols.duration +
            '+R' + this.interne.rows.total + 'C' + this.interne.cols.duration;

        let formulaVendu = 'R' + this.vendu.rows.total + 'C' + this.vendu.cols.duration + '+' +
            'R' + this.maintenance.rows.total + 'C' + this.maintenance.cols.duration;
        return this.row(48,
            this.cell('s56', this.dataTypes.string, 'TOTAL HEURES VENDUES').xml +
            this.emptyCell(56, 2) +
            this.cell('s57', this.dataTypes.number, '',
                ['ss:Formula="=SUM(' + formulaVendu + ')"']).xml
            ).xml +
            this.row(30,
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
