class XmlBase {


    constructor(week, acronyme) {
        this.xmlCellIndent = '';
        this.cellIndent = 4;
        this.xmlColIndent = '';
        this.colIndent = 3;

        this.dataTypes = {
            string: 'String',
            number: 'Number',
        };

        // this.content = this.workBook + this.styleSheet + this.dhaHeader(week, acronyme);

    }

    /**
     * Set cell Indent by space number
     *
     * @param space {int}
     */
    set cellIndent(space) {
        let indent = '';
        for (let i = 0; i < space; i++) {
            indent += ' ';
        }
        console.log('set cell indent', space, indent.length);

        this.xmlCellIndent = indent;
    }

    /**
     * return space numbers to cell indent
     * @returns {string}
     */
    get cellIndent() {
        console.log('cml cel indent', this.xmlCellIndent.length);
        return this.xmlCellIndent;
    }

    /**
     * Set col Indent by space number
     * @param space {int}
     */
    set colIndent(space) {
        let indent = '';
        for (let i = 0; i < space; i++) {
            indent += ' ';
        }
        this.xmlColIndent = indent;
    }

    /**
     * return space numbers to col indent

     * @returns {string}
     */
    get colIndent() {
        return this.xmlColIndent;
    }


    /**
     *
     * @param styleID {string}
     * @param width {int}
     * @param specs {array}
     * @returns {string}
     */
    col(styleID, width, specs = []) {
        return this.colIndent + '<Column ' + specs.join(' ') + ' ss:StyleID="' + styleID + '" ss:AutoFitWidth="0" ss:Width="' + width + '"/>\n';

    }

    /**
     *
     * @param styleID {string}
     * @param dataType {dataTypes}
     * @param dataContent {string}
     * @param specs {array}
     * @returns {string}
     */
    cell(styleID, dataType, dataContent = '', specs = []) {
        return this.cellIndent + '<Cell ' + specs.join(' ') + ' ss:StyleID="' + styleID + '"><Data ss:Type="' + dataType + '">' + dataContent + '</Data></Cell>\n';
    }

    row(height, content = '', specs = []) {
        return this.colIndent + '<Row  ' + specs.join(' ') + ' ss:AutoFitHeight="0" ss:Height="' + height + '">\n' +
            content +
            this.colIndent + '</Row>\n';
    }



    acronymeHeader(content) {
        return this.cell('m140462106056232', this.dataTypes.string, content,['ss:MergeDown="1"']);
        // return this.cellIndent + '<Cell ss:MergeDown="1" ss:StyleID="m140462106056232"><Data ss:Type="String">content</Data></Cell>\n';
    }

    weekHeader(content) {
        return this.cell('m140462106056252', this.dataTypes.string, content,['ss:MergeDown="1"']);
        // return this.cellIndent + '<Cell ss:MergeDown="1" ss:StyleID="m140462106056252"><Data ss:Type="String">content</Data></Cell>\n';
    }

    clientHeader(content) {
        return this.cell('m140462106056272', this.dataTypes.string, content,['ss:MergeDown="1"']);
        // return this.cellIndent + '<Cell ss:MergeDown="1" ss:StyleID="m140462106056252"><Data ss:Type="String">content</Data></Cell>\n';
    }


    projectHeader(content) {
        return this.cell('m140462106056292', this.dataTypes.string, content,['ss:MergeDown="1"']);
        // return this.cellIndent + '<Cell ss:MergeDown="1" ss:StyleID="m140462106056292"><Data ss:Type="String">content</Data></Cell>\n';
    }

    familyHeader(content) {
        return this.cell('m140462106056312', this.dataTypes.string, content,['ss:MergeDown="1"']);
        // return this.cellIndent + '<Cell ss:MergeDown="1" ss:StyleID="m140462106056312"><Data ss:Type="String">content</Data></Cell>\n';
    }

    tacheHeader(content) {
        return this.cell('m140462106056212', this.dataTypes.string, content,['ss:MergeDown="1"']);
        // return this.cellIndent + '<Cell ss:MergeDown="1" ss:StyleID="m140462106056212"><Data ss:Type="String">content</Data></Cell>\n';
    }

    durationHeader(content) {
        return this.cell('m140462106024432', this.dataTypes.string, content,['ss:MergeDown="1"']);
        // return this.cellIndent + '<Cell ss:MergeDown="1" ss:StyleID="m140462106024432"><Data ss:Type="String">content</Data></Cell>\n';
    }

    commentHeader(content, merge = false) {
        if (!merge) {
            return this.cell('m140462106056192', this.dataTypes.string, content,['ss:MergeDown="1"']);
        } else {
            return this.cell('m140462106056612', this.dataTypes.string, content,['ss:MergeDown="1"', 'ss:MergeAcross="1"']);
        }
        // return this.cellIndent + '<Cell ss:MergeDown="1" ss:StyleID="m140462106056192"><Data ss:Type="String">content</Data></Cell>\n';
    }



    acronymeCell(acronyme) {
        return this.cell('s34', this.dataTypes.string, acronyme);
        // return this.cellIndent + '<Cell ss:StyleID="s34"><Data ss:Type="String">' + acronyme.trim() + '</Data></Cell>\n';
    }

    weekCell(week) {
        return this.cell('s34', this.dataTypes.string, 'S' + week);
        // return this.cellIndent + '<Cell ss:StyleID="s34"><Data ss:Type="String">S' + week + '</Data></Cell>\n';
    }

    clientCell(client) {
        return this.cell('s34', this.dataTypes.string, client);
        // return this.cellIndent + '<Cell ss:StyleID="s34"><Data ss:Type="String">' + client + '</Data></Cell>\n';
    }

    projectCell(project) {
        return this.cell('s34', this.dataTypes.string, project);
        // return this.cellIndent + '<Cell ss:StyleID="s34"><Data ss:Type="String">' + project + '</Data></Cell>\n';
    }

    familyCell(family) {
        return this.cell('s36', this.dataTypes.string, family);
        // return this.cellIndent + '<Cell ss:StyleID="s36"><Data ss:Type="String">' + parse_family(family) + '</Data></Cell>\n';
    }

    detailCell(detail) {
        return this.cell('s39', this.dataTypes.string, detail);
        // return this.cellIndent + '<Cell ss:StyleID="s39"><Data ss:Type="String">' + removeBalise(detail) + '</Data></Cell>\n';
    }

    durationCell(duration) {
        return this.cell('s76', this.dataTypes.number, duration);
        // return this.cellIndent + '<Cell ss:StyleID="s76"><Data ss:Type="Number">' + duration + '</Data></Cell>\n';
    }

    commentCell(comment, merge = false) {
        if (!merge) {
            return this.cell('s39', this.dataTypes.string, comment);
        } else {
            return this.cell('m140462106056532', this.dataTypes.string, comment, ['ss:MergeAcross="1"']);
        }
        // return this.cellIndent + '<Cell ss:StyleID="s39"><Data ss:Type="String">' + removeBalise(comment) + '</Data></Cell>\n';
    }

    emptyCell(style, nbr) {
        let text = '';
        for (let i = 0; i < nbr; i++) {
            text += this.cellIndent + '<Cell ss:StyleID="s' + style + '"/>\n';
        }
        return text;
    }
}