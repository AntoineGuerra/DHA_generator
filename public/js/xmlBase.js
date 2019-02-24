class XmlBase {

    static currentRow = 0;
    static currentCol = 0;
    static  maxCol = 0

    static xmlCellIndent = '';
    static xmlColIndent = '';

    /**
     * Set cell Indent by space number
     *
     * @param space {int}
     */
    static set cellIndent(space) {
        let indent = '';
        for (let i = 0; i < space; i++) {
            indent += ' ';
        }
        XmlBase.xmlCellIndent = indent;
    }

    /**
     * return space numbers to cell indent
     * @returns {string}
     */
    static get cellIndent() {
        console.log('cml cel indent', this.xmlCellIndent.length);
        return XmlBase.xmlCellIndent;
    }


    /**
     * Set col Indent by space number
     * @param space {int}
     */
    static set colIndent(space) {
        let indent = '';
        for (let i = 0; i < space; i++) {
            indent += ' ';
        }
        XmlBase.xmlColIndent = indent;
    }

    /**
     * return space numbers to col indent

     * @returns {string}
     */
    static get colIndent() {
        return XmlBase.xmlColIndent;
    }


    constructor() {

        this.dataTypes = {
            string: 'String',
            number: 'Number',
        };
    }




    /**
     *
     * @param styleID {string}
     * @param width {int}
     * @param specs {array}
     * @returns {string}
     */
    col(styleID, width, specs = []) {
        XmlBase.maxCol++;
        return this.colIndent + '<Column ' + specs.join(' ') + ' ss:StyleID="' + styleID + '" ss:AutoFitWidth="0" ss:Width="' + width + '"/>\n';

    }

    /**
     *
     * @param styleID {string}
     * @param dataType {dataTypes}
     * @param dataContent {string}
     * @param specs {array}
     * @returns {{number: number, xml: string}}
     */
    cell(styleID, dataType, dataContent = '', specs = []) {
        XmlBase.currentCol++;
        let xml = XmlBase.cellIndent + '<Cell ' + specs.join(' ') +
            ' ss:StyleID="' + styleID + '"><Data ss:Type="' + dataType + '">' + dataContent + '</Data></Cell>\n';
        let data = {xml: xml, number: XmlBase.currentCol};
        return data;
    }

    row(height, content = '', specs = []) {
        XmlBase.currentCol = 0;
        XmlBase.currentRow++;
        let xml = this.colIndent + '<Row  ' + specs.join(' ') + ' ss:AutoFitHeight="0" ss:Height="' + height + '">\n' +
            content +
            this.colIndent + '</Row>\n';
        return {xml: xml, number: XmlBase.currentRow};
    }



    acronymeHeader(content) {
        return this.cell('m140462106056232', this.dataTypes.string, content,['ss:MergeDown="1"']);
    }

    weekHeader(content) {
        return this.cell('m140462106056252', this.dataTypes.string, content,['ss:MergeDown="1"']);
    }

    clientHeader(content) {
        return this.cell('m140462106056272', this.dataTypes.string, content,['ss:MergeDown="1"']);
    }


    projectHeader(content) {
        return this.cell('m140462106056292', this.dataTypes.string, content,['ss:MergeDown="1"']);
    }

    familyHeader(content) {
        return this.cell('m140462106056312', this.dataTypes.string, content,['ss:MergeDown="1"']);
    }

    tacheHeader(content) {
        return this.cell('m140462106056212', this.dataTypes.string, content,['ss:MergeDown="1"']);
    }

    durationHeader(content) {
        return this.cell('m140462106024432', this.dataTypes.string, content,['ss:MergeDown="1"']);
    }

    commentHeader(content, merge = false) {
        if (!merge) {
            return this.cell('m140462106056192', this.dataTypes.string, content,['ss:MergeDown="1"']);
        } else {
            return this.cell('m140462106056612', this.dataTypes.string, content,['ss:MergeDown="1"', 'ss:MergeAcross="1"']);
        }
    }



    acronymeCell(acronyme) {
        return this.cell('s34', this.dataTypes.string, acronyme).xml;
    }

    weekCell(week) {
        return this.cell('s34', this.dataTypes.string, 'S' + week).xml;
    }

    clientCell(client) {
        return this.cell('s34', this.dataTypes.string, XmlBase.upperFirst(client)).xml;
    }

    projectCell(project) {
        return this.cell('s34', this.dataTypes.string, XmlBase.upperFirst(project)).xml;
    }

    familyCell(family) {
        return this.cell('s36', this.dataTypes.string, family).xml;
    }

    tacheCell(tache) {
        return this.cell('s39', this.dataTypes.string, tache).xml;
    }

    durationCell(duration) {
        return this.cell('s76', this.dataTypes.number, duration).xml;
    }

    commentCell(comment, merge = false) {
        if (!merge) {
            return this.cell('s39', this.dataTypes.string, comment).xml;
        } else {
            return this.cell('m140462106056532', this.dataTypes.string, comment, ['ss:MergeAcross="1"']).xml;
        }
    }

    emptyCell(style, nbr) {
        let text = '';
        for (let i = 0; i < nbr; i++) {
            text += XmlBase.cellIndent + '<Cell ss:StyleID="s' + style + '"/>\n';
        }
        return text;
    }

    /**
     * upper first letter
     * @param string {string}
     * @returns {string}
     */
    static upperFirst(string) {
        string = string.toLowerCase();
        return string.replace(/\w\S*/g, function (txt) {
            return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
        });
    }
}