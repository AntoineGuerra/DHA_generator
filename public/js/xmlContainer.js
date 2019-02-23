

const xmlWorkBook = '<?xml version="1.0"?>\n' +
    '<?mso-application progid="Excel.Sheet"?>\n' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n' +
    ' xmlns:o="urn:schemas-microsoft-com:office:office"\n' +
    ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n' +
    ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"\n' +
    ' xmlns:html="http://www.w3.org/TR/REC-html40">\n' +
    ' <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">\n' +
    '  <LastAuthor>Antoine Guerra</LastAuthor>\n' +
    '  <Created>2019-01-25T15:49:35Z</Created>\n' +
    '  <LastSaved>2019-01-25T15:49:35Z</LastSaved>\n' +
    '  <Version>16.00</Version>\n' +
    ' </DocumentProperties>\n' +
    ' <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">\n' +
    '  <AllowPNG/>\n' +
    ' </OfficeDocumentSettings>\n' +
    ' <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">\n' +
    '  <WindowHeight>15540</WindowHeight>\n' +
    '  <WindowWidth>25600</WindowWidth>\n' +
    '  <WindowTopX>32767</WindowTopX>\n' +
    '  <WindowTopY>460</WindowTopY>\n' +
    '  <ProtectStructure>False</ProtectStructure>\n' +
    '  <ProtectWindows>False</ProtectWindows>\n' +
    ' </ExcelWorkbook>\n';

const endXmlWorkBook = '  </Table>\n' +
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
    '   <Range>R29C8:R32C8,R13C8</Range>\n' +
    '   <Type>List</Type>\n' +
    '   <Value>R38C11:R39C11</Value>\n' +
    '   <InputHide/>\n' +
    '  </DataValidation>\n' +
    '  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">\n' +
    '   <Range>R39C10:R40C10,R50C10:R51C10</Range>\n' +
    '   <Type>List</Type>\n' +
    '   <Value>R39C10:R40C10</Value>\n' +
    '   <InputHide/>\n' +
    '  </DataValidation>\n' +
    '  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">\n' +
    '   <Range>R11C5:R12C5,R15C5:R21C5</Range>\n' +
    '   <Type>List</Type>\n' +
    '   <Value>\'Nomemclature activités\'!R2C1:R9C1</Value>\n' +
    '   <InputHide/>\n' +
    '  </DataValidation>\n' +
    ' </Worksheet>\n' +
    // ' <Worksheet ss:Name="Nomemclature activités">\n' +
    // '  <Table ss:ExpandedColumnCount="22" ss:ExpandedRowCount="1000" x:FullColumns="1"\n' +
    // '   x:FullRows="1" ss:StyleID="s15" ss:DefaultColumnWidth="67"\n' +
    // '   ss:DefaultRowHeight="15">\n' +
    // '   <Column ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="287"/>\n' +
    // '   <Column ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="412"/>\n' +
    // '   <Column ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="65" ss:Span="3"/>\n' +
    // '   <Column ss:Index="7" ss:StyleID="s15" ss:AutoFitWidth="0" ss:Width="63"\n' +
    // '    ss:Span="15"/>\n' +
    // '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
    // '    <Cell ss:StyleID="s16"><Data ss:Type="String">Famille d\'activités</Data></Cell>\n' +
    // '    <Cell ss:StyleID="s16"><Data ss:Type="String">Exemples</Data></Cell>\n' +
    // emptyCell(18, 19) +
    // '   </Row>\n' +
    // '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Direction conseil, technique et éditoriale</Data></Cell>\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes les tâches d\'audit, benchmarmark, réalisation de recommandations, définition de plans de com, de stratégies éditoriales, de choix techniques…</Data></Cell>\n' +
    // emptyCell(21, 19) +
    // '   </Row>\n' +
    // '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Pilotage de projet</Data></Cell>\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes les réunions projets, les tâches liées au suivi, les briefs, le pilotage de l\'équipe interne ou des prestataires, les points téléphoniques client, les reportings…et les livrables associés (compte-rendus, cadrage…)</Data></Cell>\n' +
    // emptyCell(21, 19) +
    // '   </Row>\n' +
    // '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Conception</Data></Cell>\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Tous les brainstormings, zonings, le développement technique du projet, les recherches de modules techniques, le SEO, les  chemins de fer,  storyboards… et les livrables associés (specs, prés des zonings…)</Data></Cell>\n' +
    // emptyCell(21, 19) +
    // '   </Row>\n' +
    // '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Graphisme</Data></Cell>\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes la veille graphique, la conception graphique, les chartes graphiques, création de logos, retouche d\'image, mise en page, création de pictos, infographies…et les livrables associés (maquettes, moodboards…)</Data></Cell>\n' +
    // emptyCell(21, 19) +
    // '   </Row>\n' +
    // '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Mise en production</Data></Cell>\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Toutes les tâches liées à la mise en production comme les recettes, changements de DNS, déploiements, déclaration Cnil, test d\'email, repo GIT….</Data></Cell>\n' +
    // emptyCell(21, 19) +
    // '   </Row>\n' +
    // '   <Row ss:AutoFitHeight="0" ss:Height="61.5">\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Formation</Data></Cell>\n' +
    // '    <Cell ss:StyleID="s19"><Data ss:Type="String">Tout ce qui est production de document de formation, formation en présentiel, assistance téléphonique…</Data></Cell>\n' +
    // emptyCell(21, 19) +
    // '   </Row>\n' +
    // '   <Row ss:AutoFitHeight="0" ss:Height="18.75">\n' +
    // '    <Cell ss:StyleID="s26"><Data ss:Type="String">Contenus et CM</Data></Cell>\n' +
    // '    <Cell ss:StyleID="s28"><Data ss:Type="String">Tout ce qui est naming, travail sur du contenu (accroches, textes…), community management, insertion de contenu dans le back-office, recherche de visuel pour des articles… et les livrables associés</Data></Cell>\n' +
    // emptyCell(18, 19) +
    // '   </Row>\n' +
    // '   <Row ss:AutoFitHeight="0" ss:Height="54.75">\n' +
    // '    <Cell ss:StyleID="s32"><Data ss:Type="String">Maintenance et interventions post-projet</Data></Cell>\n' +
    // '    <Cell ss:StyleID="s26"><Data ss:Type="String">Tout ce qui est maintenance applicative, maintenance évolutive, gestion des urgences mais aussi tout ce qui concerne les interventions hors période de garantie</Data></Cell>\n' +
    // emptyCell(32, 19) +
    // '   </Row>\n' +
    // '   <Row ss:AutoFitHeight="0" ss:Height="15.75" ss:Span="779"/>\n' +
    // '  </Table>\n' +
    // '  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">\n' +
    // '   <PageSetup>\n' +
    // '    <Layout x:Orientation="Landscape"/>\n' +
    // '    <Header x:Margin="0"/>\n' +
    // '    <Footer x:Margin="0"/>\n' +
    // '   </PageSetup>\n' +
    // '   <Unsynced/>\n' +
    // '   <Print>\n' +
    // '    <ValidPrinterInfo/>\n' +
    // '    <HorizontalResolution>600</HorizontalResolution>\n' +
    // '    <VerticalResolution>600</VerticalResolution>\n' +
    // '   </Print>\n' +
    // '   <Panes>\n' +
    // '    <Pane>\n' +
    // '     <Number>3</Number>\n' +
    // '     <ActiveRow>5</ActiveRow>\n' +
    // '    </Pane>\n' +
    // '   </Panes>\n' +
    // '   <ProtectObjects>False</ProtectObjects>\n' +
    // '   <ProtectScenarios>False</ProtectScenarios>\n' +
    // '  </WorksheetOptions>\n' +
    // ' </Worksheet>\n' +
    '</Workbook>\n'

const xmlStyle = ' <Styles>\n' +
    '  <Style ss:ID="Default" ss:Name="Normal">\n' +
    '   <Alignment ss:Vertical="Bottom"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <Interior/>\n' +
    '   <NumberFormat/>\n' +
    '   <Protection/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056512">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056532">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056552">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056572">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056592">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056612">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056632">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056192">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056212">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056232">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056252">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056272">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056292">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106056312">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106024352">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106024372">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106024392">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106024412">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106024432">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106024452">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106024472">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106024492">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106024512">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106024532">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055328">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055348">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055368">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055388">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055408">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055428">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055448">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055468">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055488">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055508">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055528">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055548">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055568">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055588">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055608">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="m140462106055628">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s15">\n' +
    '   <Alignment ss:Vertical="Bottom"/>\n' +
    '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s16">\n' +
    '   <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000" ss:Bold="1"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s17">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="22" ss:Color="#000000" ss:Bold="1"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s18">\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s19">\n' +
    '   <Alignment ss:Vertical="Center" ss:WrapText="1"/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s20">\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="11" ss:Color="#000000"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s21">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s22">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="14" ss:Color="#000000" ss:Bold="1"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s23">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="8" ss:Color="#000000" ss:Bold="1"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s24">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#FF9300" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s25">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#000000" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#FF9300" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s26">\n' +
    '   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s27">\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <Interior ss:Color="#FF9300" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s28">\n' +
    '   <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s29">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#9BBB59" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s30">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#000000" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#9BBB59" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s31">\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <Interior ss:Color="#9BBB59" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s32">\n' +
    '   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s33">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000" ss:Bold="1"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s34">\n' +
    '   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s35">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s36">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#3F3F3F"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s37">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s38">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
    '   <Interior ss:Color="#C2D69B" ss:Pattern="Solid"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s39">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s40">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s41">\n' +
    '   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s42">\n' +
    '   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s43">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s44">\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
    '   <Interior ss:Color="#9BBB59" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s45">\n' +
    '   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s46">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s47">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s48">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#BFBFBF" ss:Pattern="Solid"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s49">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#B2A1C7" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s50">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#000000" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#B2A1C7" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s51">\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <Interior ss:Color="#B2A1C7" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s52">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
    '   <Interior ss:Color="#CCC0D9" ss:Pattern="Solid"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s53">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#0080FF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s54">\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
    '   <Interior ss:Color="#0080FF" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s55">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"\n' +
    '     ss:Color="#000000"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#FFFFFF"/>\n' +
    '   <Interior ss:Color="#8DB3E2" ss:Pattern="Solid"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s56">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="16" ss:Color="#000000" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s57">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="16" ss:Color="#000000"\n' +
    '    ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>\n' +
    '   <NumberFormat ss:Format="0.00"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s58">\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Calibri" ss:Size="12" ss:Color="#000000"/>\n' +
    '   <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s70">\n' +
    '   <Alignment ss:Vertical="Center" ss:WrapText="1"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Open Sans" ss:Size="12" ss:Color="#000000"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s71">\n' +
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Rubik Regular" ss:Size="22" ss:Color="#404040" ss:Bold="1"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s72">\n' +
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Borders/>\n' +
    '   <Font ss:FontName="Rubik Regular" ss:Size="22" ss:Color="#404040" ss:Bold="1"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s73">\n' + // Maintenance / vendu Header
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Font ss:FontName="Rubik Medium" ss:Size="16" ss:Color="#FFFFFF" ss:Bold="1"/>\n' +
    '   <Interior ss:Color="#1EAC8C" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s74">\n' + // avant vente Header
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Font ss:FontName="Rubik Medium" ss:Size="16" ss:Color="#FFFFFF"/>\n' +
    '   <Interior ss:Color="#FF9300" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '  <Style ss:ID="s75">\n' + // Interne Header
    '   <Alignment ss:Vertical="Center"/>\n' +
    '   <Font ss:FontName="Rubik Medium" ss:Size="16" ss:Color="#FFFFFF"/>\n' +
    '   <Interior ss:Color="#16CABD" ss:Pattern="Solid"/>\n' +
    '  </Style>\n' +
    '<Style ss:ID="s76">\n' + // Number
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>\n' +
    '   <Borders>\n' +
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n' +
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n' +
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n' +
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n' +
    '   </Borders>\n' +
    '   <Font ss:FontName="Rubik Regular" ss:Size="12" ss:Color="#404040"/>\n' +
    '   <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>\n' +
    '   <NumberFormat ss:Format="0.0"/>\n' +
    '   <Protection/>\n' +
    '  </Style>\n' +
    ' </Styles>\n';
function emptyCell(style, nbr) {
    let text = '';
    for (let i = 0; i < nbr; i++) {
        text += '    <Cell ss:StyleID="s' + style + '"/>\n';
    }
    return text;
}