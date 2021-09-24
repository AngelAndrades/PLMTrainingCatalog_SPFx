import * as $ from 'jquery';
import * as JSZip from 'jszip';
import '@progress/kendo-ui';
import { ds, dsExpand } from  './datasource';

export class ModelState extends kendo.data.ObservableObject {
    constructor() {
        super();
    }

    public tabName: string;
    public redirectUrl: string;
}

export interface Params {
    catalogGuid: string;
    rolesGuid: string;
    maturityGuid: string;
    safeLink: string;
    wamLink: string;
}

export class SPA {
    protected static catalogGridOptions: kendo.ui.GridOptions;
    protected static catalogGrid: kendo.ui.Grid;
    protected static devSecOpsGridOptions: kendo.ui.GridOptions;
    protected static devSecOpsGrid: kendo.ui.Grid;
    protected static readingGridOptions: kendo.ui.GridOptions;
    protected static readingGrid: kendo.ui.Grid;
    protected static organizationDropDownListOptions: kendo.ui.DropDownListOptions;
    protected static organizationDropDownList: kendo.ui.DropDownList;
    protected static teamDropDownListOptions: kendo.ui.DropDownListOptions;
    protected static teamDropDownList: kendo.ui.DropDownList;
    protected static roleDropDownListOptions: kendo.ui.DropDownListOptions;
    protected static roleDropDownList: kendo.ui.DropDownList;
    protected static maturityDropDownListOptions: kendo.ui.DropDownListOptions;
    protected static maturityDropDownList: kendo.ui.DropDownList;
    protected static dsoMaturityDropDownList: kendo.ui.DropDownList;
    protected static tabStrip: kendo.ui.TabStrip;
    protected static tabStripOptions: kendo.ui.TabStripOptions;
    protected static dialog: kendo.ui.Dialog;
    protected static dialogOptions: kendo.ui.DialogOptions;
    private static instance: SPA;

    constructor() {}

    public static getInstance(args: Params): SPA {
        var appState = new ModelState();

        const dsCatalog = dsExpand({
            guid: args.catalogGuid,
            dsName: 'dsCatalog',
            schema: {
                id: 'Id',
                fields: {
                    Id: { type: 'number' },
                    Title: { type: 'string' },
                    CourseSeries: { type: 'string' },
                    LearningHours: { type: 'number' },
                    Asset_x0020_Type: { type: 'string' },
                    DSO_x0020_Maturity_x0020_Level: { type: 'string' },
                    PLM_x0020_Roadmap_x0020_Focus: { type: 'string' },
                    Link_x0020_to_x0020_Resource: { type: 'string' },
                    Role_x0028_s_x0029_: { type: 'string' },
                    Keywords: { type: 'string' },
                    TMSItemID: { type: 'string' },
                    DevSecOps: { type: 'boolean' },
                    Recommended_x0020_Reading: { type: 'boolean' }
                }
            },
            pageSize: 5,
            sort: { field: 'Title', dir: 'asc' },
            expand: ['Role_x0028_s_x0029_','Asset_x0020_Type','PLM_x0020_Roadmap_x0020_Focus'],
            expandedFields: ['Role_x0028_s_x0029_/Subrole','Asset_x0020_Type/Title','PLM_x0020_Roadmap_x0020_Focus/Title'],
            //filter: { field: 'DevSecOps', operator: 'equals', value: 0 }
        });

        const dsMaturityDataSource = ds({
            guid: args.maturityGuid,
            dsName: 'dsMaturityDataSource',
            schema: {
                id: 'Id',
                fields: {
                    Id: { type: 'number' },
                    Title: { type: 'string' }
                }
            },
            sort: {field: 'Id', dir: 'asc'}
        });

$(() => {
            this.tabStripOptions = {
                tabPosition: 'top',
                animation: { open: { effects: 'fadeIn' } },
                navigatable: true,
                select: e => {
                    switch (e.item.textContent) {
                        case 'By Role':
                            dsCatalog.filter([
                                { field: 'DevSecOps', operator: 'eq', value: false },
                                { field: 'Recommended_x0020_Reading', operator: 'eq', value: false }
                            ]);
                            // reset maturity filter
                            this.maturityDropDownList.select(0);
                            break;
                        case 'DevSecOps':
                            dsCatalog.filter([
                                { field: 'DevSecOps', operator: 'eq', value: true },
                                { field: 'Recommended_x0020_Reading', operator: 'eq', value: false }
                            ]);
                            // reset maturity filter
                            this.dsoMaturityDropDownList.select(0);
                            break;
                        case 'Recommended Reading':
                            this.readingGrid.dataSource.filter({ field: 'Recommended_x0020_Reading', operator: 'eq', value: true });
                            break;
                        default:
                            appState.tabName = e.item.textContent;
                            appState.redirectUrl = (e.item.textContent == 'SAFe' ? args.safeLink : args.wamLink);

                            this.dialog.content('<p>Click OK if you like to open a new browser window to display the ' + appState.tabName + ' information, otherwise click Cancel...</p>');
                            this.dialog.setOptions({
                                actions: [ 
                                    { text: 'OK', primary: true, action: _ => { this.tabStrip.select(0); this.dialog.open(); window.open(appState.redirectUrl); return true; } },
                                    { text: 'Cancel', action: _ => { this.dialog.close(); this.tabStrip.select(0); return false; } },
                                ]
                            });
                            this.dialog.open();
                    }
                }
            };
            this.tabStrip = $('#tabstrip').kendoTabStrip(this.tabStripOptions).data('kendoTabStrip');

            let org = null;
            let team = null;
            let role = null;

            const dsFilterSharedDataSource = ds({
                guid: args.rolesGuid,
                dsName: 'dsFilterSharedDataSource',
                schema: {
                    id: 'Id',
                    fields: {
                        Id: { type: 'number' },
                        Title: { type: 'string' },
                        Role: { type: 'string' },
                        Subrole: { type: 'string' }
                    }
                }
            });

            dsFilterSharedDataSource.read().then(_ => {
                // Remove duplicate objects from the data source
                org = dsFilterSharedDataSource.data().map(item => ({ Title: item['Title'] }));
                org = org.filter((item, index, array) => array.findIndex(i => i.Title == item.Title) == index);
                org.sort((x,y) => (x.Title > y.Title) ? 1 : -1);

                this.organizationDropDownListOptions = {
                    autoBind: false,
                    dataSource: org,
                    dataTextField: 'Title',
                    dataValueField: 'Title',
                    optionLabel: 'Select your organizational level...',
                    change: e => {
                        // Always set default grid filter
                        dsCatalog.filter({
                            logic: 'and',
                            filters: [
                                { field: 'DevSecOps', operator: 'eq', value: false },
                                { field: 'Recommended_x0020_Reading', operator: 'eq', value: false }
                            ]
                        });

                        // Always Disable the Role DropDownList
                        this.roleDropDownList.enable(false);

                        if (e.sender.value() === 'Product Team Level') {
                            // Filter duplicates
                            team = dsFilterSharedDataSource.data().filter(item => item['Title'] == 'Product Team Level');
                            team = team.filter((item, index, array) => array.findIndex(i => i.Role == item.Role) == index);

                            // Set new data source (no duplicates)
                            this.teamDropDownList.setDataSource(team);
                        } else {
                            // Restore the original data source
                            this.teamDropDownList.setDataSource(dsFilterSharedDataSource);
                        }

                        if (e.sender.value() === 'Account Manager' || e.sender.value() === 'Business Owner') {
                            this.teamDropDownList.enable(false);

                            // Update grid filter for these values
                            dsCatalog.filter().logic = 'and';
                            let filters = dsCatalog.filter().filters;
                            let appendFilter = { field: 'Role_x0028_s_x0029_', operator: 'contains', value: e.sender.value() };
                            filters.push(appendFilter);
                            dsCatalog.filter(filters);
                        }
                    }
                };
                this.organizationDropDownList = $('#organization').kendoDropDownList(this.organizationDropDownListOptions).data('kendoDropDownList');
    
                this.teamDropDownListOptions = {
                    autoBind: false,
                    cascadeFrom: 'organization',
                    dataSource: dsFilterSharedDataSource,
                    dataTextField: 'Role',
                    dataValueField: 'Role',
                    optionLabel: 'Select your team/role...',
                    change: e => {
                        // Always set default grid filter
                        dsCatalog.filter({
                            logic: 'and',
                            filters: [
                                { field: 'DevSecOps', operator: 'eq', value: false },
                                { field: 'Recommended_x0020_Reading', operator: 'eq', value: false }
                            ]
                        });

                        if(['Portfolio Level','Product Line Level','Supporting Role','Technical Manager/Solution Architect'].includes(this.organizationDropDownList.value())) {
                            this.roleDropDownList.enable(false);

                            // Update grid filter for these values
                            dsCatalog.filter().logic = 'and';
                            let filters = dsCatalog.filter().filters;
                            let appendFilter = { field: 'Role_x0028_s_x0029_', operator: 'contains', value: e.sender.value() };
                            filters.push(appendFilter);
                            dsCatalog.filter(filters);
                        } else {
                            // Filter duplicates
                            role = dsFilterSharedDataSource.data().filter(item => item['Role'] == e.sender.value());
                            role = role.filter((item, index, array) => array.findIndex(i => i.Subrole == item.Subrole) == index);

                            // Set new data source (no duplicates)
                            this.roleDropDownList.setDataSource(role);
                            this.roleDropDownList.enable(true);
                        }
                    }
                };
                this.teamDropDownList = $('#team').kendoDropDownList(this.teamDropDownListOptions).data('kendoDropDownList');

                this.roleDropDownListOptions = {
                    autoBind: false,
                    dataSource: dsFilterSharedDataSource,
                    dataTextField: 'Subrole',
                    dataValueField: 'Subrole',
                    optionLabel: 'Select your role...',
                    change: e => {
                        // Always set default grid filter
                        dsCatalog.filter({
                            logic: 'and',
                            filters: [
                                { field: 'DevSecOps', operator: 'eq', value: false },
                                { field: 'Recommended_x0020_Reading', operator: 'eq', value: false }
                            ]
                        });

                        // Update grid filter for these values
                        dsCatalog.filter().logic = 'and';
                        let filters = dsCatalog.filter().filters;
                        let appendFilter = { field: 'Role_x0028_s_x0029_', operator: 'contains', value: e.sender.value() };
                        filters.push(appendFilter);
                        dsCatalog.filter(filters);
                    }
                };
                this.roleDropDownList = $('#role').kendoDropDownList(this.roleDropDownListOptions).data('kendoDropDownList');

                this.maturityDropDownListOptions = {
                    dataSource: dsMaturityDataSource,
                    dataTextField: 'Title',
                    dataValueField: 'Title',
                    optionLabel: 'Select your DSO Maturity Level...',
                    change: e => {
                        // reset default filters
                        let currentFilters = dsCatalog.filter().filters.filter(obj => obj['field'] !== 'DSO_x0020_Maturity_x0020_Level');

                        // append new filter if applicable
                        if (e.sender.value() !== '') currentFilters.push({field: 'DSO_x0020_Maturity_x0020_Level', operator: 'Contains', value: e.sender.value()});
                        dsCatalog.filter(currentFilters);
                    }
                };
                this.maturityDropDownList = $('#byRoleMaturity').kendoDropDownList(this.maturityDropDownListOptions).data('kendoDropDownList');
                this.dsoMaturityDropDownList = $('#dsoMaturity').kendoDropDownList(this.maturityDropDownListOptions).data('kendoDropDownList');
            });
            
            this.catalogGridOptions = {
                dataSource: dsCatalog,
                columnMenu: true,
                editable: false,
                filterable: true,
                groupable: false,
                navigatable: true,
                pageable: {
                    alwaysVisible: true,
                    buttonCount: 3,
                    pageSizes: [5, 10, 20, 'All']
                },
                reorderable: true,
                resizable: true,
                scrollable: { virtual: 'column' },
                sortable: {
                    allowUnsort: false,
                    initialDirection: 'asc',
                    mode: 'single',
                    showIndexes: true
                },
                toolbar: [ 'excel', 'pdf' ],
                excel: {
                    fileName: 'DSO Training Export.xlsx',
                    filterable: true
                },
                pdf: {
                    allPages: true,
                    avoidLinks: false,
                    paperSize: 'letter',
                    margin: { top: '1cm', left: '1cm', right: '1cm', bottom: '1cm' },
                    landscape: false,
                    scale: 1.0
                },
                columns: [
                    { field: 'Title', title: 'Course Name', width: 350, template: dataItem => { if (dataItem.Link_x0020_to_x0020_Resource != '') return '<a href="' + dataItem.Link_x0020_to_x0020_Resource + '" title="Link to course for ' + dataItem.Title + '" target="_blank">' + dataItem.Title + '</a>'; return dataItem.Title; } },
                    { field: 'LearningHours', title: 'Learning Hours', width: 150 },
                    { field: 'Asset_x0020_Type', title: 'Asset Type', width: 300 },
                    { field: 'CourseSeries', title: 'Course Series', width: 150 },
                    { field: 'TMSItemID', title: 'TMS Item ID', width: 175 },
                    { field: 'DSO_x0020_Maturity_x0020_Level', title: 'DSO Maturity Level', width: 225, hidden: false, template: dataItem => { if (dataItem.DSO_x0020_Maturity_x0020_Level != null) return dataItem.DSO_x0020_Maturity_x0020_Level.replaceAll(',', ', '); else return ''; } },
                    //{ field: 'PLM_x0020_Roadmap_x0020_Focus', title: 'Roadmap Focus', width: 225 },
                    { field: 'Role_x0028_s_x0029_', title: 'Roles', hidden: true },
                    { field: 'Keywords', title: 'Keywords', hidden: true }
                ]
            };
            this.catalogGrid = $('#grid').kendoGrid(this.catalogGridOptions).data('kendoGrid');
            this.catalogGrid.dataSource.filter([
                { field: 'DevSecOps', operator: 'eq', value: false },
                { field: 'Recommended_x0020_Reading', operator: 'eq', value: false }
            ]);

            this.devSecOpsGridOptions = {
                dataSource: dsCatalog,
                columnMenu: true,
                editable: false,
                filterable: true,
                groupable: false,
                navigatable: true,
                pageable: {
                    alwaysVisible: true,
                    buttonCount: 3,
                    pageSizes: [5, 10, 20, 'All']
                },
                reorderable: true,
                resizable: true,
                scrollable: { virtual: 'column' },
                sortable: {
                    allowUnsort: false,
                    initialDirection: 'asc',
                    mode: 'single',
                    showIndexes: true
                },
                //toolbar: [ 'search' ],
                columns: [
                    { field: 'Title', title: 'Course Name', width: 350, template: dataItem => { if (dataItem.Link_x0020_to_x0020_Resource != '') return '<a href="' + dataItem.Link_x0020_to_x0020_Resource + '" title="Link to course for ' + dataItem.Title + '" target="_blank">' + dataItem.Title + '</a>'; return dataItem.Title; } },
                    { field: 'LearningHours', title: 'Learning Hours', width: 150 },
                    { field: 'Asset_x0020_Type', title: 'Asset Type', width: 300 },
                    { field: 'CourseSeries', title: 'Course Series', width: 400 },
                    { field: 'TMSItemID', title: 'TMS Item ID', width: 175 },
                    { field: 'DSO_x0020_Maturity_x0020_Level', title: 'DSO Maturity Level', width: 225, hidden: false, template: dataItem => { if (dataItem.DSO_x0020_Maturity_x0020_Level != null) return dataItem.DSO_x0020_Maturity_x0020_Level.replaceAll(',', ', '); else return ''; } },
                    { field: 'Role_x0028_s_x0029_', title: 'Roles', hidden: true },
                    { field: 'Keywords', title: 'Keywords', hidden: true }
                ]

            };
            this.devSecOpsGrid = $('#grid2').kendoGrid(this.devSecOpsGridOptions).data('kendoGrid');

            this.readingGridOptions = {
                dataSource: dsCatalog,
                columnMenu: true,
                editable: false,
                filterable: true,
                groupable: false,
                navigatable: true,
                pageable: {
                    alwaysVisible: true,
                    buttonCount: 3,
                    pageSizes: [5, 10, 20, 'All']
                },
                reorderable: true,
                resizable: true,
                scrollable: { virtual: 'column' },
                sortable: {
                    allowUnsort: false,
                    initialDirection: 'asc',
                    mode: 'single',
                    showIndexes: true
                },
                toolbar: [ 'search' ],
                columns: [
                    { field: 'Title', title: 'Course Name', width: 350, template: dataItem => { if (dataItem.Link_x0020_to_x0020_Resource != '') return '<a href="' + dataItem.Link_x0020_to_x0020_Resource + '" title="Link to course for ' + dataItem.Title + '" target="_blank">' + dataItem.Title + '</a>'; return dataItem.Title; } },
                    { field: 'TMSItemID', title: 'TMS Item ID', width: 150 },
                    { field: 'LearningHours', title: 'Learning Hours', width: 150 },
                    { field: 'Asset_x0020_Type', title: 'Asset Type', width: 300 },
                    { field: 'Role_x0028_s_x0029_', title: 'Roles', hidden: true },
                    { field: 'Keywords', title: 'Keywords', hidden: true }
                ]

            };
            this.readingGrid = $('#grid3').kendoGrid(this.readingGridOptions).data('kendoGrid');
    
            this.dialogOptions = {
                width: 500,
                title: 'Content Redirection',
                closable: false,
                modal: false
            };

            this.dialog = $('#dialog').kendoDialog(this.dialogOptions).data('kendoDialog');
            this.dialog.close();
        });

        return SPA.instance;
    }
}