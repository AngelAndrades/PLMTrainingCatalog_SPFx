import * as $ from 'jquery';
import * as JSZip from 'jszip';
import '@progress/kendo-ui';
import { ds, dsExpand } from  './datasource';
import { sp, objectToSPKeyValueCollection } from '@pnp/sp/presets/all';
import { data } from 'jquery';

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
    safeGuid: string;
    wamLink: string;
    enableMaturity: boolean;
    enableExport: boolean;
}

export class SPA {
    protected static catalogGridOptions: kendo.ui.GridOptions;
    protected static catalogGrid: kendo.ui.Grid;
    protected static devSecOpsGridOptions: kendo.ui.GridOptions;
    protected static devSecOpsGrid: kendo.ui.Grid;
    protected static readingGridOptions: kendo.ui.GridOptions;
    protected static readingGrid: kendo.ui.Grid;
    protected static wamGridOptions: kendo.ui.GridOptions;
    protected static wamGrid: kendo.ui.Grid;
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
    protected static safeGridOptions: kendo.ui.GridOptions;
    protected static safeGrid: kendo.ui.Grid;
    private static instance: SPA;

    constructor() {}

    public static getInstance(args: Params): SPA {
        var appState = new ModelState();

        // Toolbar configuration
        let toolbar = [];
        let readingToolbar = ['search'];
        if (args.enableExport) {
            //toolbar.push('excel');
            toolbar.push('pdf');
            //readingToolbar.push('excel');
            readingToolbar.push('pdf');
        }

        // Required for Excel Export to work with Grid
        window['JSZip'] = JSZip;

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
            expandedFields: ['Role_x0028_s_x0029_/Subrole','Asset_x0020_Type/Title','PLM_x0020_Roadmap_x0020_Focus/Title']
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

        const dsSafe = dsExpand({
            guid: args.safeGuid,
            dsName: 'dsSafe',
            schema: {
                id: 'Id',
                fields: {
                    Id: { type: 'number' },
                    Title: { type: 'string' },
                    TrainingLink: { type: 'object' },
                    Certification: { type: 'string' },
                    CourseLevel: { type: 'string' },
                    LearningHours: { type: 'string' },
                    Roles: { type: 'object' }
                }
            },
            sort: { field: 'Title', dir: 'asc' },
            expand: ['Roles'],
            expandedFields: ['Roles/Title']
        });
        //dsSafe.read().then(res => console.log(res));

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
                            if (args.enableMaturity) this.maturityDropDownList.select(0);
                            break;
                        case 'DevSecOps':
                            dsCatalog.filter([
                                { field: 'DevSecOps', operator: 'eq', value: true },
                                { field: 'Recommended_x0020_Reading', operator: 'eq', value: false }
                            ]);
                            // reset maturity filter
                            if (args.enableMaturity) this.dsoMaturityDropDownList.select(0);
                            break;
                        case 'Recommended Reading':
                            this.readingGrid.dataSource.filter({ field: 'Recommended_x0020_Reading', operator: 'eq', value: true });
                            break;
                        case 'SAFe':
                            break;
                        default:
                            appState.tabName = e.item.textContent;
                            appState.redirectUrl = (e.item.textContent == 'SAFe' ? args.safeGuid : args.wamLink);

                            if (appState.redirectUrl.startsWith('http'))
                            {
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

                        // Always Disable and Reset the Role DropDownList
                        this.roleDropDownList.enable(false);
                        this.roleDropDownList.setDataSource(role);

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

                        if (e.sender.value() === 'Account Manager' || e.sender.value() === 'Business Owner' || e.sender.value() === 'Product Owners' ) {
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
                if (args.enableMaturity) {
                    this.maturityDropDownList = $('#byRoleMaturity').kendoDropDownList(this.maturityDropDownListOptions).data('kendoDropDownList');
                    this.dsoMaturityDropDownList = $('#dsoMaturity').kendoDropDownList(this.maturityDropDownListOptions).data('kendoDropDownList');
                } else {
                    $('#byRoleMaturity').hide();
                    $('#dsoFilters').hide();
                }
                
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
                toolbar: toolbar,
                excel: {
                    fileName: 'DSO ByRole Training Export.xlsx',
                    filterable: true,
                    allPages: true
                },
                pdf: {
                    fileName: 'DSO ByRole Training Export.pdf',
                    allPages: true,
                    avoidLinks: false,
                    paperSize: 'letter',
                    margin: { top: '1cm', left: '1cm', right: '1cm', bottom: '1cm' },
                    landscape: true,
                    scale: 0.5
                },
                columns: [
                    { field: 'Title', title: 'Course Name', width: 350, template: dataItem => { if (dataItem.Link_x0020_to_x0020_Resource != '') return '<a href="' + dataItem.Link_x0020_to_x0020_Resource + '" title="Link to course for ' + dataItem.Title + '" target="_blank">' + dataItem.Title + '</a>'; return dataItem.Title; } },
                    { field: 'LearningHours', title: 'Learning Hours', width: 150 },
                    { field: 'Asset_x0020_Type', title: 'Asset Type', width: 300 },
                    { field: 'CourseSeries', title: 'Course Series', width: 150 },
                    { field: 'TMSItemID', title: 'TMS Item ID', width: 175 },
                    { field: 'DSO_x0020_Maturity_x0020_Level', title: 'DSO Maturity Level', width: 225, hidden: true, template: dataItem => { if (dataItem.DSO_x0020_Maturity_x0020_Level != null) return dataItem.DSO_x0020_Maturity_x0020_Level.replaceAll(',', ', '); else return ''; } },
                    //{ field: 'PLM_x0020_Roadmap_x0020_Focus', title: 'Roadmap Focus', width: 225 },
                    { field: 'Role_x0028_s_x0029_', title: 'Roles', width: 200, hidden: true },
                    { field: 'Link_x0020_to_x0020_Resource', title: 'Link', width: 500 },
                    { field: 'Keywords', title: 'Keywords', hidden: true, exportable: false }
                ],

                excelExport: e => {
                    this.catalogGrid.showColumn('Link_x0020_to_x0020_Resource');
                }
            };
            this.catalogGrid = $('#grid').kendoGrid(this.catalogGridOptions).data('kendoGrid');
            this.catalogGrid.hideColumn('Link_x0020_to_x0020_Resource');
            if(args.enableMaturity) this.catalogGrid.showColumn('DSO_x0020_Maturity_x0020_Level');
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
                toolbar: toolbar,
                excel: {
                    fileName: 'DevSecOps Training Export.xlsx',
                    filterable: true
                },
                pdf: {
                    fileName: 'DevSecOps Training Export.pdf',
                    allPages: true,
                    avoidLinks: false,
                    paperSize: 'letter',
                    margin: { top: '1cm', left: '1cm', right: '1cm', bottom: '1cm' },
                    landscape: true,
                    scale: 0.5
                },
                columns: [
                    { field: 'Title', title: 'Course Name', width: 350, template: dataItem => { if (dataItem.Link_x0020_to_x0020_Resource != '') return '<a href="' + dataItem.Link_x0020_to_x0020_Resource + '" title="Link to course for ' + dataItem.Title + '" target="_blank">' + dataItem.Title + '</a>'; return dataItem.Title; } },
                    { field: 'LearningHours', title: 'Learning Hours', width: 150 },
                    { field: 'Asset_x0020_Type', title: 'Asset Type', width: 300 },
                    { field: 'CourseSeries', title: 'Course Series', width: 400 },
                    { field: 'TMSItemID', title: 'TMS Item ID', width: 175 },
                    { field: 'DSO_x0020_Maturity_x0020_Level', title: 'DSO Maturity Level', width: 225, hidden: true, template: dataItem => { if (dataItem.DSO_x0020_Maturity_x0020_Level != null) return dataItem.DSO_x0020_Maturity_x0020_Level.replaceAll(',', ', '); else return ''; } },
                    { field: 'Role_x0028_s_x0029_', title: 'Roles', hidden: true },
                    { field: 'Keywords', title: 'Keywords', hidden: true, exportable: false }
                ]

            };
            this.devSecOpsGrid = $('#grid2').kendoGrid(this.devSecOpsGridOptions).data('kendoGrid');
            if(args.enableMaturity) this.devSecOpsGrid.showColumn('DSO_x0020_Maturity_x0020_Level');

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
                toolbar: readingToolbar,
                excel: {
                    fileName: 'DevSecOps Recommended Reading Export.xlsx',
                    filterable: true
                },
                pdf: {
                    fileName: 'DevSecOps Recommended Reading Export.pdf',
                    allPages: true,
                    avoidLinks: false,
                    paperSize: 'letter',
                    margin: { top: '1cm', left: '1cm', right: '1cm', bottom: '1cm' },
                    landscape: true,
                    scale: 0.5
                },
                
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

            if (!args.wamLink.startsWith('http'))
            {
                const dsWAM = new kendo.data.DataSource({
                    transport: {
                        read: async options => {
                            // This library belongs to the Agile Coaching subsite, changing context for this call
                            const isolatedSP = await sp.createIsolated();
                            isolatedSP.setup({
                                sp: {
                                    baseUrl: 'https://dvagov.sharepoint.com/sites/OITACOEPortal/agilecoach'
                                }
                            });

                            isolatedSP.web.lists.getById(args.wamLink).items.select('Title,Topic,DateofPresentation,FileRef,Recording').filter("Topic ne 'Archived' and Topic ne 'Recordings'").top(5000).getAll()
                            .then(response => {
                                options.success(response);
                            })
                            .catch(err => {
                                console.log('Error accessing Agile Coaching site: ', err);
                            });
                        }
                    },
                    schema: {
                        model: {
                            fields: {
                                Title: { type: 'string' },
                                Topic: { type: 'string' },
                                DateofPresentation: { type: 'date' },
                                FileRef: {type: 'string' },
                                Recording: { type: 'Object' }
                            }
                        }
                    },
                    pageSize: 10,
                    sort: [
                        { field: 'DateofPresentation', dir: 'desc'}
                    ]
                });
                
                this.wamGridOptions = {
                    dataSource: dsWAM,
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
                    toolbar: readingToolbar,
                    excel: {
                        fileName: 'Weekly Agile Meetings Export.xlsx',
                        filterable: true
                    },
                    pdf: {
                        fileName: 'Weekly Agile Meetings Export.pdf',
                        allPages: true,
                        avoidLinks: false,
                        paperSize: 'letter',
                        margin: { top: '1cm', left: '1cm', right: '1cm', bottom: '1cm' },
                        landscape: true,
                        scale: 0.5
                    },
    

                    columns: [
                        { field: 'Title', title: 'File Name', width: 400, template: dataItem => { 
                            if(dataItem.Title === null) 
                                return '<a href="' + dataItem.FileRef + '" target="_blank">' + dataItem.FileRef.substring(dataItem.FileRef.lastIndexOf('/' + 1)) + '</a>';
                            else
                                return '<a href="' + dataItem.FileRef + '" title="Link to presentation for ' + dataItem.Title + '" target="_blank">' + dataItem.Title + '</a>';
                        }},
                        { field: 'Topic', title: 'Topic', width: 350 },
                        { field: 'DateofPresentation', title: 'Date', width: 150, template: '#= new Date(DateofPresentation.getTime() + DateofPresentation.getTimezoneOffset()*60000).toLocaleString("en-US", { month: "short", day: "numeric", year: "numeric" }) #' },
                        { field: 'Recording', template: (dataItem: any) => { if (dataItem.Recording !== null) return '<a href="' + dataItem.Recording.Url + '" title="link to ' + dataItem.Recording.Description + ' video stream" target="_blank">Video Stream</a>'; else return 'N/A'; } }
                    ]
                    /*
                    dataBound: e => {
                        this.wamGrid.tbody.find('tr.k-master-row').each((idx,elem) => {
                            this.wamGrid.collapseRow(elem);
                        });
                    }*/
                };
                this.wamGrid = $('#grid4').kendoGrid(this.wamGridOptions).data('kendoGrid');    

            }

            this.safeGridOptions = {
                dataSource: dsSafe,
                columnMenu: true,
                editable: false,
                filterable: true,
                groupable: false,
                navigatable: true,
                pageable: false,
                reorderable: true,
                resizable: true,
                scrollable: { virtual: 'column' },
                sortable: {
                    allowUnsort: false,
                    initialDirection: 'asc',
                    mode: 'single',
                    showIndexes: true
                },
                toolbar: readingToolbar,
                excel: {
                    fileName: 'SAFe Training Export.xlsx',
                    filterable: true
                },
                pdf: {
                    fileName: 'SAFe Training Export.pdf',
                    allPages: true,
                    avoidLinks: false,
                    paperSize: 'letter',
                    margin: { top: '1cm', left: '1cm', right: '1cm', bottom: '1cm' },
                    landscape: true,
                    scale: 0.5
                },
                
                columns: [
                    { field: 'Title', title: 'Course Name', width: 350, template: dataItem => { return '<a href="' + dataItem.TrainingLink.Url + '" title="Link to course for ' + dataItem.Title + '" target="_blank">' + dataItem.Title + '</a>'; } },
                    { field: 'Certification', title: 'Certification', width: 150 },
                    { field: 'CourseLevel', title: 'Course Level', width: 200 },
                    { field: 'LearningHours', title: 'Learning Hours', width: 200 },
                    { field: 'Roles', title: 'Roles', width: 250, template: dataItem => { return dataItem.Roles.map(item => item.Title).join(', '); } },
                ]
            };
            this.safeGrid = $('#grid5').kendoGrid(this.safeGridOptions).data('kendoGrid');
        });

        return SPA.instance;
    }
}