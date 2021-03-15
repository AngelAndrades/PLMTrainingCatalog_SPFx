import * as $ from 'jquery';
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
    safeLink: string;
    wamLink: string;
    maturity: boolean;
}

export class SPA {
    protected static catalogGridOptions: kendo.ui.GridOptions;
    protected static catalogGrid: kendo.ui.Grid;
    protected static devSecOpsGridOptions: kendo.ui.GridOptions;
    protected static devSecOpsGrid: kendo.ui.Grid;
    protected static organizationDropDownListOptions: kendo.ui.DropDownListOptions;
    protected static organizationDropDownList: kendo.ui.DropDownList;
    protected static teamDropDownListOptions: kendo.ui.DropDownListOptions;
    protected static teamDropDownList: kendo.ui.DropDownList;
    protected static roleDropDownListOptions: kendo.ui.DropDownListOptions;
    protected static roleDropDownList: kendo.ui.DropDownList;
    protected static tabStrip: kendo.ui.TabStrip;
    protected static tabStripOptions: kendo.ui.TabStripOptions;
    protected static dialog: kendo.ui.Dialog;
    protected static dialogOptions: kendo.ui.DialogOptions;
    private static instance: SPA;

    constructor() {}

    public static getInstance(args: Params): SPA {
        var appState = new ModelState();

        $(() => {
            this.tabStripOptions = {
                tabPosition: 'top',
                animation: { open: { effects: 'fadeIn' } },
                navigatable: true,
                select: e => {
                    switch (e.item.textContent) {
                        case 'By Role':
                            this.catalogGrid.dataSource.filter({ field: 'DevSecOps', operator: 'eq', value: false });
                            break;
                        case 'DevSecOps':
                            this.catalogGrid.dataSource.filter({ field: 'DevSecOps', operator: 'eq', value: true });
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

                this.organizationDropDownListOptions = {
                    autoBind: false,
                    dataSource: org,
                    dataTextField: 'Title',
                    dataValueField: 'Title',
                    optionLabel: 'Select your organizational level...',
                    change: e => {
                        // Always set default grid filter
                        this.catalogGrid.dataSource.filter({ field: 'DevSecOps', operator: 'eq', value: false });

                        // Always Disable the Role DropDownList
                        this.roleDropDownList.enable(false);

                        if (e.sender.value() === 'Product Team') {
                            // Filter duplicates
                            team = dsFilterSharedDataSource.data().filter(item => item['Title'] == 'Product Team');
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
                            this.catalogGrid.dataSource.filter().logic = 'and';
                            let filters = this.catalogGrid.dataSource.filter().filters;
                            let appendFilter = { field: 'Role_x0028_s_x0029_', operator: 'contains', value: e.sender.value() };
                            filters.push(appendFilter);
                            this.catalogGrid.dataSource.filter(filters);
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
                        this.catalogGrid.dataSource.filter({ field: 'DevSecOps', operator: 'eq', value: false });

                        if (this.organizationDropDownList.value() === 'Portfolio Team' || this.organizationDropDownList.value() === 'Product Line Team') {
                            this.roleDropDownList.enable(false);

                            // Update grid filter for these values
                            this.catalogGrid.dataSource.filter().logic = 'and';
                            let filters = this.catalogGrid.dataSource.filter().filters;
                            let appendFilter = { field: 'Role_x0028_s_x0029_', operator: 'contains', value: e.sender.value() };
                            filters.push(appendFilter);
                            this.catalogGrid.dataSource.filter(filters);
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
                        this.catalogGrid.dataSource.filter({ field: 'DevSecOps', operator: 'eq', value: false });

                        // Update grid filter for these values
                        this.catalogGrid.dataSource.filter().logic = 'and';
                        let filters = this.catalogGrid.dataSource.filter().filters;
                        let appendFilter = { field: 'Role_x0028_s_x0029_', operator: 'contains', value: e.sender.value() };
                        filters.push(appendFilter);
                        this.catalogGrid.dataSource.filter(filters);
                    }
                };
                this.roleDropDownList = $('#role').kendoDropDownList(this.roleDropDownListOptions).data('kendoDropDownList');
            });
            
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
                        PLM_x0020_Maturity_x0020_Level: { type: 'string' },
                        PLM_x0020_Roadmap_x0020_Focus: { type: 'string' },
                        Link_x0020_to_x0020_Resource: { type: 'string' },
                        Role_x0028_s_x0029_: { type: 'string' },
                        Keywords: { type: 'string' },
                        TMSItemID: { type: 'string' },
                        DevSecOps: { type: 'boolean' }
                    }
                },
                pageSize: 5,
                sort: { field: 'Title', dir: 'asc' },
                expand: ['Role_x0028_s_x0029_','Asset_x0020_Type','PLM_x0020_Maturity_x0020_Level','PLM_x0020_Roadmap_x0020_Focus'],
                expandedFields: ['Role_x0028_s_x0029_/Subrole','Asset_x0020_Type/Title','PLM_x0020_Maturity_x0020_Level/Title','PLM_x0020_Roadmap_x0020_Focus/Title'],
                //filter: { field: 'DevSecOps', operator: 'equals', value: 0 }
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
                    { field: 'LearningHours', title: 'Learning Hours', width: 150 },
                    { field: 'Asset_x0020_Type', title: 'Asset Type', width: 300 },
                    { field: 'CourseSeries', title: 'Course Series', width: 400 },
                    { field: 'TMSItemID', title: 'TMS Item ID', width: 150 },
                    { field: 'PLM_x0020_Maturity_x0020_Level', title: 'Maturity Level', width: 150 },
                    { field: 'PLM_x0020_Roadmap_x0020_Focus', title: 'Roadmap Focus', width: 225 },
                    { field: 'Role_x0028_s_x0029_', title: 'Roles', hidden: true },
                    { field: 'Keywords', title: 'Keywords', hidden: true }
                ]
            };

            this.catalogGrid = $('#grid').kendoGrid(this.catalogGridOptions).data('kendoGrid');
            this.catalogGrid.dataSource.filter({ field: 'DevSecOps', operator: 'eq', value: false });

            // Property page switch to remove maturity grouping switch and hide maturity level & roadmap focus columns
            if (!args.maturity) 
            {
                this.catalogGrid.setOptions({
                    toolbar: [ 
                        { template: '<div style="display: inline-block; margin-top: 10px;"><input type="checkbox" id="maturity-switch" aria-label="Maturity Level Grouping" /> Group By Maturity Level</div>' },
                        'search'
                    ]
                });
                $('#maturity-switch').kendoSwitch({
                change: e => {
                    if (e.checked) this.catalogGrid.dataSource.group({ field: 'PLM_x0020_Maturity_x0020_Level' });
                    else this.catalogGrid.dataSource.group([]);
                }
                });
                this.catalogGrid.showColumn('PLM_x0020_Maturity_x0020_Level');
                this.catalogGrid.showColumn('PLM_x0020_Roadmap_x0020_Focus');
            } else {
                this.catalogGrid.setOptions({
                    toolbar: [ 'search' ]
                });
                this.catalogGrid.hideColumn('PLM_x0020_Maturity_x0020_Level');
                this.catalogGrid.hideColumn('PLM_x0020_Roadmap_x0020_Focus');
            }

            this.devSecOpsGridOptions = {
                dataSource: dsCatalog,
                columnMenu: true,
                editable: false,
                filterable: true,
                groupable: false,
                navigatable: true,
                pageable: {
                    alwaysVisible: true,
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
                    { field: 'LearningHours', title: 'Learning Hours', width: 150 },
                    { field: 'Asset_x0020_Type', title: 'Asset Type', width: 300 },
                    { field: 'CourseSeries', title: 'Course Series', width: 400 },
                    { field: 'TMSItemID', title: 'TMS Item ID', width: 150 },
                    { field: 'Role_x0028_s_x0029_', title: 'Roles', hidden: true },
                    { field: 'Keywords', title: 'Keywords', hidden: true }
                ]

            };
            this.devSecOpsGrid = $('#grid2').kendoGrid(this.devSecOpsGridOptions).data('kendoGrid');
    
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