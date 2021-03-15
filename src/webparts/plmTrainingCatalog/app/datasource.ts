import * as $ from 'jquery';
import '@progress/kendo-ui';
import { sp, objectToSPKeyValueCollection } from '@pnp/sp/presets/all';

interface DataSourceConfig {
    guid: string;
    dsName: string;
    schema: kendo.data.DataSourceSchemaModel;
    pageSize?: number;
    filter?: kendo.data.DataSourceFilterItem;
    group?: kendo.data.DataSourceGroupItem;
    sort?: kendo.data.DataSourceSortItem;
    top?: number;
    expand?: string[];
    expandedFields?: string[];
}

//Object.keys($('#faqGrid').data('kendoGrid').dataSource.options.schema.model.fields)
const cleanseModel = (dataItem: object, spFields: string[]): object => {
    $.each(dataItem, (k,v) => {
        if (spFields.join(',').indexOf(k) == -1) delete dataItem[k];
    });
    return dataItem;
};

export const dsExpand = (args: DataSourceConfig): kendo.data.DataSource => {
    var dataArray = [];

    return new kendo.data.DataSource({
        transport: {
            create: async options => {
                //console.log('create: ', options.data);
                //console.log('keys: ', Object.keys(args.schema.fields));                
                
                await sp.web.lists.getById(args.guid).items.add(cleanseModel(options.data, Object.keys(args.schema.fields)))
                .then(response => {
                    options.success(response.data);
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to create item');
                });
            },
            read: async options => {
                // Handle lookup fields
                let selectedFields = Object.keys(args.schema.fields);
                selectedFields = selectedFields.map(elem => {
                    args.expand.forEach((item, index) => {
                        if (item == elem) return elem = args.expandedFields[index];
                    });
                    return elem;
                });

                await sp.web.lists.getById(args.guid).items.select(selectedFields.join()).expand(args.expand.join()).top(1000).getPaged()
                .then(response => {
                    const recurse = (next: any) => {
                        next.getNext().then(nestedResponse => {
                            dataArray = [...dataArray, ...nestedResponse.results];
                            if (nestedResponse.hasNext) recurse(nestedResponse);
                            else {
                                dataArray.map(e => {
                                    e.Asset_x0020_Type = (e.Asset_x0020_Type != undefined) ? e.Asset_x0020_Type.Title : '';
                                    e.Link_x0020_to_x0020_Resource = (e.Link_x0020_to_x0020_Resource != undefined) ? e.Link_x0020_to_x0020_Resource : '';
                                    e.PLM_x0020_Maturity_x0020_Level = (e.PLM_x0020_Maturity_x0020_Level != undefined) ? e.PLM_x0020_Maturity_x0020_Level.Title : '' ;
                                    e.PLM_x0020_Roadmap_x0020_Focus = (e.PLM_x0020_Roadmap_x0020_Focus != undefined) ? e.PLM_x0020_Roadmap_x0020_Focus.Title : '';
                                    e.Role_x0028_s_x0029_ =  (e.Role_x0028_s_x0029_ != undefined) ? e.Role_x0028_s_x0029_.map(item => item.Subrole).join() : '';
                                });
                                options.success(dataArray);
                            }
                        });
                    };
                    
                    dataArray = response.results;
                    if (response.hasNext) recurse(response);
                    else {
                        dataArray.map(e => {
                            e.Asset_x0020_Type = (e.Asset_x0020_Type != undefined) ? e.Asset_x0020_Type.Title : '';
                            e.Link_x0020_to_x0020_Resource = (e.Link_x0020_to_x0020_Resource != undefined) ? e.Link_x0020_to_x0020_Resource : '';
                            e.PLM_x0020_Maturity_x0020_Level = (e.PLM_x0020_Maturity_x0020_Level != undefined) ? e.PLM_x0020_Maturity_x0020_Level.Title : '' ;
                            e.PLM_x0020_Roadmap_x0020_Focus = (e.PLM_x0020_Roadmap_x0020_Focus != undefined) ? e.PLM_x0020_Roadmap_x0020_Focus.Title : '';
                            e.Role_x0028_s_x0029_ =  (e.Role_x0028_s_x0029_ != undefined) ? e.Role_x0028_s_x0029_.map(item => item.Subrole).join() : '';
                        });
                        options.success(dataArray);
                    }
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to read items');
                });
            },
            update: async options => {
                await sp.web.lists.getById(args.guid).items.getById(options.data.Id).update(cleanseModel(options.data, Object.keys(args.schema.fields)))
                .then(response => {
                    options.success();
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to update item');
                });
            },
            destroy: async options => {
                await sp.web.lists.getById(args.guid).items.getById(options.data.Id).recycle()
                .then(response => {
                    options.success();
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to delete item');
                });
            }
        },
        schema: { model: args.schema },
        pageSize: args.pageSize,
        filter: args.filter,
        group: args.group,
        sort: args.sort
    });
};

export const ds = (args: DataSourceConfig): kendo.data.DataSource => {
    var dataArray = [];

    return new kendo.data.DataSource({
        transport: {
            create: async options => {
                //console.log('create: ', options.data);
                //console.log('keys: ', Object.keys(args.schema.fields));                
                
                await sp.web.lists.getById(args.guid).items.add(cleanseModel(options.data, Object.keys(args.schema.fields)))
                .then(response => {
                    options.success(response.data);
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to create item');
                });
            },
            read: async options => {
                await sp.web.lists.getById(args.guid).items.select(Object.keys(args.schema.fields).join()).top(1000).getPaged()
                .then(response => {
                    const recurse = (next: any) => {
                        next.getNext().then(nestedResponse => {
                            dataArray = [...dataArray, ...nestedResponse.results];
                            if (nestedResponse.hasNext) recurse(nestedResponse);
                            else options.success(dataArray);
                        });
                    };

                    dataArray = response.results;
                    if (response.hasNext) recurse(response);
                    else options.success(dataArray);
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to read items');
                });
            },
            update: async options => {
                await sp.web.lists.getById(args.guid).items.getById(options.data.Id).update(cleanseModel(options.data, Object.keys(args.schema.fields)))
                .then(response => {
                    options.success();
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to update item');
                });
            },
            destroy: async options => {
                await sp.web.lists.getById(args.guid).items.getById(options.data.Id).recycle()
                .then(response => {
                    options.success();
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to delete item');
                });
            }
        },
        schema: { model: args.schema },
        pageSize: args.pageSize,
        filter: args.filter,
        group: args.group,
        sort: args.sort
    });
};