import { Dashboard, IDashboardProps } from "dattatable";
import { Components } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { DataSource, IListItem } from "./ds";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    private _dashboard: Dashboard = null;

    // Constructor
    constructor(el: HTMLElement, dashboardType?: string) {
        // Render the dashboard
        this.render(el, dashboardType);
    }

    // Refreshes the dashboard
    private refresh() {
        // Refresh the data
        DataSource.refresh().then(() => {
            // Refresh the table
            this._dashboard.refresh(DataSource.ListItems);
        });
    }

    // Renders the dashboard
    private render(el: HTMLElement, dashboardType: string = "table") {
        // Define the dashboard properties
        let props: IDashboardProps = {
            el,
            hideHeader: true,
            useModal: true,
            filters: {
                items: [{
                    header: "By Status",
                    items: DataSource.StatusFilters,
                    onFilter: (value: string) => {
                        // Filter the table
                        this._dashboard.filter(2, value);
                    }
                }]
            },
            navigation: {
                title: Strings.ProjectName,
                items: [
                    {
                        className: "btn-outline-light",
                        text: "Create Item",
                        isButton: true,
                        onClick: () => {
                            // Show the new form
                            DataSource.List.newForm({
                                onUpdate: () => {
                                    // Refresh the dashboard
                                    this.refresh();
                                }
                            });
                        }
                    }
                ]
            },
            footer: {
                itemsEnd: [
                    {
                        text: "v" + Strings.Version
                    }
                ]
            }
        };

        // Set the dashboard type
        switch (dashboardType) {
            case "accordion":
                props.accordion = {
                    items: DataSource.ListItems,
                    titleField: "Title",
                    bodyField: "ItemType",
                    filterField: "Status"
                };
                break;
            case "tiles":
                props.tiles = {
                    items: DataSource.ListItems,
                    colSize: 4,
                    titleField: "Title",
                    subTitleField: "Status",
                    bodyField: "ItemType",
                    filterField: "Status",
                    showFooter: false,
                    showHeader: false
                };
                break;
            default:
                props.table = {
                    rows: DataSource.ListItems,
                    dtProps: {
                        dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                        columnDefs: [
                            {
                                "targets": 0,
                                "orderable": false,
                                "searchable": false
                            }
                        ],
                        createdRow: function (row, data, index) {
                            jQuery('td', row).addClass('align-middle');
                        },
                        drawCallback: function (settings) {
                            let api = new jQuery.fn.dataTable.Api(settings) as any;
                            jQuery(api.context[0].nTable).removeClass('no-footer');
                            jQuery(api.context[0].nTable).addClass('tbl-footer');
                            jQuery(api.context[0].nTable).addClass('table-striped');
                            jQuery(api.context[0].nTableWrapper).find('.dataTables_info').addClass('text-center');
                            jQuery(api.context[0].nTableWrapper).find('.dataTables_length').addClass('pt-2');
                            jQuery(api.context[0].nTableWrapper).find('.dataTables_paginate').addClass('pt-03');
                        },
                        headerCallback: function (thead, data, start, end, display) {
                            jQuery('th', thead).addClass('align-middle');
                        },
                        // Order by the 1st column by default; ascending
                        order: [[1, "asc"]]
                    },
                    columns: [
                        {
                            name: "",
                            title: "Title",
                            onRenderCell: (el, column, item: IListItem) => {
                                // Render a tooltip
                                Components.TooltipGroup({
                                    el,
                                    tooltips: [
                                        {
                                            content: "Click to view the " + item.Title + " item.",
                                            btnProps: {
                                                text: "View",
                                                type: Components.ButtonTypes.OutlinePrimary,
                                                onClick: () => {
                                                    // Show the display form
                                                    DataSource.List.viewForm({
                                                        itemId: item.Id
                                                    });
                                                }
                                            }
                                        },
                                        {
                                            content: "Click to edit the item.",
                                            btnProps: {
                                                text: "Edit",
                                                type: Components.ButtonTypes.OutlineSuccess,
                                                onClick: () => {
                                                    // Show the edit form
                                                    DataSource.List.editForm({
                                                        itemId: item.Id,
                                                        onUpdate: () => {
                                                            // Refresh the dashboard
                                                            this.refresh();
                                                        }
                                                    });
                                                }
                                            }
                                        }
                                    ]
                                });
                            }
                        },
                        {
                            name: "ItemType",
                            title: "Item Type"
                        },
                        {
                            name: "Status",
                            title: "Status"
                        }
                    ]
                }
                break;
        }

        // Create the dashboard
        this._dashboard = new Dashboard(props);
    }
}