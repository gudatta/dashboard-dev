import { List } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import { Security } from "./security";
import Strings from "./strings";

/**
 * List Item
 * Add your custom fields here
 */
export interface IListItem extends Types.SP.ListItem {
    ItemType: string;
    Status: string;
}

/**
 * Data Source
 */
export class DataSource {
    // List
    private static _list: List<IListItem> = null;
    static get List(): List<IListItem> { return this._list; }

    // List Items
    static get ListItems(): IListItem[] { return this.List.Items; }

    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = null;
    static get StatusFilters(): Components.ICheckboxGroupItem[] { return this._statusFilters; }
    static loadStatusFilters() {
        // Clear the filters
        this._statusFilters = [];

        // Get the status field
        let fld: Types.SP.FieldChoice = this.List.getField("Status");

        // Parse the choices
        for (let i = 0; i < fld.Choices.results.length; i++) {
            // Add an item
            this._statusFilters.push({
                label: fld.Choices.results[i],
                type: Components.CheckboxGroupTypes.Switch
            });
        }
    }

    // Initializes the application
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Initialize the list
            this._list = new List<IListItem>({
                listName: Strings.Lists.Main,
                itemQuery: {
                    GetAllItems: true,
                    OrderBy: ["Title"],
                    Top: 5000
                },
                onInitError: reject,
                onInitialized: () => {
                    // Load the status filters
                    this.loadStatusFilters();

                    // Initialize the security
                    Security.init().then(resolve, reject);
                }
            });
        });
    }

    // Refreshes the list data
    static refresh(): PromiseLike<IListItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Refresh the data
            DataSource.List.refresh().then(resolve, reject);
        });
    }
}