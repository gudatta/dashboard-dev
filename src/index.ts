import { InstallationRequired, LoadingDialog } from "dattatable";
import { Components, ContextInfo, ThemeManager } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { Security } from "./security";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// App Properties
interface IAppProps {
    el: HTMLElement;
    context?: any;
    dashboardType?: string;
    listName?: string;
    sourceUrl?: string
}

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    getVersion: () => {
        // Return the version #
        return Strings.Version;
    },
    render: (props: IAppProps) => {
        // Clear the element
        while (props.el.firstChild) { props.el.removeChild(props.el.firstChild); }

        // See if the page context exists
        if (props.context) {
            // Set the context
            setContext(props.context, props.sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(props.sourceUrl || ContextInfo.webServerRelativeUrl);
        }

        // Show a loading dialog
        LoadingDialog.setHeader("Loading Application");
        LoadingDialog.setBody("This will close after the data is loaded...");
        LoadingDialog.show();

        // Initialize the application
        DataSource.init().then(
            // Success
            () => {
                // Load the current theme and apply it to the components
                ThemeManager.load(true).then(() => {
                    // Create the application
                    new App(props.el, props.dashboardType);

                    // Hide the loading dialog
                    LoadingDialog.hide();
                });
            },

            // Error
            () => {
                // Update the loading dialog
                LoadingDialog.setHeader("Error Loading App");
                LoadingDialog.setBody("Checking to see what went wrong...");

                // See if an installation is required
                InstallationRequired.requiresInstall({ cfg: Configuration }).then(installFl => {
                    let customErrors: Components.IListGroupItem[] = [];

                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // See if the custom groups exist
                    if (Security.ManagerGroup == null) {
                        // Add a custom error
                        customErrors.push({
                            content: "The security groups have not been created."
                        });
                    }

                    // See if an install is required
                    if (installFl || customErrors.length > 0) {
                        // Show the dialog
                        InstallationRequired.showDialog({
                            errors: customErrors,
                            onFooterRendered: el => {
                                // Append a button
                                Components.Tooltip({
                                    el,
                                    content: "Click to configure the security.",
                                    btnProps: {
                                        text: "Security",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the security dialog
                                            Security.show(() => {
                                                // Refresh the page
                                                window.location.reload();
                                            });
                                        }
                                    }
                                });
                            }
                        });
                    } else {
                        // Log
                        console.error("[" + Strings.ProjectName + "] Error initializing the solution.");
                    }
                });
            }
        );
    },
    updateTheme: (themeInfo) => {
        // Update the theme
        ThemeManager.setCurrentTheme(themeInfo, true);
    }
};

// Make is available in the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Render the application
    GlobalVariable.render({ el: elApp });
}