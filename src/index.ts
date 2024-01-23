import { InstallationRequired, LoadingDialog } from "dattatable";
import { Components, ContextInfo, ThemeManager } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { Security } from "./security";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    render: (el, context?, sourceUrl?: string) => {
        // See if the page context exists
        if (context) {
            // Set the context
            setContext(context, sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(sourceUrl || ContextInfo.webServerRelativeUrl);
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
                    new App(el);

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
    }
};

// Make is available in the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Render the application
    GlobalVariable.render(elApp);
}