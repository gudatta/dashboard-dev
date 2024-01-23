import { ContextInfo } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the source url
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "sp-dashboard",
    GlobalVariable: "SPDashboard",
    Lists: {
        Main: "Dashboard"
    },
    SecurityGroups: {
        Managers: {
            Name: "Dashboard Admins",
            Description: "Group containing all of the dashboard admins."
        }
    },
    ProjectName: "SP Dashboard",
    ProjectDescription: "Created using the gd-sprest-bs library.",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.1"
};
export default Strings;