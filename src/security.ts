import { ListSecurity, ListSecurityDefaultGroups } from "dattatable";
import { ContextInfo, SPTypes, Types } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * Security
 * Code related to the security groups the user belongs to.
 */
export class Security {
    private static _listSecurity: ListSecurity = null;

    // Admin
    private static _isAdmin: boolean = false;
    static get IsAdmin(): boolean { return this._isAdmin; }

    // Dashboard Managers
    private static _isManager: boolean = false;
    static get IsManager(): boolean { return this._isManager; }
    private static _managerGroup: Types.SP.Group = null;
    static get ManagerGroup(): Types.SP.Group { return this._managerGroup; }
    private static _managerGroupInfo: Types.SP.GroupCreationInformation = {
        AllowMembersEditMembership: false,
        Description: Strings.SecurityGroups.Managers.Description,
        OnlyAllowMembersViewMembership: false,
        Title: Strings.SecurityGroups.Managers.Name
    };

    // Members
    private static _memberGroup: Types.SP.Group = null;
    static get MemberGroup(): Types.SP.Group { return this._memberGroup; }

    // Owners
    private static _ownerGroup: Types.SP.Group = null;
    static get OwnerGroup(): Types.SP.Group { return this._ownerGroup; }

    // Visitors
    private static _visitorGroup: Types.SP.Group = null;
    static get VisitorGroup(): Types.SP.Group { return this._visitorGroup; }

    // Initializes the class
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            this._listSecurity = new ListSecurity({
                groups: [this._managerGroupInfo],
                listItems: [
                    {
                        listName: Strings.Lists.Main,
                        groupName: this._managerGroupInfo.Title,
                        permission: SPTypes.RoleType.WebDesigner
                    },
                    {
                        listName: Strings.Lists.Main,
                        groupName: ListSecurityDefaultGroups.Owners,
                        permission: SPTypes.RoleType.Administrator
                    },
                    {
                        listName: Strings.Lists.Main,
                        groupName: ListSecurityDefaultGroups.Members,
                        permission: SPTypes.RoleType.Contributor
                    },
                    {
                        listName: Strings.Lists.Main,
                        groupName: ListSecurityDefaultGroups.Visitors,
                        permission: SPTypes.RoleType.Contributor
                    }
                ],
                onGroupsLoaded: () => {
                    // Set the groups
                    this._managerGroup = this._listSecurity.getGroup(this._managerGroupInfo.Title);
                    this._memberGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Members);
                    this._ownerGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Owners);
                    this._visitorGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Visitors);

                    // Set the user flags
                    this._isAdmin = this._listSecurity.CurrentUser.IsSiteAdmin || this._listSecurity.isInGroup(ContextInfo.userId, ListSecurityDefaultGroups.Owners);
                    this._isManager = this._listSecurity.isInGroup(ContextInfo.userId, this._managerGroupInfo.Title);

                    // Ensure the groups exist
                    if (this._ownerGroup && this.ManagerGroup && this._memberGroup && this._visitorGroup) {
                        // Resolve the request
                        resolve();
                    } else {
                        // Reject the request
                        reject();
                    }
                }
            });
        });
    }

    // Displays the security group configuration
    static show(onComplete: () => void) {
        // Create the groups
        this._listSecurity.show(true, onComplete);
    }
}