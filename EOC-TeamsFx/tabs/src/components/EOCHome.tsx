import { initializeIcons, MessageBar, MessageBarType } from "@fluentui/react";
import { FluentProvider, teamsDarkTheme, teamsHighContrastTheme, teamsLightTheme, Theme } from "@fluentui/react-components";
import { Button, Dialog, Loader } from "@fluentui/react-northstar";
import loadable from "@loadable/component";
import { ApplicationInsights } from "@microsoft/applicationinsights-web";
import { Graph, Providers, ProviderState } from "@microsoft/mgt-element";
import { SimpleProvider } from "@microsoft/mgt-react";
import { Client, Context, HTTPMessageHandler, Middleware, RedirectHandler, RedirectHandlerOptions, RetryHandler, RetryHandlerOptions } from "@microsoft/microsoft-graph-client";
import * as microsoftTeams from "@microsoft/teams-js";
import "bootstrap/dist/css/bootstrap.min.css";
import React from "react";
import LocalizedStrings from "react-localization";
import CommonService, { IListItem } from "../common/CommonService";
import * as constants from "../common/Constants";
import * as graphConfig from "../common/graphConfig";
import siteConfig from "../config/siteConfig.json";
import { localizedStrings } from "../locale/LocaleStrings";
import "../scss/EOCHome.module.scss";
import EocHeader from "./EocHeader";

const Dashboard = loadable(() => import("./Dashboard"));
const ActiveBridge = loadable(() => import("./ActiveBridge"));
const AdminSettings = loadable(() => import("./AdminSettings"));
const IncidentDetails = loadable(() => import("./IncidentDetails"));
const IncidentHistory = loadable(() => import("./IncidentHistory"));


initializeIcons();
//Global Variables
let appInsights: ApplicationInsights;
//Get site name from ARMS template(environment variable)
//Replace spaces from environment variable to get site URL
let siteName = process.env.REACT_APP_SHAREPOINT_SITE_NAME?.toString().replace(/\s+/g, '');

//Get graph base URL from ARMS template(environment variable)
let graphBaseURL = process.env.REACT_APP_GRAPH_BASE_URL?.toString().replace(/\s+/g, '');
graphBaseURL = graphBaseURL || constants.defaultGraphBaseURL;

interface IEOCHomeState {
    showLoginPage: boolean;
    graph: Client;
    tenantName: string;
    graphContextURL: string;
    siteId: string;
    showIncForm: boolean;
    showSuccessMessageBar: boolean;
    showErrorMessageBar: boolean;
    successMessage: string;
    errorMessage: string;
    locale: string;
    currentUserName: string;
    currentUserId: string;
    loaderMessage: string;
    selectedIncident: any;
    existingTeamMembers: any;
    isOwner: boolean;
    isEditMode: boolean,
    showLoader: boolean,
    showNoAccessMessage: boolean;
    userPrincipalName: any;
    showAdminSettings: boolean;
    showIncidentHistory: boolean;
    incidentId: string;
    showActiveBridge: boolean;
    isOwnerOrMember: boolean;
    currentUserDisplayName: string;
    currentUserEmail: string;
    isRolesEnabled: boolean;
    isUserAdmin: boolean;
    configRoleData: any;
    settingsLoader: boolean;
    tenantID: any;
    currentTeamsTheme: Theme;
    currentThemeName: string;
    activeDashboardIncidentId: string;
    fromActiveDashboardTab: boolean;
    appSettings: any;
    isMapViewerEnabled: boolean;
    azureMapsKeyConfigData: any;
    appTitle: string;
    appTitleData: any;
    editIncidentAccessRole: string;
    editIncidentAccessRoleData: any;
    authError: string;
}

interface IEOCHomeProps {
    teamsUserCredential: any;
}

let localeStrings = new LocalizedStrings(localizedStrings);

class SSOAuthMiddleware implements Middleware {
    private next!: Middleware;
    setNext(next: Middleware) { this.next = next; }
    async execute(context: Context): Promise<void> {
        const token = await new Promise<string>((resolve, reject) => {
            microsoftTeams.authentication.getAuthToken({
                successCallback: resolve,
                failureCallback: (err) => reject(new Error(err))
            });
        });
        const headers = new Headers((context.options as RequestInit).headers as HeadersInit);
        headers.set("Authorization", `Bearer ${token}`);
        (context.options as RequestInit).headers = headers;
        return this.next.execute(context);
    }
}

export default class EOCHome extends React.Component<IEOCHomeProps, IEOCHomeState> {
    private dataService = new CommonService();
    private successMessagebarRef: React.RefObject<HTMLDivElement>;
    private errorMessagebarRef: React.RefObject<HTMLDivElement>;
    private _tokenRefreshInterval: number | undefined;

    constructor(props: any) {
        super(props);

        this.successMessagebarRef = React.createRef();
        this.errorMessagebarRef = React.createRef();
        this.state = {
            showLoginPage: false,
            graph: null as any,
            tenantName: '',
            graphContextURL: '',
            siteId: '',
            showIncForm: false,
            showSuccessMessageBar: false,
            showErrorMessageBar: false,
            successMessage: "",
            errorMessage: "",
            locale: "",
            currentUserName: "",
            currentUserId: "",
            loaderMessage: localeStrings.genericLoaderMessage,
            selectedIncident: [],
            existingTeamMembers: [],
            isOwner: false,
            isEditMode: false,
            showLoader: false,
            showNoAccessMessage: false,
            userPrincipalName: null,
            showAdminSettings: false,
            showIncidentHistory: false,
            incidentId: "",
            showActiveBridge: false,
            isOwnerOrMember: false,
            currentUserDisplayName: "",
            currentUserEmail: "",
            isRolesEnabled: false,
            isUserAdmin: false,
            configRoleData: {},
            settingsLoader: false,
            tenantID: "",
            currentTeamsTheme: teamsLightTheme,
            currentThemeName: constants.defaultMode,
            activeDashboardIncidentId: "",
            fromActiveDashboardTab: false,
            appSettings: {},
            isMapViewerEnabled: false,
            azureMapsKeyConfigData: {},
            appTitle: siteConfig.appTitle,
            appTitleData: {},         
            editIncidentAccessRole: "",
            editIncidentAccessRoleData: {},
            authError: ''
        }

        this.showActiveBridge = this.showActiveBridge.bind(this);
        this.updateIncidentData = this.updateIncidentData.bind(this);
        this.setState = this.setState.bind(this);
    }

    componentWillUnmount() {
        if (this._tokenRefreshInterval !== undefined) {
            window.clearInterval(this._tokenRefreshInterval);
        }
    }

    async componentDidMount() {
        try {
            const ssoToken = await new Promise<string>((resolve, reject) => {
                microsoftTeams.authentication.getAuthToken({
                    successCallback: resolve,
                    failureCallback: (err) => reject(new Error(err))
                });
            });

            const meRes = await fetch('/api/graph/me', {
                headers: { Authorization: `Bearer ${ssoToken}` }
            });
            if (!meRes.ok) throw new Error(`Graph proxy returned ${meRes.status}`);
            const currentUser = await meRes.json();

            const graph = this.createMicrosoftGraphClient();

            this.setState({
                showLoginPage: false,
                graph,
                currentUserName: currentUser.givenName || '',
                currentUserId: currentUser.id || '',
                currentUserDisplayName: currentUser.displayName || '',
                currentUserEmail: currentUser.mail || currentUser.userPrincipalName || ''
            });

            await this.initGraphToolkit();

            try {
                microsoftTeams.app.getContext().then(ctx => {
                    microsoftTeams.pages.tabs.getMruTabInstances().then((tabInfo: any) => {
                        if (ctx.channel?.id && ctx.channel?.displayName && tabInfo.teamTabs[0].tabName === constants.activeDashboardTabTitle) {
                            this.setState({
                                activeDashboardIncidentId: ctx.sharePointSite?.teamSitePath?.split("_")[1] as any,
                                fromActiveDashboardTab: true
                            });
                        }
                    });
                });

                microsoftTeams.app.getContext().then(ctx => {
                    if (ctx?.app?.locale !== "") {
                        this.setState({
                            locale: ctx.app.locale,
                            userPrincipalName: ctx.user?.userPrincipalName,
                            tenantID: ctx.user?.tenant?.id
                        });
                    } else {
                        this.setState({
                            locale: constants.defaultLocale,
                            userPrincipalName: ctx.user?.userPrincipalName,
                            tenantID: ctx.user?.tenant?.id
                        });
                    }
                    const theme = ctx.app.theme ?? constants.defaultMode;
                    this.updateTheme(theme);
                });

                microsoftTeams.pages.getConfig().then((settings) => {
                    console.log(constants.infoLogPrefix + "settings ", settings);
                    this.setState({ appSettings: settings });
                });

                microsoftTeams.app.registerOnThemeChangeHandler((theme: string) => {
                    this.updateTheme(theme);
                });

                appInsights = new ApplicationInsights({
                    config: {
                        instrumentationKey: process.env.REACT_APP_APPINSIGHTS_INSTRUMENTATIONKEY ? process.env.REACT_APP_APPINSIGHTS_INSTRUMENTATIONKEY : ''
                    }
                });
                appInsights.loadAppInsights();
            } catch (error) {
                this.setState({ locale: constants.defaultLocale });
                this.dataService.trackException(appInsights, error, constants.componentNames.EOCHomeComponent, 'ComponentDidMount', this.state.userPrincipalName);
            }

            await this.getTenantAndSiteDetails();
            await this.getConfigSettings();

            // Keep server-side OBO cache warm — zero UI side effects
            this._tokenRefreshInterval = window.setInterval(async () => {
                try {
                    const freshToken = await new Promise<string>((resolve, reject) => {
                        microsoftTeams.authentication.getAuthToken({
                            successCallback: resolve,
                            failureCallback: (err) => reject(new Error(err))
                        });
                    });
                    await fetch('/api/auth/refresh', {
                        method: 'POST',
                        headers: { Authorization: `Bearer ${freshToken}` }
                    });
                } catch (err) {
                    console.error(constants.errorLogPrefix + 'EOCHome_TokenRefresh_Failed', err);
                }
            }, 45 * 60 * 1000);

        } catch (error: any) {
            console.error(constants.errorLogPrefix + 'EOCHome_Auth_Failed', error);
            this.setState({
                locale: constants.defaultLocale,
                authError: 'Sign-in failed. Please reload the tab and try again.'
            });
        }
    }
    createMicrosoftGraphClient() {
        return Client.initWithMiddleware({
            middleware: [
                new SSOAuthMiddleware(),
                new RetryHandler(new RetryHandlerOptions()),
                new RedirectHandler(new RedirectHandlerOptions()),
                new HTTPMessageHandler()
            ],
            baseUrl: '/api/graph'
        });
    }

    //method to perform actions based on state changes
    componentDidUpdate(_prevProps: Readonly<IEOCHomeProps>, prevState: Readonly<IEOCHomeState>): void {
        if (prevState.showSuccessMessageBar !== this.state.showSuccessMessageBar && this.state.showSuccessMessageBar) {
            const classes = this.successMessagebarRef?.current?.getElementsByClassName("ms-MessageBar-content")[0].getAttribute("class");
            this.successMessagebarRef?.current?.getElementsByClassName("ms-MessageBar-content")[0].setAttribute("class", classes + " container");
        }
        if (prevState.showErrorMessageBar !== this.state.showErrorMessageBar && this.state.showErrorMessageBar) {
            const classes = this.errorMessagebarRef?.current?.getElementsByClassName("ms-MessageBar-content")[0].getAttribute("class");
            this.errorMessagebarRef?.current?.getElementsByClassName("ms-MessageBar-content")[0].setAttribute("class", classes + " container");
        }
    }

    //method to set the current theme to state variables
    updateTheme = (theme: string) => {
        switch (theme.toLocaleLowerCase()) {
            case constants.defaultMode:
                this.setState({
                    currentTeamsTheme: teamsLightTheme,
                    currentThemeName: constants.defaultMode
                });
                break;
            case constants.darkMode:
                this.setState({
                    currentTeamsTheme: teamsDarkTheme,
                    currentThemeName: constants.darkMode
                });
                break;
            case constants.contrastMode:
                this.setState({
                    currentTeamsTheme: teamsHighContrastTheme,
                    currentThemeName: constants.contrastMode
                });
                break;
        }
    };


    async initGraphToolkit() {
        Providers.globalProvider = new SimpleProvider(
            () => new Promise<string>((resolve, reject) => {
                microsoftTeams.authentication.getAuthToken({
                    successCallback: resolve,
                    failureCallback: (err) => reject(new Error(err))
                });
            }),
            async () => {},
            async () => {}
        );
        Providers.globalProvider.setState(ProviderState.SignedIn);
        Providers.globalProvider.graph = new Graph(this.state.graph as any);
    }

    // this method fetches tenant name and SharePoint site Id from the server-side config endpoint.
    // The config endpoint uses app-only credentials so this works for all users regardless of
    // SharePoint site membership.
    public async getTenantAndSiteDetails() {
        try {
            const ssoToken = await new Promise<string>((resolve, reject) => {
                microsoftTeams.authentication.getAuthToken({
                    successCallback: resolve,
                    failureCallback: (err) => reject(new Error(err))
                });
            });
            const res = await fetch('/api/config/tenant', {
                headers: { Authorization: `Bearer ${ssoToken}` }
            });
            if (!res.ok) throw new Error(`Config endpoint returned ${res.status}`);
            const { sharePointRootUrl, siteId } = await res.json();
            this.setState({
                tenantName: new URL(sharePointRootUrl).hostname,
                graphContextURL: 'https://graph.microsoft.com/v1.0/',
                siteId
            });
        } catch (error: any) {
            console.error(
                constants.errorLogPrefix + "_EOCHome_GetTenantAndSiteDetails \n",
                JSON.stringify(error)
            );
            this.dataService.trackException(appInsights, error, constants.componentNames.EOCHomeComponent, 'GetTenantAndSiteDetails', this.state.userPrincipalName);
        }
    }

    // this method connects with service layer to get the current user details
    public async getCurrentUserDetails() {
        try {
            // get the tenant name
            const currentUser = await this.dataService.getGraphData(graphConfig.meGraphEndpoint, this.state.graph);

            this.setState({
                currentUserName: currentUser.givenName,
                currentUserId: currentUser.id,
                currentUserDisplayName: currentUser.displayName,
                currentUserEmail: currentUser.mail
            })
        } catch (error: any) {
            console.error(
                constants.errorLogPrefix + "_EOCHome_GetCurrentUserDetails \n",
                JSON.stringify(error)
            );
            //log exception to AppInsights
            this.dataService.trackException(appInsights, error, constants.componentNames.EOCHomeComponent, 'GetCurrentUserDetails', this.state.userPrincipalName);
        }
    }

    //Get data from TEOC-Config sharepoint list
    private getConfigSettings = async () => {
        try {
            this.setState({
                settingsLoader: true
            });

            if (!this.state.siteId) {
                this.setState({ settingsLoader: false });
                return;
            }

            //graph endpoint to get data from TEOC-Config list
            let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.state.siteId}/lists/${siteConfig.configurationList}/items?$expand=fields&$Top=5000`;
            const configDataRecords = [constants.enableRoles, constants.azureMapsKey, constants.appTitleKey,  constants.editIncidentAccessRoleKey];
            const configData = await this.dataService.getConfigData(graphEndpoint, this.state.graph, configDataRecords);
            const rolesEnabledItem = configData.find((item: any) => item.title === constants.enableRoles);
            const rolesEnabled = rolesEnabledItem?.value === "True";
            if (rolesEnabled) {
                await this.checkUserRoleIsAdmin();
            }
            const appTitleItem = configData.filter((item: any) => item.title === constants.appTitleKey);
            const azureMapItem = configData.filter((item: any) => item.title === constants.azureMapsKey);
            const editIncidentAccessRole = configData.filter((item: any) => item.title === constants.editIncidentAccessRoleKey);
            
            if (appTitleItem.length > 0) {
                this.setState({
                    appTitle: appTitleItem[0].value,
                    appTitleData: appTitleItem[0]
                });
            }
            if (azureMapItem.length > 0) {
                this.setState({
                    isMapViewerEnabled: azureMapItem[0].value?.trim() !== "" && azureMapItem[0].value?.trim() !== undefined,
                    azureMapsKeyConfigData: azureMapItem[0]
                });
            }
            if (editIncidentAccessRole.length > 0) {
                this.setState({
                    editIncidentAccessRole: editIncidentAccessRole[0].value,
                    editIncidentAccessRoleData: editIncidentAccessRole[0]
                });
            }

            this.setState({
                isRolesEnabled: rolesEnabled,
                configRoleData: rolesEnabledItem ?? {},
                settingsLoader: false
            });
        }
        catch (error: any) {
            console.error(
                constants.errorLogPrefix + `${constants.componentNames.EOCHomeComponent}_getConfigSetting \n`,
                JSON.stringify(error)
            );
            this.setState({
                settingsLoader: false
            });
            // Log Exception
            this.dataService.trackException(appInsights, error,
                constants.componentNames.EOCHomeComponent,
                `${constants.componentNames.EOCHomeComponent}_getConfigSetting`, this.state.userPrincipalName);
        }
    }

    //Check if user's role is Admin in user roles list
    private checkUserRoleIsAdmin = async () => {
        try {
            let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.state.siteId}/lists/${siteConfig.userRolesList}/items?$expand=fields($select=Title,Role)`;

            const usersData = await this.dataService.getGraphData(graphEndpoint, this.state.graph);
            //check if the role of user is Admin
            const currentUsersdata = usersData.value.filter((item: any) => {
                return item.fields.Title?.toLowerCase().trim() === this.state.currentUserEmail?.toLowerCase().trim() && item.fields.Role === constants.adminRole
            });

            this.setState({
                isUserAdmin: currentUsersdata.length > 0
            });
        }
        catch (error) {
            console.error(
                constants.errorLogPrefix + `${constants.componentNames.EOCHomeComponent}_checkUserRoleExists \n`,
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(appInsights, error,
                constants.componentNames.EOCHomeComponent,
                `${constants.componentNames.EOCHomeComponent}_checkUserRoleExists`, this.state.userPrincipalName);
        }
    }

    // changes state to hide message bar
    private hideMessageBar = () => {
        this.setState({
            showSuccessMessageBar: false,
            showErrorMessageBar: false,
            successMessage: "",
            errorMessage: ""
        })
    }

    // changes state to show message bar
    private showMessageBar = (message: string, type: string) => {
        if (type === constants.messageBarType.success) {
            this.setState({
                showSuccessMessageBar: true,
                successMessage: message.trim()
            });
        }
        if (type === constants.messageBarType.error) {
            this.setState({
                showErrorMessageBar: true,
                errorMessage: message.trim()
            });
        }
    }

    // changes state to show create incident form
    private showNewForm = () => {
        this.setState({ showIncForm: true, selectedIncident: [] });
        this.hideMessageBar();
    }

    // changes state to show update incident form
    private showEditForm = async (incidentData: any) => {
        this.hideMessageBar();
        try {
            this.setState({
                showLoader: true
            })
            const teamGroupId = incidentData.teamWebURL.split("?")[1].split("&")[0].split("=")[1].trim();
            // check if current user is owner of the team
            await this.checkIfUserHasPermissionToEdit(teamGroupId);
            this.setState({
                showIncForm: true,
                showActiveBridge: false,
                selectedIncident: incidentData
            })

        } catch (error) {
            this.setState({
                showIncForm: true,
                showActiveBridge: false,
                selectedIncident: incidentData
            });
            console.error(
                constants.errorLogPrefix + "_EOCHome_showEditForm \n",
                JSON.stringify(error)
            );
            //log exception to AppInsights
            this.dataService.trackException(appInsights, error, constants.componentNames.EOCHomeComponent, 'showEditForm', this.state.userPrincipalName);
        }
    }

    // check if the user is owner of the team
    private checkIfUserHasPermissionToEdit = async (teamId: string): Promise<any> => {
        let isOwner = false;
        return new Promise(async (resolve, reject) => {
            try {
                const graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + teamId + graphConfig.membersGraphEndpoint;
                const existingMembers = await this.dataService.getExistingTeamMembers(graphEndpoint, this.state.graph);

                existingMembers.value.forEach((members: any) => {
                    if (members.roles.length > 0 && members.userId === this.state.currentUserId) {
                        isOwner = true;
                    }
                });

                if (isOwner) {
                    this.setState({
                        existingTeamMembers: existingMembers.value,
                        isEditMode: true,
                        showLoader: false,
                        isOwner: isOwner,
                        showNoAccessMessage: false
                    })
                }
                else {
                    this.setState({
                        existingTeamMembers: existingMembers.value,
                        isEditMode: true,
                        isOwner: isOwner,
                        showLoader: false,
                        showNoAccessMessage: true
                    })
                }
                resolve(isOwner);
            } catch (error) {
                this.setState({
                    isOwner: isOwner,
                    isEditMode: true,
                    showLoader: false,
                    showNoAccessMessage: true
                })
                reject(isOwner);
            }
        });
    }

    //set state to show Active Bridge of an incident.
    private async showActiveBridge(incidentData: any) {
        this.hideMessageBar();
        try {
            const teamGroupId = incidentData.teamWebURL.split("?")[1].split("&")[0].split("=")[1].trim();

            // check if current user is owner or member of the team
            await this.checkIfUserCanAccessBridge(teamGroupId);

            this.setState({
                showActiveBridge: true,
                selectedIncident: incidentData
            });
        } catch (error) {
            this.setState({
                showActiveBridge: true,
                selectedIncident: incidentData
            });
            console.error(
                constants.errorLogPrefix + "_EOCHome_showActiveBridge \n",
                JSON.stringify(error)
            );
            //log exception to AppInsights
            this.dataService.trackException(appInsights, error, constants.componentNames.EOCHomeComponent, 'showActiveBridge', this.state.userPrincipalName);
        }
    }

    // check if the user is owner/member of the team
    private checkIfUserCanAccessBridge = async (teamId: string): Promise<any> => {
        let isOwnerOrMember = false;
        let isOwner = false;
        return new Promise(async (resolve, reject) => {
            try {
                const graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + teamId + graphConfig.membersGraphEndpoint;
                const existingMembers = await this.dataService.getExistingTeamMembers(graphEndpoint, this.state.graph);

                existingMembers.value.forEach((members: any) => {
                    //check if the user is owner of the team
                    if (members.roles.length > 0 && members.userId === this.state.currentUserId) {
                        isOwner = true;
                    }
                    //check if the user is owner or member of the team
                    if (members.userId === this.state.currentUserId) {
                        isOwnerOrMember = true;
                    }
                });

                if (isOwner) {
                    this.setState({ isOwner: isOwner });
                }

                if (isOwnerOrMember) {
                    this.setState({
                        isOwnerOrMember: isOwnerOrMember,
                        showLoader: false,
                        showNoAccessMessage: false
                    })
                }
                else {
                    this.setState({
                        isOwnerOrMember: isOwnerOrMember,
                        showLoader: false,
                        showNoAccessMessage: true
                    })
                }
                resolve(isOwnerOrMember);
            } catch (error) {
                this.setState({
                    isOwnerOrMember: isOwnerOrMember,
                    showLoader: false,
                    showNoAccessMessage: true
                })
                reject(isOwnerOrMember);
            }
        });
    }

    private updateIncidentData = async (incidentData: IListItem) => {
        this.setState({ selectedIncident: incidentData });
    }

    // changes state to show message bar and dashboard
    private handleBackClick = (messageBarType: string) => {
        if (messageBarType === constants.messageBarType.error || messageBarType === constants.messageBarType.success) {
            this.setState({
                showIncForm: false,
                isEditMode: false,
                showAdminSettings: false,
                showIncidentHistory: false,
                showActiveBridge: false
            });
        }
        else {
            this.setState({
                showIncForm: false,
                showErrorMessageBar: false,
                showSuccessMessageBar: false,
                isEditMode: false,
                showAdminSettings: false,
                showIncidentHistory: false,
                showActiveBridge: false
            });
        }
    }

    // hide the message bar and reset the flags for unauthorized edit button click
    private hideUnauthorizedMessage = () => {
        this.setState({
            showNoAccessMessage: false,
            showIncForm: false,
            showSuccessMessageBar: false,
            showErrorMessageBar: false,
            isEditMode: false,
            showActiveBridge: false
        })
    }

    // changes state to show Admin Settings Page
    private onShowAdminSettings = () => {
        this.setState({ showAdminSettings: true });
        this.hideMessageBar();
    }

    //changes state to show Incident History of an incident.
    private onShowIncidentHistory = (incidentId: string) => {
        this.setState({ showIncidentHistory: true, incidentId: incidentId });
        this.hideMessageBar();
    }


    public render() {
        if (this.state.locale && this.state.locale !== "") {
            localeStrings.setLanguage(this.state.locale);
        }
        return (
            <FluentProvider theme={this.state.currentTeamsTheme}>
                {this.state.locale === "" ?
                    <>
                        <Loader className="loaderAlign" label={this.state.loaderMessage} size="largest" />
                    </>
                    :
                    <>
                        <EocHeader clickcallback={() => { }}
                            localeStrings={localeStrings}
                            currentUserName={this.state.currentUserName}
                            currentThemeName={this.state.currentThemeName}
                            appTitle={this.state.appTitle}                           
                        />

                        {this.state.authError &&
                            <div className='loginButton'>
                                <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>
                                    {this.state.authError}
                                </MessageBar>
                            </div>
                        }
                        {!this.state.showLoginPage && this.state.siteId !== "" &&
                            <div>
                                {this.state.showSuccessMessageBar &&
                                    <div ref={this.successMessagebarRef}>
                                        <MessageBar
                                            messageBarType={MessageBarType.success}
                                            isMultiline={false}
                                            dismissButtonAriaLabel="Close"
                                            onDismiss={() => this.setState({ showSuccessMessageBar: false, successMessage: "" })}
                                            className="message-bar"
                                            role="alert"
                                            aria-live="polite"
                                        >
                                            {this.state.successMessage}
                                        </MessageBar>
                                    </div>
                                }
                                {this.state.showErrorMessageBar &&
                                    <div ref={this.errorMessagebarRef}>
                                        <MessageBar
                                            messageBarType={MessageBarType.error}
                                            isMultiline={true}
                                            dismissButtonAriaLabel="Close"
                                            onDismiss={() => this.setState({ showErrorMessageBar: false, errorMessage: "" })}
                                            className="message-bar"
                                            role="alert"
                                            aria-live="polite"
                                        >
                                            {this.state.errorMessage}
                                        </MessageBar>
                                    </div>
                                }
                                {this.state.showLoader ?
                                    <>
                                        <Loader label={this.state.loaderMessage} size="largest" />
                                    </>
                                    : this.state.showAdminSettings ?
                                        <AdminSettings
                                            localeStrings={localeStrings}
                                            appTitle={this.state.appTitle}
                                            appTitleData={this.state.appTitleData}
                                            onBackClick={this.handleBackClick}
                                            siteId={this.state.siteId}
                                            graph={this.state.graph}
                                            appInsights={appInsights}
                                            userPrincipalName={this.state.userPrincipalName}
                                            showMessageBar={this.showMessageBar}
                                            hideMessageBar={this.hideMessageBar}
                                            currentUserDisplayName={this.state.currentUserDisplayName}
                                            currentUserEmail={this.state.currentUserEmail}
                                            isRolesEnabled={this.state.isRolesEnabled}
                                            isUserAdmin={this.state.isUserAdmin}
                                            configRoleData={this.state.configRoleData}
                                            setState={this.setState}
                                            tenantName={this.state.tenantName}
                                            siteName={siteName}
                                            currentThemeName={this.state.currentThemeName}
                                            isMapViewerEnabled={this.state.isMapViewerEnabled}
                                            azureMapsKeyConfigData={this.state.azureMapsKeyConfigData}
                                            editIncidentAccessRole={this.state.editIncidentAccessRole}
                                            editIncidentAccessRoleData={this.state.editIncidentAccessRoleData}
                                        />
                                        : this.state.showIncidentHistory ?
                                            <IncidentHistory
                                                localeStrings={localeStrings}
                                                onBackClick={this.handleBackClick}
                                                siteId={this.state.siteId}
                                                graph={this.state.graph}
                                                appInsights={appInsights}
                                                userPrincipalName={this.state.userPrincipalName}
                                                showMessageBar={this.showMessageBar}
                                                hideMessageBar={this.hideMessageBar}
                                                incidentId={this.state.incidentId}
                                                currentThemeName={this.state.currentThemeName}
                                            />
                                            : this.state.showActiveBridge ?
                                                <>
                                                    {(this.state.isOwnerOrMember && !this.state.showNoAccessMessage) ?
                                                        <ActiveBridge
                                                            localeStrings={localeStrings}
                                                            onBackClick={this.handleBackClick}
                                                            incidentData={this.state.selectedIncident}
                                                            graph={this.state.graph}
                                                            siteId={this.state.siteId}
                                                            appInsights={appInsights}
                                                            userPrincipalName={this.state.userPrincipalName}
                                                            onShowIncidentHistory={this.onShowIncidentHistory}
                                                            currentUserId={this.state.currentUserId}
                                                            updateIncidentData={this.updateIncidentData}
                                                            isOwner={this.state.isOwner}
                                                            onEditButtonClick={this.showEditForm}
                                                            graphContextURL={this.state.graphContextURL}
                                                            tenantID={this.state.tenantID}
                                                            fromActiveDashboardTab={this.state.fromActiveDashboardTab}
                                                            currentThemeName={this.state.currentThemeName}
                                                        />
                                                        :
                                                        <Dialog
                                                            confirmButton={localeStrings.okLabel}
                                                            content={localeStrings.bridgeAccessMessage}
                                                            header={localeStrings.noAccessLabel}
                                                            onConfirm={(e) => this.hideUnauthorizedMessage()}
                                                            open={this.state.showNoAccessMessage}
                                                        />
                                                    }
                                                </> :
                                                <>
                                                    {!this.state.showIncForm ?
                                                        <Dashboard
                                                            graph={this.state.graph}
                                                            tenantName={this.state.tenantName}
                                                            siteId={this.state.siteId}
                                                            onCreateTeamClick={this.showNewForm}
                                                            onEditButtonClick={this.showEditForm}
                                                            localeStrings={localeStrings}
                                                            onBackClick={this.handleBackClick}
                                                            showMessageBar={this.showMessageBar}
                                                            hideMessageBar={this.hideMessageBar}
                                                            appInsights={appInsights}
                                                            userPrincipalName={this.state.userPrincipalName}
                                                            siteName={siteName}
                                                            onShowAdminSettings={this.onShowAdminSettings}
                                                            onShowIncidentHistory={this.onShowIncidentHistory}
                                                            onShowActiveBridge={this.showActiveBridge}
                                                            isRolesEnabled={this.state.isRolesEnabled}
                                                            isUserAdmin={this.state.isUserAdmin}
                                                            settingsLoader={this.state.settingsLoader}
                                                            currentThemeName={this.state.currentThemeName}
                                                            activeDashboardIncidentId={this.state.activeDashboardIncidentId}
                                                            fromActiveDashboardTab={this.state.fromActiveDashboardTab}
                                                            isMapViewerEnabled={this.state.isMapViewerEnabled}
                                                            azureMapsKeyConfigData={this.state.azureMapsKeyConfigData}
                                                            graphBaseUrl={graphBaseURL}
                                                            currentUserId={this.state.currentUserId}
                                                        />
                                                        :
                                                        <>
                                                            {this.state.isEditMode ?
                                                                <>
                                                                    {(this.state.isOwner && !this.state.showNoAccessMessage) ?
                                                                        <IncidentDetails
                                                                            graph={this.state.graph}
                                                                            graphBaseUrl={graphBaseURL}
                                                                            graphContextURL={this.state.graphContextURL}
                                                                            siteId={this.state.siteId}
                                                                            onBackClick={this.handleBackClick}
                                                                            showMessageBar={this.showMessageBar}
                                                                            hideMessageBar={this.hideMessageBar}
                                                                            localeStrings={localeStrings}
                                                                            currentUserId={this.state.currentUserId}
                                                                            incidentData={this.state.selectedIncident}
                                                                            existingTeamMembers={this.state.existingTeamMembers}
                                                                            isEditMode={this.state.isEditMode}
                                                                            appInsights={appInsights}
                                                                            userPrincipalName={this.state.userPrincipalName}
                                                                            tenantID={this.state.tenantID}
                                                                            currentThemeName={this.state.currentThemeName}
                                                                            appSettings={this.state.appSettings}
                                                                            editIncidentAccessRole={this.state.editIncidentAccessRole}
                                                                        />
                                                                        :
                                                                        <Dialog
                                                                            confirmButton={localeStrings.okLabel}
                                                                            content={localeStrings.editIncidentAccessMessage}
                                                                            header={localeStrings.noAccessLabel}
                                                                            onConfirm={(e) => this.hideUnauthorizedMessage()}
                                                                            open={this.state.showNoAccessMessage}
                                                                        />
                                                                    }
                                                                </>
                                                                :
                                                                <IncidentDetails
                                                                    graph={this.state.graph}
                                                                    graphBaseUrl={graphBaseURL}
                                                                    graphContextURL={this.state.graphContextURL}
                                                                    siteId={this.state.siteId}
                                                                    onBackClick={this.handleBackClick}
                                                                    showMessageBar={this.showMessageBar}
                                                                    hideMessageBar={this.hideMessageBar}
                                                                    localeStrings={localeStrings}
                                                                    currentUserId={this.state.currentUserId}
                                                                    incidentData={this.state.selectedIncident}
                                                                    existingTeamMembers={this.state.existingTeamMembers}
                                                                    isEditMode={this.state.isEditMode}
                                                                    appInsights={appInsights}
                                                                    userPrincipalName={this.state.userPrincipalName}
                                                                    tenantID={this.state.tenantID}
                                                                    currentThemeName={this.state.currentThemeName}
                                                                    appSettings={this.state.appSettings}
                                                                    editIncidentAccessRole={this.state.editIncidentAccessRole}
                                                                />
                                                            }
                                                        </>
                                                    }
                                                </>
                                }
                            </div>
                        }
                    </>
                }
            </FluentProvider>
        )
    }
}
