import { ChevronStartIcon, Loader } from '@fluentui/react-northstar';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { MessageBar } from '@fluentui/react/lib/MessageBar';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { Client } from "@microsoft/microsoft-graph-client";
import React from 'react';
import Col from 'react-bootstrap/esm/Col';
import Row from 'react-bootstrap/esm/Row';
import { IListItem } from '../common/CommonService';
import * as graphConfig from '../common/graphConfig';
import siteConfig from '../config/siteConfig.json';
import "../scss/ActiveBridge.module.scss";
import Bridge from './Bridge';
import Members from './Members';
import Tasks from './Tasks';
import * as constants from '../common/Constants';

export interface ActiveBridgeProps {
    onBackClick(showMessageBar: string): void;
    localeStrings: any;
    incidentData: IListItem;
    graph: Client;
    siteId: string;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    onShowIncidentHistory: Function;
    currentUserId: string;
    updateIncidentData: Function;
    onEditButtonClick: Function;
    isOwner: boolean;
    graphContextURL: string;
    tenantID: any;
    fromActiveDashboardTab: boolean;
    currentThemeName: string;
}

export interface ActiveBridgeState {
    messageType: number;
    messageText: string;
    showBridgeLoader: boolean;
    showTasksLoader: boolean;
    assessments: any[];
    assessmentsLoading: boolean;
    activityFeed: any[];
    activityLoading: boolean;
}

export default class ActiveBridge extends React.Component<ActiveBridgeProps, ActiveBridgeState> {

    constructor(props: ActiveBridgeProps) {
        super(props);

        this.state = {
            messageType: -1,
            messageText: "",
            showBridgeLoader: false,
            showTasksLoader: false,
            assessments: [],
            assessmentsLoading: true,
            activityFeed: [],
            activityLoading: true
        };

        this.updateMessagebar = this.updateMessagebar.bind(this);
    }

    async componentDidMount() {
        await Promise.all([this.loadAssessments(), this.loadActivityFeed()]);
    }

    // Load assessments from the incident team's SharePoint site
    private loadAssessments = async () => {
        try {
            const teamWebURL = this.props.incidentData.teamWebURL || '';
            if (!teamWebURL) { this.setState({ assessmentsLoading: false }); return; }

            // Derive the team group ID and then get its SharePoint site
            const groupId = teamWebURL.split("?")[1]?.split("&")[0]?.split("=")[1]?.trim();
            if (!groupId) { this.setState({ assessmentsLoading: false }); return; }

            const teamSiteEndpoint = `${graphConfig.teamsGraphEndpoint}/${groupId}/drive/root`;
            // Get the team site ID via the group
            const groupEndpoint = `${graphConfig.teamGroupsGraphEndpoint}/${groupId}/sites/root`;
            const teamSite = await this.props.graph.api(groupEndpoint).get();
            const teamSiteId = teamSite.id;

            const assessmentEndpoint = `${graphConfig.spSiteGraphEndpoint}${teamSiteId}/${graphConfig.listsGraphEndpoint}/${siteConfig.assessmentsList}/items?$expand=fields&$Top=50&$orderby=fields/Created desc`;
            const result = await this.props.graph.api(assessmentEndpoint).get();
            this.setState({ assessments: result.value || [], assessmentsLoading: false });
        } catch {
            this.setState({ assessmentsLoading: false });
        }
    };

    // Load recent activity from the incident's version history
    private loadActivityFeed = async () => {
        try {
            if (!this.props.incidentData.incidentId || !this.props.siteId) {
                this.setState({ activityLoading: false }); return;
            }
            const versionsEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incidentsList}/items/${this.props.incidentData.incidentId}/versions?$Top=10`;
            const result = await this.props.graph.api(versionsEndpoint).get();
            this.setState({ activityFeed: result.value || [], activityLoading: false });
        } catch {
            this.setState({ activityLoading: false });
        }
    };

    // Parse the role-assignment string "Role: User1|id|email, User2|id|email; Role2: ..."
    private parseRosterRows(): Array<{ role: string; users: string[]; lead: string }> {
        const raw = this.props.incidentData.roleAssignments || '';
        const leadsRaw = this.props.incidentData.roleLeads || '';

        const leadMap: Record<string, string> = {};
        leadsRaw.split(';').forEach(segment => {
            const parts = segment.split(':');
            if (parts.length >= 2) {
                const role = parts[0].trim();
                const lead = parts[1].trim().split('|')[0].trim();
                if (role && lead) leadMap[role] = lead;
            }
        });

        // Add Incident Commander as the first row
        const rows: Array<{ role: string; users: string[]; lead: string }> = [];
        const ic = this.props.incidentData.incidentCommander || '';
        if (ic) rows.push({ role: constants.incidentCommanderRoleName, users: [ic], lead: '' });

        raw.split(';').forEach(segment => {
            const parts = segment.split(':');
            if (parts.length < 2) return;
            const role = parts[0].trim();
            const userList = parts[1].trim().split(',')
                .map((u: string) => u.split('|')[0].trim())
                .filter((u: string) => u.length > 0);
            if (role && userList.length > 0) {
                rows.push({ role, users: userList, lead: leadMap[role] || '' });
            }
        });

        return rows;
    }

    private updateMessagebar = (messageType: number, message: string,
        showBridgeLoader = false, showTasksLoader = false) => {
        this.setState({
            messageType,
            messageText: message,
            showBridgeLoader,
            showTasksLoader
        });
    };

    render() {
        const isDarkOrContrastTheme = this.props.currentThemeName === constants.darkMode || this.props.currentThemeName === constants.contrastMode;
        const rosterRows = this.parseRosterRows();

        return (
            <div
                className={`active-bridge-wrapper${(this.state.showBridgeLoader || this.state.showTasksLoader) ? " disable-active-bridge" : ""}`}>
                {!this.props.fromActiveDashboardTab &&
                    <div className=".col-xs-12 .col-sm-8 .col-md-4 container" id="active-bridge-path">
                        <label>
                            <span onClick={() => this.props.onBackClick("")} className="go-back">
                                <ChevronStartIcon id="path-back-icon" />
                                <span className="back-label" title={this.props.localeStrings.back}>{this.props.localeStrings.back}</span>
                            </span> &nbsp;&nbsp;
                            <span className="right-border">|</span>
                            <span title={this.props.localeStrings.activeDashboard}>&nbsp;&nbsp;{this.props.localeStrings.activeDashboard}</span>
                        </label>
                    </div>
                }
                <div className={`active-bridge-area${isDarkOrContrastTheme ? " active-bridge-area-darkcontrast" : ""}`}>
                    <div className="container">
                        <div className='active-bridge-heading'>
                            {this.props.localeStrings.activeDashboard} - {this.props.incidentData.incidentId}
                        </div>

                        {/* ── Row 1: Team panel + Bridge/Tasks ── */}
                        <Row xl={2} lg={2} md={1}>
                            <Col xl={4} lg={4} md={12} className="members-tab-wrapper">
                                <div className='members-tab'>
                                    <div className="members-tab-heading">{this.props.localeStrings.teamLabel}</div>
                                    <Members
                                        incidentData={this.props.incidentData}
                                        graph={this.props.graph}
                                        appInsights={this.props.appInsights}
                                        userPrincipalName={this.props.userPrincipalName}
                                        localeStrings={this.props.localeStrings}
                                        isOwner={this.props.isOwner}
                                        onEditButtonClick={this.props.onEditButtonClick}
                                    />
                                </div>
                            </Col>
                            <Col xl={8} lg={8} md={12} className="bridge-tasks-wrapper">
                                <div>
                                    <div className='bridge-tab'>
                                        <div className="bridge-tab-heading">{this.props.localeStrings.bridgeLabel}</div>
                                        {this.state.showBridgeLoader && (
                                            <Loader
                                                label={this.props.localeStrings.processingLabel}
                                                size="smallest"
                                                labelPosition="start"
                                                className="bridge-spinner"
                                            />
                                        )}
                                        {this.state.messageType !== -1 &&
                                            <MessageBar
                                                messageBarType={this.state.messageType}
                                                title={this.state.messageText}
                                                className="bridge-message-bar"
                                                actions={
                                                    <IconButton
                                                        iconProps={{ iconName: "Cancel" }}
                                                        title={this.props.localeStrings.cancelIcon}
                                                        ariaLabel={this.props.localeStrings.cancelIcon}
                                                        onClick={() =>
                                                            this.setState({ messageText: "", messageType: -1 })}
                                                    />
                                                }
                                                isMultiline={false}
                                                role="status"
                                            >
                                                {this.state.messageText}
                                            </MessageBar>
                                        }
                                        <Bridge
                                            currentUserId={this.props.currentUserId}
                                            onShowIncidentHistory={this.props.onShowIncidentHistory}
                                            incidentData={this.props.incidentData}
                                            graph={this.props.graph}
                                            siteId={this.props.siteId}
                                            appInsights={this.props.appInsights}
                                            userPrincipalName={this.props.userPrincipalName}
                                            localeStrings={this.props.localeStrings}
                                            updateIncidentData={this.props.updateIncidentData}
                                            onEditButtonClick={this.props.onEditButtonClick}
                                            isOwner={this.props.isOwner}
                                            updateMessagebar={this.updateMessagebar}
                                        />
                                    </div>
                                    <div className='tasks-tab'>
                                        <div className="tasks-tab-heading">
                                            {this.props.localeStrings.tasksLabel}
                                            <span className="tasks-info-icon">
                                                <TooltipHost
                                                    content={this.props.localeStrings.tasksSectionInfoText}
                                                    calloutProps={{ gapSpace: 0 }}
                                                >
                                                    <Icon iconName="Info" aria-label={this.props.localeStrings.tasksSectionInfoText} />
                                                </TooltipHost>
                                            </span>
                                        </div>
                                        {this.state.showTasksLoader && (
                                            <Loader
                                                label={this.props.localeStrings.createPlanloaderMessage + " " + this.props.localeStrings.incidentCreationLoaderMessage}
                                                size="smallest"
                                                className="tasks-spinner"
                                            />
                                        )}
                                        <Tasks
                                            incidentData={this.props.incidentData}
                                            graph={this.props.graph}
                                            siteId={this.props.siteId}
                                            appInsights={this.props.appInsights}
                                            userPrincipalName={this.props.userPrincipalName}
                                            updateMessagebar={this.updateMessagebar}
                                            showTasksLoader={this.state.showTasksLoader}
                                            localeStrings={this.props.localeStrings}
                                            graphContextURL={this.props.graphContextURL}
                                            tenantID={this.props.tenantID}
                                        />
                                    </div>
                                </div>
                            </Col>
                        </Row>

                        {/* ── Row 2: Assessment | Team Roster | Activity Feed ── */}
                        <Row xl={3} lg={3} md={1} className="mt-3">

                            {/* Assessment panel */}
                            <Col xl={5} lg={5} md={12}>
                                <div className="irg-panel">
                                    <div className="irg-panel-heading">
                                        <Icon iconName="ClipboardList" className="irg-panel-icon" />
                                        &nbsp;{constants.Assessment}
                                    </div>
                                    {this.state.assessmentsLoading ? (
                                        <Loader size="smallest" label={this.props.localeStrings.loadingLabel} />
                                    ) : this.state.assessments.length === 0 ? (
                                        <div className="irg-panel-empty">No assessments submitted yet.</div>
                                    ) : (
                                        <table className="irg-assessment-table" aria-label="Assessments">
                                            <thead>
                                                <tr>
                                                    <th>Contact</th>
                                                    <th>Location</th>
                                                    <th>Status</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {this.state.assessments.slice(0, 8).map((item: any, idx: number) => (
                                                    <tr key={idx}>
                                                        <td>{item.fields?.["Contact Name"] || item.fields?.Title || "—"}</td>
                                                        <td>{item.fields?.Location?.DisplayName || "—"}</td>
                                                        <td>{item.fields?.Status || "—"}</td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    )}
                                </div>
                            </Col>

                            {/* Team Roster panel */}
                            <Col xl={4} lg={4} md={12}>
                                <div className="irg-panel">
                                    <div className="irg-panel-heading">
                                        <Icon iconName="People" className="irg-panel-icon" />
                                        &nbsp;Team Roster
                                    </div>
                                    {rosterRows.length === 0 ? (
                                        <div className="irg-panel-empty">No roles assigned.</div>
                                    ) : (
                                        <table className="irg-roster-table" aria-label="Team Roster">
                                            <thead>
                                                <tr>
                                                    <th>Role</th>
                                                    <th>Assigned</th>
                                                    <th>Lead</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {rosterRows.map((row, idx) => (
                                                    <tr key={idx}>
                                                        <td><strong>{row.role}</strong></td>
                                                        <td>{row.users.join(', ')}</td>
                                                        <td>{row.lead || '—'}</td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    )}
                                </div>
                            </Col>

                            {/* Activity Feed panel */}
                            <Col xl={3} lg={3} md={12}>
                                <div className="irg-panel">
                                    <div className="irg-panel-heading">
                                        <Icon iconName="ActivityFeed" className="irg-panel-icon" />
                                        &nbsp;Activity Feed
                                    </div>
                                    {this.state.activityLoading ? (
                                        <Loader size="smallest" label={this.props.localeStrings.loadingLabel} />
                                    ) : this.state.activityFeed.length === 0 ? (
                                        <div className="irg-panel-empty">No recent activity.</div>
                                    ) : (
                                        <ul className="irg-activity-list">
                                            {this.state.activityFeed.slice(0, 8).map((ver: any, idx: number) => (
                                                <li key={idx} className="irg-activity-item">
                                                    <span className="irg-activity-version">v{ver.id}</span>
                                                    <span className="irg-activity-by">
                                                        {ver.lastModifiedBy?.user?.displayName || "—"}
                                                    </span>
                                                    <span className="irg-activity-date">
                                                        {ver.lastModifiedDateTime
                                                            ? new Date(ver.lastModifiedDateTime).toLocaleString()
                                                            : "—"}
                                                    </span>
                                                    {ver.fields?.ReasonForUpdate && (
                                                        <span className="irg-activity-reason">
                                                            {ver.fields.ReasonForUpdate}
                                                        </span>
                                                    )}
                                                </li>
                                            ))}
                                        </ul>
                                    )}
                                </div>
                            </Col>

                        </Row>
                    </div>
                </div>
            </div>
        );
    }
}
