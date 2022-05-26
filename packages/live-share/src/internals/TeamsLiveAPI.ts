/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { UserMeetingRole } from '../interfaces';
import * as microsoftTeams from '@microsoft/teams-js';
import fetch from "cross-fetch";
import { ITokenPayload, ICollabSpaceInfo, IRegisteredUsersRoles, IVerifiedUserRoles, ITokenResponse } from './contracts';
import { parseJwt } from './utils';
import { ContainerState } from './TestTeamsClientApi';

// Should be replaced with app ID lookup on the Teams Client side.
const APP_ID = `current-app-id`;

export interface IFluidContainerInfo {
    containerState: ContainerState;

    /**
     * ID of the Fluid Container to connect to if needed.
     */
    containerId: string | undefined;

    /**
     * If true, the client should attempt to create the container and then save the created container's
     *  ID using `setFluidContainerId()`.
     * 
     * If false, and 'id' is undefined, the client should wait for the specified `retryAfter` interval
     * and then call `getFluidContainerId()` again.
     */
    shouldCreate: boolean;

    /**
     * Number of milliseconds to wait before calling `getFluidContainerId()` again. Will be 0 if `id`
     * is defined or `shouldCreate` is true.
     */
    retryAfter: number;
}

export interface INtpTimeInfo {
    /**
     * The server's reported time as a string.
     */
    ntpTime: string;

    /**
     * The server's reported time as a UTC timestamp.
     */
    ntpTimeInUTC: number;
}

export interface IFluidTenantInfo {
    tenantId: string;
    ordererEndpoint: string;
    storageEndpoint: string;
}

export class TeamsLiveAPI {
    private readonly _apiHost: string;
    private readonly _isTesting: boolean;
    private _authToken?: string;
    private _context?: microsoftTeams.Context;
    private _tokenPayload?: ITokenPayload;

    constructor(apiHost = "https://dev.teamsgraph.teams.microsoft.net/", isTesting = false) {
        this._isTesting = isTesting;
        this._apiHost = apiHost;
        if (!this._apiHost.endsWith('/')) {
            this._apiHost += '/';
        }
    }

    /**
     * Returns the fluid access token for mapped container Id
     * @param containerId Fluid's container Id for the request. Undefined for new containers.
     * @returns token for connecting to Fluid's session.
     */
    public async getFluidToken(containerId?: string): Promise<string> {
        const tokenPayload = await this.getTokenPayload();
        const context = await this.getContext();
        const res = await this.fetch(`${this._apiHost}livesync/v1/fluid/token/get`, {
            method: 'POST',
            body: JSON.stringify({
                "appId": APP_ID,
                "originUri": tokenPayload.aud,
                "teamsContextType": 'meetingId',
                "teamsContext": {
                    "meetingId": context.meetingId
                },
                "containerId": containerId || null,
                "userId": tokenPayload.oid,
                "userName": 'User'
            })
        });

        // Check for request error
        if (res.status >= 400) {
            throw new Error(`TeamsLiveAPI: request for fluid token returned a status of ${res.status}`);
        }

        const response = await res.json() as ITokenResponse;
        if (!response.token) {
            throw new Error(`TeamsLiveAPI: no token in token providers response`);
        }

        return response.token;
    }

    /**
     * Returns the Fluid Tenant connection info for user's current context.
     */
    public async getFluidTenantInfo(): Promise<IFluidTenantInfo> {
        const tokenPayload = await this.getTokenPayload();
        const context = await this.getContext();
        const res = await this.fetch(`${this._apiHost}livesync/v1/fluid/tenantInfo/get`, {
            method: 'POST',
            body: JSON.stringify({
                "appId": APP_ID,
                "originUri": tokenPayload.aud,
                "teamsContextType": 'MeetingId',
                "teamsContext": {
                    "meetingId": context.meetingId
                }
            })
        });

        // Check for request error
        if (res.status >= 400) {
            throw new Error(`TeamsLiveAPI: request for fluid tenant info returned a status of ${res.status}`);
        }

        const info = await res.json() as ICollabSpaceInfo;
        if (!info?.broadcaster?.frsTenantInfo) {
            throw new Error(`TeamsLiveAPI: no FRS info in collabSpaceInfo response`);
        }

        return info.broadcaster.frsTenantInfo;
    }

    /**
     * Returns the ID of the fluid container associated with the user's current context.
     */
    public async getFluidContainerId(): Promise<IFluidContainerInfo> {
        const tokenPayload = await this.getTokenPayload();
        const context = await this.getContext();
        const res = await this.fetch(`${this._apiHost}livesync/v1/fluid/container/get`, {
            method: 'POST',
            body: JSON.stringify({
                "appId": APP_ID,
                "originUri": tokenPayload.aud,
                "teamsContextType": 'MeetingId',
                "teamsContext": {
                    "meetingId": context.meetingId
                }
            })
        });

        // Check for request error
        if (res.status >= 400) {
            throw new Error(`TeamsLiveAPI: request for fluid container id returned a status of ${res.status}`);
        }

        return await res.json();
    }

    /**
     * Sets the ID of the fluid container associated with the current context.
     * 
     * @remarks
     * If this returns false, the client should delete the container they created and then call 
     * `getFluidContainerId()` to get the ID of the container being used. 
     * @param containerId ID of the fluid container the client created.
     * @returns True if the client created the container that's being used.
     */
    public async setFluidContainerId(containerId: string): Promise<boolean> {
        const tokenPayload = await this.getTokenPayload();
        const context = await this.getContext();
        const res = await this.fetch(`${this._apiHost}livesync/v1/fluid/container/set`, {
            method: 'POST',
            body: JSON.stringify({
                "appId": APP_ID,
                "originUri": tokenPayload.aud,
                "teamsContextType": 'MeetingId',
                "teamsContext": {
                    "meetingId": context.meetingId
                },
                "containerId": containerId
            })
        });

        // Check for request error
        if (res.status >= 400) {
            throw new Error(`TeamsLiveAPI: request to set fluid container id returned a status of ${res.status}`);
        }

        // Check for success
        const response = await res.json() as IFluidContainerInfo;
        return response.containerState == ContainerState.added;
    }

    /**
     * Returns the shared clock server's current time.
     */
    public async getNtpTime(): Promise<INtpTimeInfo> {
        const url = `${this._apiHost}livesync/v1/getNTPTime`;

        const res = await this.fetch(url, { method: 'GET' });

        // Check for request error
        if (res.status >= 400) {
            throw new Error(`TeamsLiveAPI: request for current time returned a status of ${res.status}`);
        }

        // Deserialize response
        return await res.json();
    }

    /**
     * Associates the fluid clientId with a set of user roles.
     * @param clientId The ID for the current user's Fluid client. Changes on reconnects.
     * @returns The roles for the current user.
     */
    public async registerClientId(clientId: string): Promise<UserMeetingRole[]> {
        const tokenPayload = await this.getTokenPayload();
        const context = await this.getContext();
        const res = await this.fetch(`${this._apiHost}livesync/v1/clientRoles/register`, {
            method: 'POST',
            body: JSON.stringify({
                "appId": APP_ID,
                "originUri": tokenPayload.aud,
                "teamsContextType": 'MeetingId',
                "teamsContext": {
                    "meetingId": context.meetingId
                },
                "clientId": clientId
            })
        });

        // Check for request error
        if (res.status >= 400) {
            throw new Error(`TeamsLiveAPI: registering client roles returned a status of ${res.status}`);
        }

        const response = await res.json() as IRegisteredUsersRoles;
        return response.userRoles;
    }

    /**
     * Verifies that a client has the roles they are claiming to have.  
     * @param clientId The Client ID the message was received from.
     * @param roles List of roles the client is claiming to have.
     * @returns true if the client has the claimed roles.
     */
    public async verifyClientRoles(clientId: string, roles: UserMeetingRole[]): Promise<boolean> {
        const tokenPayload = await this.getTokenPayload();
        const context = await this.getContext();
        const res = await this.fetch(`${this._apiHost}livesync/v1/clientRoles/verify`, {
            method: 'POST',
            body: JSON.stringify({
                "appId": APP_ID,
                "originUri": tokenPayload.aud,
                "teamsContextType": 'MeetingId',
                "teamsContext": {
                    "meetingId": context.meetingId
                },
                "clientId": clientId,
                "userRoles": roles
            })
        });

        // Check for request error
        if (res.status >= 400) {
            throw new Error(`TeamsLiveAPI: verifying client roles returned a status of ${res.status}`);
        }

        const response = await res.json() as IVerifiedUserRoles;
        return response.rolesAccepted;
    }

    /**
     * Verifies that a client has the roles they are claiming to have.  
     * @param clientId The Client ID the message was received from.
     * @returns List of roles the client is claiming to have.
     */
    public async getClientRoles(clientId: string): Promise<UserMeetingRole[] | undefined> {
        const tokenPayload = await this.getTokenPayload();
        const context = await this.getContext();
        const res = await this.fetch(`${this._apiHost}livesync/v1/clientRoles/get`, {
            method: 'POST',
            body: JSON.stringify({
                "appId": APP_ID,
                "originUri": tokenPayload.aud,
                "teamsContextType": 'MeetingId',
                "teamsContext": {
                    "meetingId": context.meetingId
                },
                "clientId": clientId
            })
        });

        // Check for request error
        if (res.status == 404) {
            // Client isn't registered yet.
            return undefined;
        } else if (res.status >= 400) {
            throw new Error(`TeamsLiveAPI: verifying client roles returned a status of ${res.status}`);
        } else {
            const response = await res.json();
            if (response?.userRoles) {
                return (response as IRegisteredUsersRoles).userRoles;
            }
            return undefined;
        }
    }

    private async fetch(url: string, options: RequestInit): Promise<Response> {
        // Add authorization header
        if (!this._isTesting) {
            const token = await this.getAuthToken();
            options.headers = {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            };
        } else {
            options.headers = {
                'Content-Type': 'application/json'
            };
        }

        // Enable CORS
        options.mode = 'cors';

        // Make request
        return await fetch(url, options);
    }

    private async getTokenPayload(): Promise<ITokenPayload> {
        if (!this._tokenPayload) {
            if (!this._isTesting) {
                const token = await this.getAuthToken();
                this._tokenPayload = parseJwt(token);
            } else {
                this._tokenPayload = {
                    oid: 'test-user',
                    aud: 'https://example.com',
                    tid: ''
                };
            }
        }

        return this._tokenPayload!;
    }

    private getAuthToken(): Promise<string> {
        if (this._authToken) {
            return Promise.resolve(this._authToken);
        } else {
            return new Promise((resolve, reject) => {
                microsoftTeams.authentication.getAuthToken({
                    successCallback: (result) => {
                        this._authToken = result;
                        resolve(result);
                    },
                    failureCallback: (error) => reject(error),
                });
            });

        }
    }

    private getContext(): Promise<microsoftTeams.Context> {
        if (this._context) {
            return Promise.resolve(this._context);
        } else {
            return new Promise((resolve) => {
                microsoftTeams.getContext((result) => {
                    this._context = result;
                    resolve(result);
                });
            });
        }
    }
}