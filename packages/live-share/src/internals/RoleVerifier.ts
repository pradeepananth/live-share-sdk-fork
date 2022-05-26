/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the Microsoft Live Share SDK License.
 */

import { IRoleVerifier, UserMeetingRole } from '../interfaces';
import {  waitForResult } from './utils';
import { RequestCache } from './RequestCache';
import { TeamsLiveAPI } from './TeamsLiveAPI';

const EXPONENTIAL_BACKOFF_SCHEDULE = [100, 200, 200, 400, 600];
const CACHE_LIFETIME = 60 * 60 * 1000;


/**
 * @hidden
 */
export class RoleVerifier implements IRoleVerifier {
    private readonly _registerRequestCache: RequestCache<UserMeetingRole[]> = new RequestCache(CACHE_LIFETIME);
    private readonly _getRequestCache: RequestCache<UserMeetingRole[]> = new RequestCache(CACHE_LIFETIME);
    constructor(private readonly api: TeamsLiveAPI) { }
    
    public async registerClientId(clientId: string): Promise<UserMeetingRole[]> {
        return this._registerRequestCache.cacheRequest(clientId, () => {
            return waitForResult(async () => {
                return await this.api.registerClientId(clientId);
            }, (result) => {
                return Array.isArray(result);
            }, () => {
                return new Error(`RoleVerifier: timed out registering local client ID`);
            }, EXPONENTIAL_BACKOFF_SCHEDULE);
        });
    }

    public async getClientRoles(clientId: string): Promise<UserMeetingRole[]> {
        if (!clientId) {
            throw new Error(`RoleVerifier: called getCLientRoles() without a clientId`);
        }

        return this._getRequestCache.cacheRequest(clientId, () => {
            return waitForResult(async () => {
                return await this.api.getClientRoles(clientId);
            }, (result) => {
                return Array.isArray(result);
            }, () => {
                return new Error(`RoleVerifier: timed out getting roles for a remote client ID`);
            }, EXPONENTIAL_BACKOFF_SCHEDULE);
        });
    }

    public async verifyRolesAllowed(clientId: string, allowedRoles: UserMeetingRole[]): Promise<boolean> {
        if (!clientId) {
            throw new Error(`RoleVerifier: called verifyRolesAllowed() without a clientId`);
        }

        if (Array.isArray(allowedRoles) && allowedRoles.length > 0) {
            const roles = await this.getClientRoles(clientId);
            for (let i = 0; i < allowedRoles.length; i++) {
                const role = allowedRoles[i];
                if (roles.indexOf(role) >= 0) {
                    return true;
                }
            }

            return false;
        }

        return true;
    }
}
