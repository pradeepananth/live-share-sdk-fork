/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { ITokenProvider, ITokenResponse } from "@fluidframework/routerlicious-driver";
import { TeamsLiveAPI } from './TeamsLiveAPI';

/**
 * Token Provider implementation for connecting to Cowatch Cloud endpoint
 */
export class TeamsFluidTokenProvider implements ITokenProvider {
    private frsToken?: string;
    private _documentId?: string;
    private _tenantId?: string;
    /**
     * Creates a new instance using configuration parameters.
     */
    constructor(private readonly api: TeamsLiveAPI) { }

    public async fetchOrdererToken(tenantId: string, documentId?: string, refresh?: boolean): Promise<ITokenResponse> {
        const tokenResponse = await this.fetchFluidToken(tenantId, documentId, refresh);
        return tokenResponse;
    }

    public async fetchStorageToken(tenantId: string, documentId?: string, refresh?: boolean): Promise<ITokenResponse> {
        const tokenResponse = await this.fetchFluidToken(tenantId, documentId, refresh);
        return tokenResponse;
    }

    private async fetchFluidToken(tenantId: string, documentId?: string, refresh?: boolean): Promise<ITokenResponse> {
        let fromCache: boolean;
        if (!this.frsToken
            || refresh
            || this._tenantId !== tenantId
            || this._documentId !== documentId) {
            this.frsToken = await this.api.getFluidToken(documentId);
            fromCache = false;
        } else {
            fromCache = true;
        }
        this._tenantId = tenantId;
        if (documentId) {
            this._documentId = documentId;
        } 
        return {
            jwt: this.frsToken,
            fromCache,
        };
    }
}
