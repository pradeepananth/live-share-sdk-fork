/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

 import { TeamsLiveAPI, TeamsFluidTokenProvider, SharedClock, RoleVerifier, IFluidContainerInfo, ContainerState } from './internals';
 import { AzureClient, AzureConnectionConfig, AzureContainerServices, ITelemetryBaseLogger, LOCAL_MODE_TENANT_ID } from "@fluidframework/azure-client";
 import { ContainerSchema, IFluidContainer } from "@fluidframework/fluid-static";
 import { EphemeralEvent } from "./EphemeralEvent";
 
 /**
  * Information from a tabs `microsoftTeams.Context` object, needed to connect to  the fluid 
  * container for a meeting.
  */
 export interface IMeetingContext {
     /**
      * Meeting Id used by tab when running in meeting context
      */
     meetingId?: string;
 }
 
 export interface TeamsFluidClientProps {
 
     /**
      * Optional. Custom endpoint to use for connecting to Teams Collaboration Service. 
      */
     readonly apiHost?: string;
 
     /**
      * Optional. Configuration to use when connecting to a custom Azure Fluid Relay instance.
      */
     readonly connection?: AzureConnectionConfig,
      
      /**
       * Optional. A logger instance to receive diagnostic messages.
       */
     readonly logger?: ITelemetryBaseLogger,
 
     /**
      * Optional. Function to lookup the ID of the container to use for local testing. 
      * 
      * @remarks
      * The default implementation attempts to retrieve the containerId from the `window.location.hash`. 
      * 
      * If the function returns 'undefined' a new container will be created.
      */
     readonly getLocalTestContainerId?: () => string|undefined;
 
     /**
      * Optional. Function to save the ID of a newly created local test container. 
      * 
      * @remarks
      * The default implementation updates `window.location.hash` with the ID of the newly created 
      * container. 
      */
      readonly setLocalTestContainerId?: (containerId: string) => void;
 
     /**
      * Optional. If true the client is being used in a local testing mode.
      */
     readonly isTesting?: boolean;
 }
 
 
 /**
  * Client used to connect to fluid containers within a Microsoft Teams context.
  */
 export class TeamsFluidClient {
     private readonly _props: TeamsFluidClientProps;
     private readonly _api: TeamsLiveAPI;
     private _clock?: SharedClock;
     private _roleVerifier?: RoleVerifier;
 
     /**
      * Creates a new `TeamsFluidClient` instance.
      * @param props Configuration options for the client. 
      */
     constructor(props?: TeamsFluidClientProps) {
         // Save props
         this._props = Object.assign({
             apiHost: props?.apiHost ?? 'https://teams.microsoft.com/api/platform/',
             isTesting: props?.isTesting ?? true
         } as TeamsFluidClientProps, props);
 
         // Create API and setup role service
         this._api = new TeamsLiveAPI(this._props.isTesting ? 'https://fluidrelayapiservertest.azurewebsites.net/' : this._props.apiHost, this._props.isTesting);
     }
 
     /**
      * If true the client is configured to use a local test server.
      */
     public get isTesting(): boolean {
         return this._props.connection?.tenantId == LOCAL_MODE_TENANT_ID;
     }
 
     /**
      * Number of times the client should attempt to get the ID of the container to join for the 
      * current context.
      */
     public maxContainerLookupTries = 3;
 
     /**
      * Connects to the fluid container for the current teams context.
      * 
      * @remarks
      * The first client joining the container will create the container resulting in the 
      * `onContainerFirstCreated` callback being called. This callback can be used to set the initial
      * state of of the containers object prior to the container being attached.
      * @param fluidContainerSchema Fluid objects to create.
      * @param onContainerFirstCreated Optional. Callback that's called when the container is first created.
      * @returns The fluid `container` and `services` objects to use along with a `created` flag that if true means the container had to be created.
      */
     public async joinContainer(fluidContainerSchema: ContainerSchema, onContainerFirstCreated?: (container: IFluidContainer) => void): Promise<{
         container: IFluidContainer;
         services: AzureContainerServices;
         created: boolean;
     }> {
         performance.mark(`TeamsSync: join container`);
         try {
             // Configure role verifier and timestamp provider
             const pRoleVerifier = this.initializeRoleVerifier();
             const pTimestampProvider = this.initializeTimestampProvider();
 
             // Initialize FRS connection config 
             let config: AzureConnectionConfig | undefined = this._props.connection;
             if (!config) {
                 const frsTenantInfo = await this._api.getFluidTenantInfo();
                 config = {
                     tenantId: frsTenantInfo.tenantId,
                     tokenProvider: new TeamsFluidTokenProvider(this._api),
                     orderer: frsTenantInfo.ordererEndpoint,
                     storage: frsTenantInfo.storageEndpoint,
                 };
             }  
 
             // Create FRS client
             const client = new AzureClient({
                 connection: config,
                 logger: this._props.logger
             });
 
             // Create container on first access
             const pContainer = this.getOrCreateContainer(client, fluidContainerSchema, 0, onContainerFirstCreated);
 
             // Wait in parallel for everything to finish initializing.
             const result = await Promise.all([pContainer, pRoleVerifier, pTimestampProvider]);
 
             performance.mark(`TeamsSync: container connecting`);
 
             // Wait for containers socket to connect
             let connected = false;
             const { container, services } = result[0];
             container.on('connected', async () => {
                 if (!connected) {
                     connected = true;
                     performance.measure(`TeamsSync: container connected`, `TeamsSync: container connecting`);
                 }
 
                 // Register any new clientId's
                 // - registerClientId() will only register a client on first use
                 if (this._roleVerifier) {
                     const connections = services.audience.getMyself()?.connections ?? [];
                     for (let i = 0; i < connections.length; i++) {
                         const clientId = connections[i].id;
                         await this._roleVerifier?.registerClientId(clientId);
                     }
                 }
             });
 
             return result[0];
         } finally {
             performance.measure(`TeamsSync: container joined`, `TeamsSync: join container`);
         }
     }
 
     protected initializeRoleVerifier(): Promise<void> {
         if (!this._roleVerifier && !this.isTesting) {
             this._roleVerifier = new RoleVerifier(this._api);
             
             // Register role verifier as current verifier for events
             EphemeralEvent.setRoleVerifier(this._roleVerifier);
         } 
         
         return Promise.resolve();
     }
 
     protected initializeTimestampProvider(): Promise<void> {
         if (!this._clock && !this.isTesting) {
             this._clock = new SharedClock(this._api);
 
             // Register clock as current timestamp provider for events
             EphemeralEvent.setTimestampProvider(this._clock);
 
             // Start the clock
             return this._clock.start();
         } else {
             return Promise.resolve();
         }
     }
 
     protected getLocalTestContainerId(): string|undefined {
         if (this._props.getLocalTestContainerId) {
             return this._props.getLocalTestContainerId();
         } else if (window.location.hash) {
             return window.location.hash.substring(1);
         } else {
             return undefined;
         }            
     } 
 
     protected setLocalTestContainerId(containerId: string): void {
         if (this._props.setLocalTestContainerId) {
             this._props.setLocalTestContainerId(containerId);
         } else {
             window.location.hash = containerId;
         }            
     } 
 
     private async getOrCreateContainer(client: AzureClient, fluidContainerSchema: ContainerSchema, tries: number, onInitializeContainer?: (container: IFluidContainer) => void): Promise<{
         container: IFluidContainer;
         services: AzureContainerServices;
         created: boolean;
     }> {
         // Get container ID mapping
         const containerInfo = await this.getFluidContainerId();
 
         // Create container on first access
         if (containerInfo.shouldCreate) {
             return await this.createNewContainer(client, fluidContainerSchema, tries, onInitializeContainer);
         } else if (containerInfo.containerId) {
             return {created: false, ...await client.getContainer(containerInfo.containerId, fluidContainerSchema)};
         } else if (tries < this.maxContainerLookupTries && containerInfo.retryAfter > 0) {
             await this.wait(containerInfo.retryAfter);
             return await this.getOrCreateContainer(client, fluidContainerSchema, tries + 1, onInitializeContainer);
         } else {
             throw new Error(`TeamsFluidClient: timed out attempting to create or get container for current context.`);
         }
     }
 
     private async createNewContainer(client: AzureClient, fluidContainerSchema: ContainerSchema, tries: number, onInitializeContainer?: (container: IFluidContainer) => void): Promise<{
         container: IFluidContainer;
         services: AzureContainerServices;
         created: boolean;
     }> {
         // Create and initialize container
         const { container, services } = await client.createContainer(fluidContainerSchema);
         if (onInitializeContainer) {
             onInitializeContainer(container)
         }
 
         // Attach container to service
         const newContainerId = await container.attach();
 
         // Attempt to update container mapping
         if (!await this.setFluidContainerId(newContainerId)) {
             // Delete created container
             container.dispose();
 
             // Get mapped container ID
             return this.getOrCreateContainer(client, fluidContainerSchema, tries + 1, onInitializeContainer);
         } else {
             return {container, services, created: true};
         }
     }
 
     private getFluidContainerId(): Promise<IFluidContainerInfo> {
         if (!this.isTesting) {
             return this._api.getFluidContainerId();
         } else {
             const containerId = this.getLocalTestContainerId();
             return Promise.resolve({
                 containerState: containerId ? ContainerState.alreadyExists : ContainerState.notFound,
                 shouldCreate: !containerId,
                 containerId: containerId,
                 retryAfter: 0
             });
         }
     }
 
     private setFluidContainerId(containerId: string): Promise<boolean> {
         if (!this.isTesting) {
             return this._api.setFluidContainerId(containerId);
         } else {
             this.setLocalTestContainerId(containerId);
             return Promise.resolve(true);
         }
     }
 
     private wait(delay: number): Promise<void> {
         return new Promise((resolve) => {
             setTimeout(() => resolve(), delay);
         });
     }
 }