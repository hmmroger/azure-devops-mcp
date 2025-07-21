// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as azdev from "azure-devops-node-api";
import { AccessToken, DefaultAzureCredential } from "@azure/identity";
import { UserAgentComposer } from "./useragent.js";

/**
 * Manages Azure DevOps authentication and client creation
 */
export class AzureDevOpsClientManager {
  private readonly userAgentComposer: UserAgentComposer;
  private readonly productVersion: string;

  constructor(
    private orgName: string,
    userAgentComposer: UserAgentComposer,
    productVersion: string
  ) {
    this.userAgentComposer = userAgentComposer;
    this.productVersion = productVersion;
  }

  public async getOrgName(): Promise<string> {
    return this.orgName;
  }

  public async setOrgName(name: string): Promise<void> {
    this.orgName = name;
  }

  /**
   * Gets an Azure DevOps access token using DefaultAzureCredential
   */
  public async getToken(): Promise<AccessToken> {
    if (process.env.ADO_MCP_AZURE_TOKEN_CREDENTIALS) {
      process.env.AZURE_TOKEN_CREDENTIALS = process.env.ADO_MCP_AZURE_TOKEN_CREDENTIALS;
    } else {
      process.env.AZURE_TOKEN_CREDENTIALS = "dev";
    }
    const credential = new DefaultAzureCredential(); // CodeQL [SM05138] resolved by explicitly setting AZURE_TOKEN_CREDENTIALS
    const token = await credential.getToken("499b84ac-1321-427f-aa17-267ca6975798/.default");
    return token;
  }

  /**
   * Creates and returns an Azure DevOps WebApi client
   */
  public async getClient(): Promise<azdev.WebApi> {
    const token = await this.getToken();
    const authHandler = azdev.getBearerHandler(token.token);
    const orgName = await this.getOrgName();
    const orgUrl = `https://dev.azure.com/${orgName}`;
    const connection = new azdev.WebApi(orgUrl, authHandler, undefined, {
      productName: "AzureDevOps.MCP",
      productVersion: this.productVersion,
      userAgent: this.userAgentComposer.userAgent,
    });
    return connection;
  }

  /**
   * Returns a function that creates an Azure DevOps WebApi client
   * This is for compatibility with existing tool configuration patterns
   */
  public getClientFactory(): () => Promise<azdev.WebApi> {
    return () => this.getClient();
  }

  /**
   * Returns the user agent string
   */
  public getUserAgent(): string {
    return this.userAgentComposer.userAgent;
  }
}
