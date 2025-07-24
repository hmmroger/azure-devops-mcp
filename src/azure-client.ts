// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as azdev from "azure-devops-node-api";
import { AccessToken, DefaultAzureCredential } from "@azure/identity";
import { UserAgentComposer } from "./useragent.js";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

/**
 * Manages Azure DevOps authentication and client creation
 */
export class AzureDevOpsClientManager {
  private readonly userAgentComposer: UserAgentComposer;
  private readonly productVersion: string;
  private readonly configDir: string;
  private readonly configFile?: string;

  constructor(
    userAgentComposer: UserAgentComposer,
    productVersion: string,
    private orgName?: string
  ) {
    this.userAgentComposer = userAgentComposer;
    this.productVersion = productVersion;

    // Set up config directory and file paths
    this.configDir = path.join(os.homedir(), ".azure-devops-mcp");
    this.configFile = path.join(this.configDir, "config.json");
    try {
      if (!fs.existsSync(this.configDir)) {
        fs.mkdirSync(this.configDir, { recursive: true });
      }
    } catch (error) {
      this.configFile = undefined;
      console.warn("Organization setting will not be persisted.", error);
    }

    // Load org name from file if not provided and file exists
    if (!this.orgName) {
      this.orgName = this.loadOrgNameFromFile();
    }
  }

  public async getOrgName(): Promise<string | undefined> {
    return this.orgName;
  }

  public async setOrgName(name: string): Promise<void> {
    this.orgName = name;
    this.saveOrgNameToFile(name);
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
    if (!orgName) {
      throw new Error("Azure DevOps organization name not defined. Confirm with user on the name then use 'change-azure-devops-org' to specify the organization.");
    }
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

  /**
   * Loads the organization name from the config file
   */
  private loadOrgNameFromFile(): string | undefined {
    try {
      if (!this.configFile) {
        return undefined;
      }

      if (fs.existsSync(this.configFile)) {
        const configData = fs.readFileSync(this.configFile, "utf8");
        const config = JSON.parse(configData);
        return config.orgName;
      }
    } catch (error) {
      // Silently ignore errors when loading config
      console.warn("Failed to load organization name from config file:", error);
    }
    return undefined;
  }

  /**
   * Saves the organization name to the config file
   */
  private saveOrgNameToFile(orgName: string): void {
    try {
      if (!this.configFile) {
        return;
      }

      const config = { orgName };
      fs.writeFileSync(this.configFile, JSON.stringify(config, null, 2), "utf8");
    } catch (error) {
      console.warn("Failed to save organization name to config file:", error);
    }
  }
}
