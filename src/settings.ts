// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export type Settings = {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  subscriptionId: string;
  resourceGroup: string;
  eventGridTopic: string;
  location: string;
  portNumber: number;
};

function missingSetting(settingName: string): string {
  throw new Error(
    `${settingName} missing from environment. See README for configuration instructions.`,
  );
}

const settings: Settings = {
  clientId: process.env.CLIENT_ID || missingSetting('clientId'),
  clientSecret: process.env.CLIENT_SECRET || missingSetting('clientSecret'),
  tenantId: process.env.TENANT_ID || missingSetting('tenantId'),
  subscriptionId:
    process.env.SUBSCRIPTION_ID || missingSetting('subscriptionId'),
  resourceGroup: process.env.RESOURCE_GROUP || missingSetting('resourceGroup'),
  eventGridTopic:
    process.env.EVENT_GRID_TOPIC || missingSetting('eventGridTopic'),
  location: process.env.LOCATION || missingSetting('location'),
  portNumber: parseInt(process.env.PORT_NUMBER ?? '3000'),
};

export default settings;
