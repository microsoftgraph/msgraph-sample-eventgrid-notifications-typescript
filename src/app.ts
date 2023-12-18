// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import 'dotenv-flow/config';
import * as express from 'express';
import Router from 'express-promise-router';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { Client, PageCollection } from '@microsoft/microsoft-graph-client';
import settings from './settings';
import {
  handleNotificationAsync,
  handleValidationRequest,
} from './handlers/notificationEventHandler';
import { Subscription } from '@microsoft/microsoft-graph-types';

// Setup the Graph client
const credential = new ClientSecretCredential(
  settings.tenantId,
  settings.clientId,
  settings.clientSecret,
);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default'],
});

const app = express();
app.use(express.json({ type: '*/*' }));
const router = Router();
router.options('/', handleValidationRequest);
router.post('/', handleNotificationAsync);
app.use(router);

app.locals.graphClient = Client.initWithMiddleware({
  authProvider: authProvider,
});

app.listen(settings.portNumber, () => {
  console.log(`Server running at http://localhost:${settings.portNumber}`);
});

ensureSubscription();

// Ensure Graph subscription exists
async function ensureSubscription() {
  const subscriptions: PageCollection = await app.locals.graphClient
    .api('/subscriptions')
    .get();

  if (subscriptions.value.length > 0) {
    console.log('Subscription already exists');
  } else {
    console.log('No existing subscription found');

    const eventGridUrl =
      `EventGrid:?azuresubscriptionid=${settings.subscriptionId}` +
      `&resourcegroup=${settings.resourceGroup}` +
      `&partnertopic=${settings.eventGridTopic}` +
      `&location=${settings.location}`;

    const newSubscription: Subscription = {
      changeType: 'updated,deleted,created',
      resource: 'users',
      clientState: 'SomeSecretValue',
      notificationUrl: eventGridUrl,
      lifecycleNotificationUrl: eventGridUrl,
      // Set expiration to one hour for testing purposes
      expirationDateTime: new Date(Date.now() + 3600000).toISOString(),
    };

    try {
      const createdSubscription: Subscription = await app.locals.graphClient
        .api('/subscriptions')
        .post(newSubscription);

      if (createdSubscription) {
        console.log(
          `Created new subscription with ID ${createdSubscription.id}`,
        );
        console.log(
          `Please activate the ${settings.eventGridTopic} partner topic in the Azure portal and create an event subscription.`,
        );
        console.log('See README for details.');
      }
    } catch (err) {
      console.error('Error creating subscription: ', JSON.stringify(err));
    }
  }
}
