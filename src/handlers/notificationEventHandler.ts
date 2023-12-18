// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Request, Response } from 'express';
import { Client, GraphError } from '@microsoft/microsoft-graph-client';
import { CloudEventNotification } from '../types/cloudEventNotification';
import { Subscription, User } from '@microsoft/microsoft-graph-types';

export function handleValidationRequest(req: Request, res: Response) {
  // See https://github.com/cloudevents/spec/blob/v1.0/http-webhook.md#4-abuse-protection
  // Event Grid sends the host that emits events in this header as a request
  // for our webhook to allow them to send
  const origin = req.header('Webhook-Request-Origin');
  const rate = req.header('Webhook-Request-Rate');

  if (origin) {
    res.setHeader('Webhook-Allowed-Origin', origin);
  }

  if (rate) {
    res.setHeader('Webhook-Allowed-Rate', rate);
  }

  res.sendStatus(200);
}

export async function handleNotificationAsync(req: Request, res: Response) {
  const notification = req.body as CloudEventNotification;
  console.log(`Received ${notification.type} notification from Event Grid`);

  switch (notification.type.toLowerCase()) {
    case 'microsoft.graph.userupdated':
      await handleUserUpdateAsync(notification, req.app.locals.graphClient);
      break;
    case 'microsoft.graph.userdeleted':
      handleUserDelete(notification);
      break;
    case 'microsoft.graph.subscriptionreauthorizationrequired':
      await handleSubscriptionRenewalAsync(
        notification,
        req.app.locals.graphClient,
      );
      break;
    default:
      console.log('Unexpected notification type - no action taken');
  }
  res.sendStatus(202);
}

async function handleUserUpdateAsync(
  notification: CloudEventNotification,
  graphClient: Client,
) {
  // The user was either created, updated, or soft-deleted.
  // The notification only contains the user's ID, so
  // get the user from Microsoft Graph if other details are needed.
  // If the user isn't found, then it was likely soft-deleted.
  // The notification has the relative URL to the user.
  if (notification.data.resource) {
    try {
      const user: User = await graphClient
        .api(notification.data.resource)
        .get();

      console.log(
        `User ${user.displayName} (ID: ${user.id}) was created or updated`,
      );
    } catch (err) {
      if (
        err instanceof GraphError &&
        err.code?.toLowerCase().includes('resourcenotfound')
      ) {
        const userId = notification.data.resource.split('/')[1];
        console.log(`User with ID ${userId} was soft-deleted`);
      } else {
        console.error('Error getting user:', JSON.stringify(err));
      }
    }
  }
}

async function handleUserDelete(notification: CloudEventNotification) {
  // The user was permanently deleted. The notification only contains
  // the user's ID, and we can no longer get the user from Graph.
  const userId = notification.data.resource?.split('/')[1];
  console.log(`User with ID ${userId} was permanently deleted`);
}

async function handleSubscriptionRenewalAsync(
  notification: CloudEventNotification,
  graphClient: Client,
) {
  // The subscription needs to be renewed.
  if (notification.data.subscriptionId) {
    const subscriptionUpdate: Subscription = {
      expirationDateTime: new Date(Date.now() + 3600000).toISOString(),
    };
    try {
      await graphClient
        .api(`/subscriptions/${notification.data.subscriptionId}`)
        .patch(subscriptionUpdate);

      console.log(
        `Subscription with ID ${notification.data.subscriptionId} renewed for another hour`,
      );
    } catch (err) {
      console.error('Error renewing subscription: ', JSON.stringify(err));
    }
  }
}
