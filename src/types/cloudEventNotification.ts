// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ChangeNotification } from '@microsoft/microsoft-graph-types';

export type CloudEventNotification = {
  id: string;
  type: string;
  source: string;
  subject: string;
  time: Date;
  dataContentType: string;
  specVersion: string;
  data: ChangeNotification;
};
