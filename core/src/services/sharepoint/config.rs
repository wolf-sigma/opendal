// Licensed to the Apache Software Foundation (ASF) under one
// or more contributor license agreements.  See the NOTICE file
// distributed with this work for additional information
// regarding copyright ownership.  The ASF licenses this file
// to you under the Apache License, Version 2.0 (the
// "License"); you may not use this file except in compliance
// with the License.  You may obtain a copy of the License at
//
//   http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

use std::fmt::Debug;
use std::fmt::Formatter;

use serde::Deserialize;
use serde::Serialize;

/// Config for [SharePoint](https://sharepoint.microsoft.com) backend support.
#[derive(Default, Serialize, Deserialize, Clone, PartialEq, Eq)]
#[serde(default)]
#[non_exhaustive]
pub struct SharepointConfig {
    /// The root path for the SharePoint service for the file access
    pub root: Option<String>,
    /// SharePoint site ID or site URL (e.g., "contoso.sharepoint.com,12345678-1234-1234-1234-123456789012,87654321-4321-4321-4321-210987654321")
    pub site_id: Option<String>,
    /// SharePoint site URL (alternative to site_id, e.g., "https://contoso.sharepoint.com/sites/sitename")
    pub site_url: Option<String>,
    /// Optional drive ID for specific document library (uses default site drive if not specified)
    pub drive_id: Option<String>,
    /// Microsoft Graph API access token
    pub access_token: Option<String>,
    /// Microsoft Graph API refresh token
    pub refresh_token: Option<String>,
    /// Microsoft Graph API Application (client) ID that is in the Azure's app registration portal
    pub client_id: Option<String>,
    /// Microsoft Graph API Application client secret that is in the Azure's app registration portal
    pub client_secret: Option<String>,
    /// Enabling version support
    pub enable_versioning: bool,
}

impl Debug for SharepointConfig {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("SharepointConfig")
            .field("root", &self.root)
            .field("site_id", &self.site_id)
            .field("site_url", &self.site_url)
            .field("drive_id", &self.drive_id)
            .finish_non_exhaustive()
    }
}