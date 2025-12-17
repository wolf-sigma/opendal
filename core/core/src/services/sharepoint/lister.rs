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

use std::sync::Arc;

use bytes::Buf;

use super::core::SharePointCore;
use super::error::parse_error;
use super::graph_model::GraphApiSharePointListResponse;
use super::graph_model::ItemType;
use super::graph_model::GENERAL_SELECT_PARAM;
use crate::raw::oio;
use crate::raw::*;
use crate::*;

pub struct SharePointLister {
    core: Arc<SharePointCore>,
    path: String,
    op: OpList,
}

impl SharePointLister {
    const DRIVE_ROOT_PREFIX: &'static str = "/drive/root:";

    pub(crate) fn new(path: String, core: Arc<SharePointCore>, args: &OpList) -> Self {
        Self {
            core,
            path,
            op: args.clone(),
        }
    }
}

impl oio::PageList for SharePointLister {
    async fn next_page(&self, ctx: &mut oio::PageContext) -> Result<()> {
        let request_url = if ctx.token.is_empty() {
            let base = format!(
                "{}/children?{}",
                self.core.sharepoint_item_url(&self.path, true).await?,
                GENERAL_SELECT_PARAM
            );
            if let Some(limit) = self.op.limit() {
                base + &format!("&$top={limit}")
            } else {
                base
            }
        } else {
            ctx.token.clone()
        };

        let response = self.core.sharepoint_get_next_list_page(&request_url).await?;

        let status_code = response.status();
        if !status_code.is_success() {
            if status_code == http::StatusCode::NOT_FOUND {
                ctx.done = true;
                return Ok(());
            }
            return Err(parse_error(response));
        }

        let bytes = response.into_body();

        // Log the raw JSON response for debugging
        let body_str = String::from_utf8_lossy(bytes.chunk());
        log::debug!("Raw SharePoint API response: {}", body_str);

        let decoded_response: GraphApiSharePointListResponse =
            serde_json::from_reader(bytes.reader()).map_err(new_json_deserialize_error)?;

        let list_with_versions = self.core.info.native_capability().list_with_versions;

        // Include the current directory itself when handling the first page of the listing.
        if ctx.token.is_empty() && !ctx.done {
            // TODO: when listing a directory directly, we could reuse the stat result,
            // cache the result when listing nested directory
            let path = if self.path == "/" {
                "".to_string()
            } else {
                self.path.clone()
            };

            let meta = self.core.sharepoint_stat(&path, OpStat::default()).await?;

            // skip `list_with_versions` intentionally because a folder doesn't have versions

            let entry = oio::Entry::new(&path, meta);
            ctx.entries.push_back(entry);
        }

        if let Some(next_link) = decoded_response.next_link {
            ctx.token = next_link;
        } else {
            ctx.done = true;
        }

        for drive_item in decoded_response.value {
            let name = drive_item.name;
            let parent_path = drive_item.parent_reference
                .as_ref()
                .and_then(|pr| pr.path.as_deref())
                .unwrap_or("/drive/root:");
            let parent_path = parent_path
                .strip_prefix(Self::DRIVE_ROOT_PREFIX)
                .unwrap_or("");

            let path = format!("{parent_path}/{name}");
            let mut normalized_path = build_rel_path(self.core.root.as_str(), path.as_str());
            let entry_mode = match drive_item.item_type {
                ItemType::Folder { .. } => EntryMode::DIR,
                ItemType::File { .. } => EntryMode::FILE,
            };

            // Add the trailing `/` because SharePoint returns a directory with the name
            if entry_mode == EntryMode::DIR {
                normalized_path.push('/');
            }

            let mut meta = Metadata::new(entry_mode);

            if let Some(etag) = drive_item.e_tag {
                meta = meta.with_etag(etag);
            }

            if let Some(size) = drive_item.size {
                meta = meta.with_content_length(size.max(0) as u64);
            }

            if let Some(last_modified_str) = drive_item.last_modified_date_time {
                let last_modified = parse_datetime_from_rfc3339(last_modified_str.as_str())?;
                meta.set_last_modified(last_modified);
            }

            // When listing a directory with `$expand=versions`, SharePoint returns 400 "Operation not supported".
            // Thus, `list_with_versions` induces N+1 requests. This N+1 is intentional.
            // N+1 is horrendous but we can't do any better without SharePoint's API support.
            // When SharePoint supports listing with versions API, remove this.
            if list_with_versions {
                let versions = self.core.sharepoint_list_versions(&path).await?;
                if let Some(version) = versions.first() {
                    meta.set_version(&version.id);
                }
            }

            let entry = oio::Entry::new(&normalized_path, meta);
            ctx.entries.push_back(entry)
        }

        Ok(())
    }
}
