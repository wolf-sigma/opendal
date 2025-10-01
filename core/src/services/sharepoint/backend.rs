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
use std::sync::Arc;

use http::Response;
use http::StatusCode;

use super::core::SharePointCore;
use super::delete::SharePointDeleter;
use super::error::parse_error;
use super::lister::SharePointLister;
use super::writer::SharePointWriter;
use crate::raw::*;
use crate::*;

#[derive(Clone)]
pub struct SharepointBackend {
    pub core: Arc<SharePointCore>,
}

impl Debug for SharepointBackend {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("SharepointBackend")
            .field("core", &self.core)
            .finish()
    }
}

impl Access for SharepointBackend {
    type Reader = HttpBody;
    type Writer = oio::OneShotWriter<SharePointWriter>;
    type Lister = oio::PageLister<SharePointLister>;
    type Deleter = oio::OneShotDeleter<SharePointDeleter>;

    fn info(&self) -> Arc<AccessorInfo> {
        self.core.info.clone()
    }

    async fn create_dir(&self, path: &str, _args: OpCreateDir) -> Result<RpCreateDir> {
        if path == "/" {
            // skip, the root path exists in SharePoint.
            return Ok(RpCreateDir::default());
        }

        let response = self.core.sharepoint_create_dir(path).await?;
        match response.status() {
            StatusCode::CREATED | StatusCode::OK => Ok(RpCreateDir::default()),
            _ => Err(parse_error(response)),
        }
    }

    async fn stat(&self, path: &str, args: OpStat) -> Result<RpStat> {
        let meta = self.core.sharepoint_stat(path, args).await?;

        Ok(RpStat::new(meta))
    }

    async fn read(&self, path: &str, args: OpRead) -> Result<(RpRead, Self::Reader)> {
        let response = self.core.sharepoint_get_content(path, &args).await?;
        match response.status() {
            StatusCode::OK | StatusCode::PARTIAL_CONTENT => {
                Ok((RpRead::default(), response.into_body()))
            }
            _ => {
                let (part, mut body) = response.into_parts();
                let buf = body.to_buffer().await?;
                Err(parse_error(Response::from_parts(part, buf)))
            }
        }
    }

    async fn write(&self, path: &str, args: OpWrite) -> Result<(RpWrite, Self::Writer)> {
        Ok((
            RpWrite::default(),
            oio::OneShotWriter::new(SharePointWriter::new(
                self.core.clone(),
                args,
                path.to_string(),
            )),
        ))
    }

    async fn delete(&self) -> Result<(RpDelete, Self::Deleter)> {
        Ok((
            RpDelete::default(),
            oio::OneShotDeleter::new(SharePointDeleter::new(self.core.clone())),
        ))
    }

    async fn list(&self, path: &str, args: OpList) -> Result<(RpList, Self::Lister)> {
        let l = SharePointLister::new(path.to_string(), self.core.clone(), &args);
        Ok((RpList::default(), oio::PageLister::new(l)))
    }

    async fn copy(&self, from: &str, to: &str, _args: OpCopy) -> Result<RpCopy> {
        let monitor_url = self.core.initialize_copy(from, to).await?;
        self.core.wait_until_complete(monitor_url).await?;
        Ok(RpCopy::default())
    }

    async fn rename(&self, from: &str, to: &str, _args: OpRename) -> Result<RpRename> {
        if from == to {
            return Ok(RpRename::default());
        }

        self.core.sharepoint_move(from, to).await?;

        Ok(RpRename::default())
    }
}