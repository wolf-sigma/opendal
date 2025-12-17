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
use std::time::Duration;

use bytes::Buf;
use bytes::Bytes;
use http::header;
use http::Request;
use http::Response;
use http::StatusCode;
use mea::mutex::Mutex;

use super::error::parse_error;
use super::graph_model::*;
use crate::raw::*;
use crate::*;

pub struct SharePointCore {
    pub info: Arc<AccessorInfo>,
    pub root: String,
    pub site_id: Option<String>,
    pub site_url: Option<String>,
    pub drive_id: Option<String>,
    pub signer: Arc<Mutex<SharePointSigner>>,
}

impl Debug for SharePointCore {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("SharePointCore")
            .field("root", &self.root)
            .field("site_id", &self.site_id)
            .field("site_url", &self.site_url)
            .field("drive_id", &self.drive_id)
            .finish_non_exhaustive()
    }
}

// SharePoint returns 400 when try to access a dir with the POSIX special directory entries
const SPECIAL_POSIX_ENTRIES: [&str; 3] = [".", "/", ""];

// organizes a few core module functions
impl SharePointCore {
    // SharePoint base URL patterns
    pub(crate) const GRAPH_BASE_URL: &str = "https://graph.microsoft.com/v1.0";

    /// Get the base drive URL for SharePoint
    ///
    /// This method handles site ID resolution and drive URL construction
    pub(crate) async fn get_drive_root_url(&self) -> Result<String> {
        let base_url = if let Some(ref site_id) = self.site_id {
            // Use provided site ID directly
            format!("{}/sites/{}/drive", Self::GRAPH_BASE_URL, site_id)
        } else if let Some(ref site_url) = self.site_url {
            // Need to resolve site URL to site ID first
            let site_id = self.resolve_site_url_to_id(site_url).await?;
            format!("{}/sites/{}/drive", Self::GRAPH_BASE_URL, site_id)
        } else {
            return Err(Error::new(
                ErrorKind::ConfigInvalid,
                "either site_id or site_url must be configured",
            ));
        };

        // If specific drive ID is provided, use it
        if let Some(ref drive_id) = self.drive_id {
            Ok(format!("{}/sites/{}/drives/{}", Self::GRAPH_BASE_URL,
                      self.get_site_id().await?, drive_id))
        } else {
            Ok(base_url)
        }
    }

    /// Get site ID from either configured site_id or resolved from site_url
    async fn get_site_id(&self) -> Result<String> {
        if let Some(ref site_id) = self.site_id {
            Ok(site_id.clone())
        } else if let Some(ref site_url) = self.site_url {
            self.resolve_site_url_to_id(site_url).await
        } else {
            Err(Error::new(
                ErrorKind::ConfigInvalid,
                "either site_id or site_url must be configured",
            ))
        }
    }

    /// Resolve SharePoint site URL to site ID
    ///
    /// Makes a request to get site information from URL
    async fn resolve_site_url_to_id(&self, site_url: &str) -> Result<String> {
        // Extract hostname and site path from URL
        let url = site_url.trim_end_matches('/');
        let site_api_url = if url.contains("/sites/") {
            // For site URLs like https://contoso.sharepoint.com/sites/sitename
            let parts: Vec<&str> = url.split("/sites/").collect();
            if parts.len() != 2 {
                return Err(Error::new(
                    ErrorKind::ConfigInvalid,
                    "invalid site_url format, expected: https://hostname/sites/sitename",
                ));
            }
            let hostname = parts[0].replace("https://", "");
            let site_path = parts[1];
            format!("{}/sites/{}:/sites/{}:", Self::GRAPH_BASE_URL, hostname, site_path)
        } else {
            return Err(Error::new(
                ErrorKind::ConfigInvalid,
                "site_url must contain /sites/ path",
            ));
        };

        let mut request = Request::get(&site_api_url)
            .extension(Operation::Stat)
            .body(Buffer::new())
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        let response = self.info.http_client().send(request).await?;
        if !response.status().is_success() {
            return Err(parse_error(response));
        }

        let bytes = response.into_body();
        let site_info: SharePointSiteInfo =
            serde_json::from_reader(bytes.reader()).map_err(new_json_deserialize_error)?;

        Ok(site_info.id)
    }

    /// Get a URL to a SharePoint item
    pub(crate) async fn sharepoint_item_url(&self, path: &str, build_absolute_path: bool) -> Result<String> {
        let drive_url = self.get_drive_root_url().await?;

        // SharePoint requires the root to be accessible via the drive root URL.
        // For files under the root, the URL pattern becomes `/drive/root:<path>:`
        if self.root == "/" && SPECIAL_POSIX_ENTRIES.contains(&path) {
            Ok(format!("{}/root", drive_url))
        } else {
            // SharePoint returns 400 when try to access a folder with a ending slash
            let absolute_path = if build_absolute_path {
                let rooted_path = build_rooted_abs_path(&self.root, path);
                rooted_path
                    .strip_suffix('/')
                    .unwrap_or(rooted_path.as_str())
                    .to_string()
            } else {
                path.to_string()
            };
            Ok(format!(
                "{}/root:{}:",
                drive_url,
                percent_encode_path(&absolute_path),
            ))
        }
    }

    /// Send a simplest stat request about a particular path
    ///
    /// See also: [`sharepoint_stat()`].
    pub(crate) async fn sharepoint_get_stat_plain(&self, path: &str) -> Result<Response<Buffer>> {
        let url: String = format!(
            "{}?{}",
            self.sharepoint_item_url(path, true).await?,
            GENERAL_SELECT_PARAM
        );
        let request = Request::get(&url);

        let mut request = request
            .extension(Operation::Stat)
            .body(Buffer::new())
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        self.info.http_client().send(request).await
    }

    /// Create a directory at path if not exist, return the metadata about the folder
    ///
    /// When the folder exist, this function works exactly the same as [`sharepoint_get_stat_plain()`].
    ///
    /// * `path` - a relative folder path
    pub(crate) async fn ensure_directory(&self, path: &str) -> Result<SharePointItem> {
        let response = self.sharepoint_get_stat_plain(path).await?;
        let item: SharePointItem = match response.status() {
            StatusCode::OK => {
                let bytes = response.into_body();
                serde_json::from_reader(bytes.reader()).map_err(new_json_deserialize_error)?
            }
            StatusCode::NOT_FOUND => {
                // We must create directory for the destination
                let response = self.sharepoint_create_dir(path).await?;
                match response.status() {
                    StatusCode::CREATED | StatusCode::OK => {
                        let bytes = response.into_body();
                        serde_json::from_reader(bytes.reader())
                            .map_err(new_json_deserialize_error)?
                    }
                    _ => return Err(parse_error(response)),
                }
            }
            _ => return Err(parse_error(response)),
        };

        Ok(item)
    }

    pub(crate) async fn sign<T>(&self, request: &mut Request<T>) -> Result<()> {
        let mut signer = self.signer.lock().await;
        signer.sign(request).await
    }
}

// SharePoint copy action is asynchronous. We query an endpoint and wait 1 second.
// This is the maximum attempts we will wait.
const MAX_MONITOR_ATTEMPT: i32 = 3600;
const MONITOR_WAIT_SECOND: u64 = 1;

// SharePoint API parameters allows using with a parameter of:
//
// - ID
// - file path
//
// `services-sharepoint` uses the file path based API for simplicity.
// Read more at https://learn.microsoft.com/en-us/graph/onedrive-addressing-driveitems
impl SharePointCore {
    /// Send a stat request about a particular path, including:
    ///
    /// - Get stat object only if ETag not matches
    /// - whether to get the object version
    ///
    /// See also [`sharepoint_get_stat_plain()`].
    pub(crate) async fn sharepoint_stat(&self, path: &str, args: OpStat) -> Result<Metadata> {
        let mut url: String = self.sharepoint_item_url(path, true).await?;
        if args.version().is_some() {
            url += "?$expand=versions(";
            url += VERSION_SELECT_PARAM;
            url += ")";
        }

        let mut request = Request::get(&url);
        if let Some(etag) = args.if_none_match() {
            request = request.header(header::IF_NONE_MATCH, etag);
        }

        let mut request = request
            .extension(Operation::Stat)
            .body(Buffer::new())
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        let response = self.info.http_client().send(request).await?;
        if !response.status().is_success() {
            return Err(parse_error(response));
        }

        let bytes = response.into_body();
        let decoded_response: SharePointItem =
            serde_json::from_reader(bytes.reader()).map_err(new_json_deserialize_error)?;

        let entry_mode: EntryMode = match decoded_response.item_type {
            ItemType::Folder { .. } => EntryMode::DIR,
            ItemType::File { .. } => EntryMode::FILE,
        };

        let mut meta = Metadata::new(entry_mode);

        if let Some(etag) = decoded_response.e_tag {
            meta = meta.with_etag(etag);
        }

        if let Some(size) = decoded_response.size {
            meta = meta.with_content_length(size.max(0) as u64);
        }

        if let Some(version) = args.version() {
            for item_version in decoded_response.versions.as_deref().unwrap_or_default() {
                if item_version.id == version {
                    meta.set_version(version);
                    break; // early exit
                }
            }

            if meta.version().is_none() {
                return Err(Error::new(
                    ErrorKind::NotFound,
                    "cannot find this version of the item",
                ));
            }
        }

        if let Some(last_modified) = decoded_response.last_modified_date_time {
            let date_utc_last_modified = last_modified.parse::<Timestamp>()?;
            meta.set_last_modified(date_utc_last_modified);
        }

        Ok(meta)
    }

    /// Return versions of an item
    ///
    /// A folder has no versions.
    ///
    /// * `path` - a relative path
    pub(crate) async fn sharepoint_list_versions(
        &self,
        path: &str,
    ) -> Result<Vec<SharePointItemVersion>> {
        // don't `$select` this endpoint to get the download URL.
        let url: String = format!(
            "{}:/versions?{}",
            self.sharepoint_item_url(path, true).await?,
            VERSION_SELECT_PARAM
        );

        let mut request = Request::get(url)
            .extension(Operation::List)
            .body(Buffer::new())
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        let response = self.info.http_client().send(request).await?;
        let decoded_response: GraphApiSharePointVersionsResponse =
            serde_json::from_reader(response.into_body().reader())
                .map_err(new_json_deserialize_error)?;
        Ok(decoded_response.value)
    }

    pub(crate) async fn sharepoint_get_next_list_page(&self, url: &str) -> Result<Response<Buffer>> {
        let mut request = Request::get(url)
            .extension(Operation::List)
            .body(Buffer::new())
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        self.info.http_client().send(request).await
    }

    /// Download a file
    ///
    /// SharePoint handles a download in 2 steps:
    /// 1. Returns a 302 with a presigned URL. If `If-None-Match` succeed, returns 304.
    /// 2. With the presigned URL, we can send a GET:
    ///   1. When getting an item succeed with a `Range` header, we get a 206 Partial Content response.
    ///   2. When succeed, we get a 200 response.
    ///
    /// Read more at https://learn.microsoft.com/en-us/graph/api/driveitem-get-content
    pub(crate) async fn sharepoint_get_content(
        &self,
        path: &str,
        args: &OpRead,
    ) -> Result<Response<HttpBody>> {
        // We can't "select" the SharePoint API response fields when reading because "select" shadows not found error
        let url: String = format!("{}/content", self.sharepoint_item_url(path, true).await?);

        let mut request = Request::get(&url).header(header::RANGE, args.range().to_header());
        if let Some(etag) = args.if_none_match() {
            request = request.header(header::IF_NONE_MATCH, etag);
        }

        let mut request = request
            .extension(Operation::Read)
            .body(Buffer::new())
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        self.info.http_client().fetch(request).await
    }

    /// Upload a file
    ///
    /// When creating a file,
    ///
    /// * SharePoint returns 201 if the file is new.
    /// * SharePoint returns 200 if successfully overwrote the file successfully.
    ///
    /// Read more at https://learn.microsoft.com/en-us/onedrive/developer/rest-api/api/driveitem_put_content
    ///
    /// This function is different than uploading a file with chunks.
    /// See also [`create_upload_session()`] and [`SharePointWriter::write_chunked`].
    pub async fn sharepoint_upload_simple(
        &self,
        path: &str,
        args: &OpWrite,
        body: Buffer,
    ) -> Result<Response<Buffer>> {
        let url = format!(
            "{}:/content?@microsoft.graph.conflictBehavior={}&{}",
            self.sharepoint_item_url(path, true).await?,
            REPLACE_EXISTING_ITEM_WHEN_CONFLICT,
            GENERAL_SELECT_PARAM
        );

        // SharePoint upload API documentation requires "text/plain" as the content type.
        // In practice, SharePoint ignores the content type,
        // but decides the type (when stating) based on the extension name.
        // Also, when the extension name is unknown to SharePoint,
        // SharePoint sets the content type as "application/octet-stream".
        // We keep the content type according to the documentation.
        let mut request = Request::put(&url)
            .header(header::CONTENT_LENGTH, body.len())
            .header(header::CONTENT_TYPE, "text/plain");

        // when creating a new file, `IF-Match` has no effect.
        // when updating a file with the `If-Match`, and if the ETag mismatched,
        // SharePoint will return 412 Precondition Failed
        if let Some(if_match) = args.if_match() {
            request = request.header(header::IF_MATCH, if_match);
        }

        let mut request = request
            .extension(Operation::Write)
            .body(body)
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        self.info.http_client().send(request).await
    }

    pub(crate) async fn sharepoint_chunked_upload(
        &self,
        url: &str,
        args: &OpWrite,
        offset: usize,
        chunk_end: usize,
        total_len: usize,
        body: Buffer,
    ) -> Result<Response<Buffer>> {
        let mut request = Request::put(url);

        let range = format!("bytes {offset}-{chunk_end}/{total_len}");
        request = request.header(header::CONTENT_RANGE, range);

        let size = chunk_end - offset + 1;
        request = request.header(header::CONTENT_LENGTH, size);

        if let Some(mime) = args.content_type() {
            request = request.header(header::CONTENT_TYPE, mime)
        }

        let request = request
            .extension(Operation::Write)
            .body(body)
            .map_err(new_request_build_error)?;
        // SharePoint documentation requires not sending the `Authorization` header

        self.info.http_client().send(request).await
    }

    /// Create a upload session for chunk uploads
    ///
    /// This endpoint supports `If-None-Match` but [`sharepoint_upload_simple()`] doesn't.
    ///
    /// Read more at https://learn.microsoft.com/en-us/onedrive/developer/rest-api/api/driveitem_createuploadsession?view=odsp-graph-online#upload-bytes-to-the-upload-session
    pub(crate) async fn sharepoint_create_upload_session(
        &self,
        path: &str,
        args: &OpWrite,
    ) -> Result<Response<Buffer>> {
        let parent_path = get_parent(path);
        let file_name = get_basename(path);
        let url = format!(
            "{}:/createUploadSession",
            self.sharepoint_item_url(parent_path, true).await?,
        );
        let mut request = Request::post(url).header(header::CONTENT_TYPE, "application/json");

        if let Some(if_match) = args.if_match() {
            request = request.header(header::IF_MATCH, if_match);
        }

        let body = SharePointUploadSessionCreationRequestBody::new(file_name.to_string());
        let body_bytes = serde_json::to_vec(&body).map_err(new_json_serialize_error)?;
        let body = Buffer::from(Bytes::from(body_bytes));
        let mut request = request
            .extension(Operation::Write)
            .body(body)
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        self.info.http_client().send(request).await
    }

    /// Create a directory
    ///
    /// When creating a folder, SharePoint returns a status code with 201.
    /// When using `microsoft.graph.conflictBehavior=replace` to replace a folder, SharePoint returns 200.
    ///
    /// * `path` - the path to the folder without the root
    pub(crate) async fn sharepoint_create_dir(&self, path: &str) -> Result<Response<Buffer>> {
        let parent_path = get_parent(path);
        let basename = get_basename(path);
        let folder_name = basename.strip_suffix('/').unwrap_or(basename);

        let url = format!(
            "{}:/children?{}",
            self.sharepoint_item_url(parent_path, true).await?,
            GENERAL_SELECT_PARAM
        );

        let payload = CreateDirPayload::new(folder_name.to_string());
        let body_bytes = serde_json::to_vec(&payload).map_err(new_json_serialize_error)?;
        let body = Buffer::from(bytes::Bytes::from(body_bytes));

        let mut request = Request::post(url)
            .header(header::CONTENT_TYPE, "application/json")
            .extension(Operation::CreateDir)
            .body(body)
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        self.info.http_client().send(request).await
    }

    /// Delete a `DriveItem`
    ///
    /// This moves the items to the recycle bin.
    pub(crate) async fn sharepoint_delete(&self, path: &str) -> Result<Response<Buffer>> {
        let url = self.sharepoint_item_url(path, true).await?;

        let mut request = Request::delete(&url)
            .extension(Operation::Delete)
            .body(Buffer::new())
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        self.info.http_client().send(request).await
    }

    /// Initialize a copy
    ///
    /// * `source` - the path to the source folder without the root
    /// * `destination` - the path to the destination folder without the root
    ///
    /// See also: [`wait_until_complete()`]
    pub(crate) async fn initialize_copy(&self, source: &str, destination: &str) -> Result<String> {
        // we must validate if source exist
        let response = self.sharepoint_get_stat_plain(source).await?;
        if !response.status().is_success() {
            return Err(parse_error(response));
        }

        // We need to stat the destination parent folder to get a parent reference
        let destination_parent = get_parent(destination).to_string();
        let basename = get_basename(destination);

        let item = self.ensure_directory(&destination_parent).await?;
        let body = SharePointPatchRequestBody {
            parent_reference: ParentReference {
                path: Some("".to_string()), // irrelevant for copy
                drive_id: item.parent_reference.as_ref().map(|pr| pr.drive_id.clone()).unwrap_or_default(),
                id: Some(item.id),
            },
            name: basename.to_string(),
        };

        // ensure the destination file or folder doesn't exist
        let response = self.sharepoint_get_stat_plain(destination).await?;
        match response.status() {
            // We must remove the file or folder because
            // SharePoint doesn't support `conflictBehavior` for the consumer SharePoint.
            // `conflictBehavior` seems to work for the consumer SharePoint sometimes could be a coincidence.
            // Read more at https://learn.microsoft.com/en-us/graph/api/driveitem-copy
            StatusCode::OK => {
                let response = self.sharepoint_delete(destination).await?;
                match response.status() {
                    StatusCode::NO_CONTENT | StatusCode::NOT_FOUND => {} // expected, intentionally empty
                    _ => return Err(parse_error(response)),
                }
            }
            StatusCode::NOT_FOUND => {} // expected, intentionally empty
            _ => return Err(parse_error(response)),
        }

        let url: String = format!("{}:/copy", self.sharepoint_item_url(source, true).await?);

        let body_bytes = serde_json::to_vec(&body).map_err(new_json_serialize_error)?;
        let buffer = Buffer::from(Bytes::from(body_bytes));
        let mut request = Request::post(&url)
            .header(header::CONTENT_TYPE, "application/json")
            .extension(Operation::Copy)
            .body(buffer)
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        let response = self.info.http_client().send(request).await?;
        match response.status() {
            StatusCode::ACCEPTED => parse_location(response.headers())?
                .ok_or_else(|| {
                    Error::new(
                        ErrorKind::Unexpected,
                        "SharePoint didn't return a location URL",
                    )
                })
                .map(String::from),
            _ => Err(parse_error(response)),
        }
    }

    pub(crate) async fn wait_until_complete(&self, monitor_url: String) -> Result<()> {
        for _attempt in 0..MAX_MONITOR_ATTEMPT {
            let mut request = Request::get(monitor_url.to_string())
                .header(header::CONTENT_TYPE, "application/json")
                .extension(Operation::Copy)
                .body(Buffer::new())
                .map_err(new_request_build_error)?;

            self.sign(&mut request).await?;

            let response = self.info.http_client().send(request).await?;
            let status: SharePointMonitorStatus =
                serde_json::from_reader(response.into_body().reader())
                    .map_err(new_json_deserialize_error)?;
            if status.status == "completed" {
                return Ok(());
            }

            tokio::time::sleep(Duration::from_secs(MONITOR_WAIT_SECOND)).await;
        }

        Err(Error::new(
            ErrorKind::Unexpected,
            "Exceed monitoring timeout",
        ))
    }

    pub(crate) async fn sharepoint_move(&self, source: &str, destination: &str) -> Result<()> {
        // We must validate if the source folder exists.
        let response = self.sharepoint_get_stat_plain(source).await?;
        if !response.status().is_success() {
            return Err(Error::new(ErrorKind::NotFound, "source not found"));
        }

        // We want a parent reference about the destination's parent, or the destination folder itself.
        let destination_parent = get_parent(destination).to_string();
        let basename = get_basename(destination);

        let item = self.ensure_directory(&destination_parent).await?;
        let body = SharePointPatchRequestBody {
            parent_reference: ParentReference {
                path: Some("".to_string()), // irrelevant for update
                // reusing `ParentReference` for convenience. The API requires this value to be correct.
                drive_id: item.parent_reference.as_ref().map(|pr| pr.drive_id.clone()).unwrap_or_default(),
                id: Some(item.id),
            },
            name: basename.to_string(),
        };
        let body_bytes = serde_json::to_vec(&body).map_err(new_json_serialize_error)?;
        let buffer = Buffer::from(Bytes::from(body_bytes));
        let url: String = format!(
            "{}?@microsoft.graph.conflictBehavior={}&$select=id",
            self.sharepoint_item_url(source, true).await?,
            REPLACE_EXISTING_ITEM_WHEN_CONFLICT
        );
        let mut request = Request::patch(&url)
            .header(header::CONTENT_TYPE, "application/json")
            .extension(Operation::Rename)
            .body(buffer)
            .map_err(new_request_build_error)?;

        self.sign(&mut request).await?;

        let response = self.info.http_client().send(request).await?;
        match response.status() {
            // can get etag, metadata, etc...
            StatusCode::OK => Ok(()),
            _ => Err(parse_error(response)),
        }
    }
}

// keeps track of OAuth 2.0 tokens and refreshes the access token.
pub struct SharePointSigner {
    pub info: Arc<AccessorInfo>, // to use `http_client`

    pub client_id: String,
    pub client_secret: String,
    pub refresh_token: String,

    pub access_token: String,
    pub expires_in: Timestamp,
}

// SharePoint is part of Graph API hence shares the same authentication and authorization processes.
// `common` applies to account types:
//
// - consumers
// - work and school account
//
// set to `common` for simplicity
const SHAREPOINT_REFRESH_TOKEN: &str = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

impl SharePointSigner {
    pub fn new(info: Arc<AccessorInfo>) -> Self {
        SharePointSigner {
            info,

            client_id: "".to_string(),
            client_secret: "".to_string(),
            refresh_token: "".to_string(),
            access_token: "".to_string(),
            expires_in: Timestamp::MIN,
        }
    }

    async fn refresh_tokens(&mut self) -> Result<()> {
        // SharePoint users must provide at least this required permission scope
        // Sites.ReadWrite.All for SharePoint sites access, Files.ReadWrite for basic file access
        let encoded_payload = format!(
            "client_id={}&client_secret={}&scope=offline_access%20Sites.ReadWrite.All%20Files.ReadWrite&refresh_token={}&grant_type=refresh_token",
            percent_encode_path(self.client_id.as_str()),
            percent_encode_path(self.client_secret.as_str()),
            percent_encode_path(self.refresh_token.as_str())
        );
        let request = Request::post(SHAREPOINT_REFRESH_TOKEN)
            .header(header::CONTENT_TYPE, "application/x-www-form-urlencoded")
            .body(Buffer::from(encoded_payload))
            .map_err(new_request_build_error)?;

        let response = self.info.http_client().send(request).await?;
        match response.status() {
            StatusCode::OK => {
                let resp_body = response.into_body();
                let data: GraphOAuthRefreshTokenResponseBody =
                    serde_json::from_reader(resp_body.reader())
                        .map_err(new_json_deserialize_error)?;
                self.access_token = data.access_token;
                self.refresh_token = data.refresh_token;
                self.expires_in = Timestamp::now() + Duration::from_secs(data.expires_in)
                    - Duration::from_secs(120); // assumes 2 mins graceful transmission for implementation simplicity
                Ok(())
            }
            _ => Err(parse_error(response)),
        }
    }

    /// Sign a request.
    pub async fn sign<T>(&mut self, request: &mut Request<T>) -> Result<()> {
        if !self.access_token.is_empty() && self.expires_in > Timestamp::now() {
            let value = format!("Bearer {}", self.access_token)
                .parse()
                .expect("access_token must be valid header value");

            request.headers_mut().insert(header::AUTHORIZATION, value);
            return Ok(());
        }

        self.refresh_tokens().await?;

        let auth_header_content = format!("Bearer {}", self.access_token)
            .parse()
            .expect("Fetched access_token is invalid as a header value");

        request
            .headers_mut()
            .insert(header::AUTHORIZATION, auth_header_content);

        Ok(())
    }
}