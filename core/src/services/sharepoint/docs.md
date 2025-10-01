## Capabilities

This service can be used to:

- [x] stat
- [x] read
- [x] write
- [x] create_dir
- [x] delete
- [x] copy
- [x] rename
- [x] list
- [ ] presign
- [ ] blocking

## Configuration

- `root`: Set the working directory for SharePoint
- `site_id`: SharePoint site ID in format "hostname,site-collection-id,site-id"
- `site_url`: Alternative to site_id, SharePoint site URL
- `drive_id`: Optional drive ID for specific document library
- `access_token`: Microsoft Graph API access token
- `refresh_token`: Microsoft Graph API refresh token for long-term access
- `client_id`: Microsoft Graph API application client ID
- `client_secret`: Microsoft Graph API application client secret
- `enable_versioning`: Enable version support

You can refer to [`SharepointBuilder`]'s docs for more information

## Example

### Via Builder

```rust
use anyhow::Result;
use opendal::services::Sharepoint;
use opendal::Operator;

#[tokio::main]
async fn main() -> Result<()> {
    // Create sharepoint backend builder
    let mut builder = Sharepoint::default();

    // Set the site_id (format: hostname,site-collection-id,site-id)
    builder.site_id("contoso.sharepoint.com,12345678-1234-1234-1234-123456789012,87654321-4321-4321-4321-210987654321");

    // Set the root of the backend
    builder.root("/path/to/dir");

    // Set the access_token
    builder.access_token("access_token");

    let op: Operator = Operator::new(builder)?.finish();

    Ok(())
}
```

### Via Config

```rust
use anyhow::Result;
use opendal::services::Sharepoint;
use opendal::Operator;
use std::collections::HashMap;

#[tokio::main]
async fn main() -> Result<()> {
    let mut map = HashMap::new();
    map.insert("root".to_string(), "/path/to/dir".to_string());
    map.insert("site_id".to_string(), "contoso.sharepoint.com,12345678-1234-1234-1234-123456789012,87654321-4321-4321-4321-210987654321".to_string());
    map.insert("access_token".to_string(), "access_token".to_string());

    let op: Operator = Operator::from_map::<Sharepoint>(map)?.finish();
    Ok(())
}
```

## Authentication

This service uses Microsoft Graph API for authentication. You need to register an application in Azure AD and obtain the necessary credentials.

### Required Permissions

For SharePoint access, your application needs the following Microsoft Graph permissions:
- `Sites.ReadWrite.All` - for SharePoint sites access
- `Files.ReadWrite` - for basic file operations
- `offline_access` - for refresh token support (optional, for long-term access)

### Getting Access Token

You can obtain an access token through:
1. Microsoft Graph Explorer (for testing)
2. OAuth 2.0 authorization code flow
3. Client credentials flow (for daemon apps)

For production use, it's recommended to use the refresh token flow for long-term access.