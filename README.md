# n8n-nodes-Sharepoint

This is an n8n community node. It lets you interact with Microsoft Sharepoint in your n8n workflows.

This is a very early version of the node. It is not ready for production use. It contains no error handeling and only supports a limited set of features.

[n8n](https://n8n.io/) is a [fair-code licensed](https://docs.n8n.io/reference/license/) workflow automation platform.

[Installation](#installation)  
[Operations](#operations)  
[Credentials](#credentials)  <!-- delete if no auth needed -->  
[Compatibility](#compatibility)  
[Usage](#usage)  <!-- delete if not using this section -->  
[Resources](#resources)  
[Version history](#version-history) 

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation.

## Operations

* Site
  * Get all sites
* File
  * Upload file
  * Get file
* Folder
  * Create folder
  * List children
  * Set Permissions
  
## Credentials

## Compatibility

## Usage

## Resources

* [n8n community nodes documentation](https://docs.n8n.io/integrations/community-nodes/)

## Version history

* v0.1.6 (2026-01-02)
  * Added new action to set permissions for folders

* v0.1.5 (2024-11-08)
  * Added support for returning all items in a folder (pagination)
  * Honoring Microsoft's Retry-After header

* v0.1.4 (2024-10-30)
  * Added support for uploading files to a folder by ID or path
  * Added support for moving files to other folder by ID or path


* v0.1.3 (2024-10-23)
  * Added support to get file by ID or path

* v0.1.2 (2024-10-17)
  * Added support for creating folders (and nested folders)

* v0.1.1
  * Restructure code + UI. Grouping operations by levels: Site, Folder, File.
  * Added support for moving files

* v0.1.0 (2024-10-12)
  * Initial release

