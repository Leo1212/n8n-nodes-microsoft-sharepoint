import { IExecuteFunctions, INodeExecutionData, NodeApiError } from "n8n-workflow";
import { MSGetItemDetailsByPath } from "../../helpers/misc";
import { makeMicrosoftRequest } from "../../helpers/makeMicrosoftRequest";

interface PermissionEntry {
	grantedToV2?: {
		user?: {
			email?: string;
			displayName?: string;
			id?: string;
		};
		group?: {
			email?: string;
			displayName?: string;
			id?: string;
		};
		siteUser?: {
			displayName?: string;
			email?: string;
			id?: string;
			loginName?: string;
		};
	};
	roles?: string[];
	id?: string;
}

interface PermissionInput {
	type: 'user' | 'group';
	email: string;
	permission: 'view' | 'edit' | 'none';
}

/**
 * NOTE: Microsoft Graph API does not support download restrictions.
 * The "Download not possible" permission available in SharePoint UI
 * is NOT exposed through the Microsoft Graph API. This is controlled via:
 * - SharePoint Information Rights Management (IRM) at the site/library level
 * - Microsoft Purview Data Loss Prevention (DLP) policies at the tenant level
 * - Custom SharePoint permission levels configured by administrators
 * 
 * See: https://learn.microsoft.com/en-us/graph/api/resources/permission?view=graph-rest-1.0#roles-property-values
 * Only supported roles: 'read', 'write', 'owner'
 */

/**
 * Sets permissions for a folder in SharePoint. Supports:
 * - Multiple users/groups with different permission levels
 * - Both by folder path and ID
 * - Adding or replacing existing permissions
 * - Applying to folder only or recursively to contents
 *
 * NOTE: The "View - Download Not Possible" permission is NOT supported by Microsoft Graph API.
 * This is a SharePoint-specific feature that can only be configured via:
 * 1. SharePoint UI with Information Rights Management (IRM)
 * 2. Microsoft Purview DLP policies
 * 3. Custom SharePoint permission levels configured at the site level
 * 
 * The current implementation will set it as a standard "read" permission, but the 
 * download restriction must be applied separately via site settings or policies.
 *
 * https://learn.microsoft.com/en-us/graph/api/driveitem-invite?view=graph-rest-1.0&tabs=http
 * @param this
 * @param i
 * @returns
 */
export async function execute(this: IExecuteFunctions, i: number): Promise<INodeExecutionData[]> {
	const siteId = this.getNodeParameter('siteId', i) as string;
	const libraryId = this.getNodeParameter('libraryId', i) as string;
	const options = this.getNodeParameter('options', i, {});

	// Get folder identifier (path or ID)
	const folderLocator = this.getNodeParameter('folderLocator', i) as {
		mode: 'path' | 'id';
		value: string;
	};

	let folderId: string;

	if (folderLocator.mode === 'path') {
		const folderPath = (folderLocator.value || '').replace(/\/$/, '');
		const res = await MSGetItemDetailsByPath(this, libraryId, folderPath);
		folderId = res.id;
	} else {
		folderId = folderLocator.value || '';
	}

	if (!folderId) {
		throw new NodeApiError(this.getNode(), {}, {
			message: 'Could not determine folder ID. Please check the folder path or ID.',
		});
	}

	// Get permission inputs (fixedCollection returns data nested under the collection name)
	const permissionsData = this.getNodeParameter('permissions', i) as any;
	const permissions = permissionsData?.permission || [];
	const permissionBehavior = (options.permissionBehavior as string) || 'add';
	const applyRecursively = (options.applyRecursively as boolean) || false;

	if (!permissions || permissions.length === 0) {
		throw new NodeApiError(this.getNode(), {}, {
			message: 'At least one permission must be specified.',
		});
	}

	// If replace mode, delete all non-owner permissions first
	if (permissionBehavior === 'replace') {
		const allPermissions = await getAllPermissions(this, siteId, folderId);
		for (const perm of allPermissions) {
			// Don't delete owner permissions
			if (perm.roles && !perm.roles.includes('owner')) {
				try {
					await removePermission(this, siteId, folderId, perm.id!);
				} catch (error) {
					this.logger.warn(`Could not remove permission ${perm.id}: ${(error as any).message}`);
				}
			}
		}
	}

	// Set permissions for each user/group
	const setPermissions: any[] = [];
	
	// Group permissions by level for batch processing
	const permsByLevel: { [key: string]: PermissionInput[] } = { view: [], edit: [], none: [] };
	for (const perm of permissions) {
		const email = (perm.email as string).trim();
		const permissionLevel = perm.permission as 'view' | 'edit' | 'none';
		const type = (perm.type as 'user' | 'group') || 'user';
		if (email) {
			permsByLevel[permissionLevel].push({ type, email, permission: permissionLevel });
		}
	}

	// Process removals ('none' permissions)
	for (const permInput of permsByLevel.none) {
		try {
			this.logger.info(`Attempting to remove permission for ${permInput.type}: ${permInput.email}`);
			const existingPermission = await findPermissionByEmail(this, siteId, folderId, permInput.email);
			
			if (existingPermission && existingPermission.id) {
				this.logger.info(`Found permission with ID: ${existingPermission.id}`);
				await removePermission(this, siteId, folderId, existingPermission.id);
				setPermissions.push({
					email: permInput.email,
					type: permInput.type,
					action: 'removed',
					permissionId: existingPermission.id,
				});
			} else {
				this.logger.warn(`Permission not found for ${permInput.email}`);
				setPermissions.push({
					email: permInput.email,
					type: permInput.type,
					action: 'not_found',
					message: 'No existing permission found to remove',
				});
			}
		} catch (error) {
			this.logger.error(`Error removing permission for ${permInput.email}: ${(error as any).message}`);
			setPermissions.push({
				email: permInput.email,
				type: permInput.type,
				action: 'failed',
				error: (error as any).message || 'Unknown error',
			});
		}
	}

	// Process view permissions as a batch
	if (permsByLevel.view.length > 0) {
		await processInviteBatch(this, siteId, folderId, permsByLevel.view, 'read', setPermissions);
	}

	// Process edit permissions as a batch
	if (permsByLevel.edit.length > 0) {
		await processInviteBatch(this, siteId, folderId, permsByLevel.edit, 'write', setPermissions);
	}

	// Get all current permissions
	const allPermissions = await getAllPermissions(this, siteId, folderId);

	return [
		{
			json: {
				success: true,
				folder: {
					id: folderId,
					siteId,
					libraryId,
				},
				setPermissions,
				allPermissions,
				appliedRecursively: applyRecursively,
			},
		},
	];
}

/**
 * Helper function to process invite batch for multiple recipients
 */
async function processInviteBatch(
	thisRef: IExecuteFunctions,
	siteId: string,
	folderId: string,
	permInputs: PermissionInput[],
	role: 'read' | 'write',
	setPermissions: any[]
): Promise<void> {
	// Build recipients array with proper driveRecipient format
	const recipients = permInputs.map(perm => {
		const recipient: any = {};
		
		if (perm.type === 'group') {
			// For groups, try email first, then alias for group names
			if (perm.email.includes('@')) {
				recipient.email = perm.email;
			} else {
				// Treat as alias (group name without @domain)
				recipient.alias = perm.email;
			}
		} else {
			// For users, always use email
			recipient.email = perm.email;
		}
		
		return recipient;
	});

	const inviteBody: any = {
		recipients,
		roles: [role],
		requireSignIn: true,
		sendInvitation: false,
	};

	try {
		const inviteResult = await makeMicrosoftRequest(
			thisRef,
			`sites/${siteId}/drive/items/${folderId}/invite`,
			{
				method: 'POST',
				body: inviteBody,
			}
		);

		// Process each recipient in the response
		const results = inviteResult.value || [];
		for (let idx = 0; idx < results.length; idx++) {
			const inviteData = results[idx];
			const permInput = permInputs[idx];

			// Check if this specific invite had an error
			if (inviteData.error) {
				setPermissions.push({
					email: permInput.email,
					type: permInput.type,
					action: 'failed',
					error: inviteData.error.message || 'Unknown error',
					errorCode: inviteData.error.code,
				});
			} else {
				setPermissions.push({
					email: permInput.email,
					type: permInput.type,
					permission: role === 'write' ? 'edit' : 'view',
					id: inviteData.id,
					action: 'set',
					grantedTo: inviteData.grantedToV2 || inviteData.grantedTo,
					roles: inviteData.roles,
					shareId: inviteData.shareId,
				});
			}
		}
	} catch (error) {
		// If the entire batch failed, mark all as failed
		for (const permInput of permInputs) {
			setPermissions.push({
				email: permInput.email,
				type: permInput.type,
				action: 'failed',
				error: (error as any).message || 'Unknown error',
			});
		}
	}
}

/**
 * Helper function to find a permission entry by email, alias, or display name
 * Tries multiple matching strategies to handle both user emails and group names
 */
async function findPermissionByEmail(
	thisRef: IExecuteFunctions,
	siteId: string,
	folderId: string,
	searchTerm: string
): Promise<PermissionEntry | null> {
	const permissions = await getAllPermissions(thisRef, siteId, folderId);
	const searchTermLower = searchTerm.toLowerCase();

	for (const perm of permissions) {
		// Try matching user email
		const userEmail = perm.grantedToV2?.user?.email?.toLowerCase();
		if (userEmail && userEmail === searchTermLower) {
			return perm;
		}

		// Try matching group email
		const groupEmail = perm.grantedToV2?.group?.email?.toLowerCase();
		if (groupEmail && groupEmail === searchTermLower) {
			return perm;
		}

		// Try matching group by alias or name (handles case where alias was used instead of email)
		const groupDisplayName = perm.grantedToV2?.group?.displayName?.toLowerCase();
		if (groupDisplayName && groupDisplayName === searchTermLower) {
			return perm;
		}

		// Try matching user display name
		const userDisplayName = perm.grantedToV2?.user?.displayName?.toLowerCase();
		if (userDisplayName && userDisplayName === searchTermLower) {
			return perm;
		}

		// Try matching by login name (handles special formats)
		const loginName = perm.grantedToV2?.siteUser?.loginName?.toLowerCase();
		if (loginName) {
			// Extract email part from login name (e.g., "i:0#.f|membership|user@domain.com" -> "user@domain.com")
			const emailMatch = loginName.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
			if (emailMatch && emailMatch[0].toLowerCase() === searchTermLower) {
				return perm;
			}
		}
	}

	return null;
}

/**
 * Helper function to get all permissions on a folder
 */
async function getAllPermissions(
	thisRef: IExecuteFunctions,
	siteId: string,
	folderId: string
): Promise<PermissionEntry[]> {
	try {
		const result = await makeMicrosoftRequest(
			thisRef,
			`sites/${siteId}/drive/items/${folderId}/permissions`
		);

		return result.value || [];
	} catch (error) {
		thisRef.logger.warn(`Could not fetch all permissions: ${(error as any).message}`);
		return [];
	}
}

/**
 * Helper function to remove a permission
 */
async function removePermission(
	thisRef: IExecuteFunctions,
	siteId: string,
	folderId: string,
	permissionId: string
): Promise<void> {
	await makeMicrosoftRequest(
		thisRef,
		`sites/${siteId}/drive/items/${folderId}/permissions/${permissionId}`,
		{
			method: 'DELETE',
		}
	);
}
