import { IExecuteFunctions, INodeExecutionData, NodeApiError } from "n8n-workflow";
import { MSGetItemDetailsByPath } from "../../helpers/misc";
import { makeMicrosoftRequest } from "../../helpers/makeMicrosoftRequest";

interface PermissionEntry {
	grantedToV2?: {
		user?: {
			email: string;
		};
		group?: {
			email: string;
		};
	};
	roles?: string[];
	id?: string;
}

/**
 * Sets permissions for a folder in SharePoint. Supports:
 * - Multiple users/groups with different permission levels
 * - Both by folder path and ID
 * - Adding or replacing existing permissions
 * - Applying to folder only or recursively to contents
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
	const permsByLevel: { [key: string]: string[] } = { view: [], edit: [], none: [] };
	for (const perm of permissions) {
		const email = (perm.email as string).trim();
		const permissionLevel = perm.permission as 'view' | 'edit' | 'none';
		if (email) {
			permsByLevel[permissionLevel].push(email);
		}
	}

	// Process removals ('none' permissions)
	for (const email of permsByLevel.none) {
		try {
			const existingPermission = await findPermissionByEmail(this, siteId, folderId, email);
			if (existingPermission && existingPermission.id) {
				await removePermission(this, siteId, folderId, existingPermission.id);
				setPermissions.push({
					email,
					action: 'removed',
					permissionId: existingPermission.id,
				});
			} else {
				setPermissions.push({
					email,
					action: 'not_found',
					message: 'No existing permission found to remove',
				});
			}
		} catch (error) {
			setPermissions.push({
				email,
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
	emails: string[],
	role: 'read' | 'write',
	setPermissions: any[]
): Promise<void> {
	const inviteBody = {
		recipients: emails.map(email => ({ email })),
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
			const email = emails[idx];

			// Check if this specific invite had an error
			if (inviteData.error) {
				setPermissions.push({
					email,
					action: 'failed',
					error: inviteData.error.message || 'Unknown error',
					errorCode: inviteData.error.code,
				});
			} else {
				setPermissions.push({
					email,
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
		for (const email of emails) {
			setPermissions.push({
				email,
				action: 'failed',
				error: (error as any).message || 'Unknown error',
			});
		}
	}
}

/**
 * Helper function to find a permission entry by email
 */
async function findPermissionByEmail(
	thisRef: IExecuteFunctions,
	siteId: string,
	folderId: string,
	email: string
): Promise<PermissionEntry | null> {
	const permissions = await getAllPermissions(thisRef, siteId, folderId);

	for (const perm of permissions) {
		const userEmail = perm.grantedToV2?.user?.email;
		const groupEmail = perm.grantedToV2?.group?.email;

		if (userEmail === email || groupEmail === email) {
			return perm;
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
