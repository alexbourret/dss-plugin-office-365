import dataiku
import pandas
from dataiku.customrecipe import get_output_names_for_role
from dataiku.customrecipe import get_recipe_config
from office365_commons import get_credentials_from_config
from office365_client import Office365Session
from datetime import datetime
from safe_logger import SafeLogger
from dss_constants import DSSConstants


logger = SafeLogger("office-365 plugin", DSSConstants.SECRET_PARAMETERS_KEYS)

TIME_FORMAT = "%Y-%m-%dT%H:%M:%SZ"


def reorder_permissions(permissions):
    owners_emails = []
    reads_emails = []
    writes_emails = []
    owners_ids = []
    reads_ids = []
    writes_ids = []
    for permission in permissions:
        permission_id = permission.get("id")
        if len(permission_id) > 23:
            roles = permission.get("roles", [])
            granted_to_v2 = permission.get("grantedToV2", {})
            grant_type = "group"
            if "user" in granted_to_v2:
                grant_type = "user"
            grant = granted_to_v2.get(grant_type, {})
            email = grant.get("email")
            id = grant.get("id")
            # Note: id is always there, email not always
            # on entra, email is the user principal name, but also in contact information / email
            # if id and grant_type == "group":
            #     group = sharepoint_drive.get_group(id) # requires admin consent
            #     print("ALX:group={}".format(group))
            if email:
                for role in roles:
                    if role == "owner":
                        owners_emails.append(email)
                    if role == "read":
                        reads_emails.append(email)
                    if role == "write":
                        writes_emails.append(email)
            if id:
                for role in roles:
                    if role == "owner":
                        owners_ids.append(id)
                    if role == "read":
                        reads_ids.append(id)
                    if role == "write":
                        writes_ids.append(id)
    return owners_emails, reads_emails, writes_emails, owners_ids, reads_ids, writes_ids


def get_last_modified(paths_details, file_path):
    # extracts the last modified date of the file_path file
    # from the paths_details dic
    for path_detail in paths_details:
        full_path = path_detail.get("fullPath").strip("/")
        is_directory = path_detail.get("directory")
        if not is_directory and file_path == full_path:
            return path_detail.get("lastModified")


def sharepoint_date_to_epoch(date):
    if date is not None:
        utc_time = datetime.strptime(date, TIME_FORMAT)
        epoch_time = (utc_time - datetime(1970, 1, 1)).total_seconds()
        return int(epoch_time) * 1000
    else:
        return None


def build_file_path(path_elements, file_name):
    return "/".join(path_elements + [file_name])


def process_folder(sharepoint_path):
    logger.info("Processing folder '{}'".format(sharepoint_path))
    item = sharepoint_drive.get_item(sharepoint_path)
    results = []
    for item in sharepoint_drive.get_next_child(sharepoint_path):
        sub_item_name = item.get("name")
        next_path = "/".join([sharepoint_path, sub_item_name])
        if "folder" in item:
            logger.info("'{}' is a folder, recursing...".format(sharepoint_path))
            try:
                results += process_folder(next_path)
            except Exception as error:
                logger.error("Error {}".format(error))
        else:
            result = {}
            download_url = item.get("@microsoft.graph.downloadUrl")
            folder_details = files_folders[0].get_path_details(path=sharepoint_path)
            paths_details = folder_details.get("children", [])
            local_last_modified = get_last_modified(paths_details, next_path)
            remote_last_modifider = sharepoint_date_to_epoch(item.get("lastModifiedDateTime"))
            result["path"] = next_path
            result["details"] = item
            if download_url:
                item_id = item.get("id")
                list = sharepoint_drive.get_permission_list(item_id)
                result["permissions"] = list.get("value")
                result["owner_email"], result["read_email"], result["write_email"], result["owner_id"], result["read_id"], result["write_id"] = reorder_permissions(list.get("value"))
            if download_url and (not local_last_modified or local_last_modified < remote_last_modifider):
                logger.info("'{}' has been modified since last sync ({} / {}) -> downloading".format(
                    next_path, local_last_modified, remote_last_modifider
                ))
                response = session.get(url=download_url)
                with files_folders[0].get_writer(next_path) as local_file_handle:
                    local_file_handle.write(response.content)
            else:
                logger.info("'{}' already downloaded, skipping".format(next_path))
            results.append(result)
    return results


file_security_names = get_output_names_for_role('file_security')
file_security_datasets = [dataiku.Dataset(name) for name in file_security_names]

files_folder_names = get_output_names_for_role('files_folder')
files_folders = [dataiku.Folder(name) for name in files_folder_names]

target_folder_paths = files_folders[0].list_paths_in_partition()
folder_details = files_folders[0].get_path_details()
# paths_details = folder_details.get("children", [])

config = get_recipe_config()
auth_token = get_credentials_from_config(config)

sharepoint_path = config.get("sharepoint_path", "/")
sharepoint_drive_id = config.get("sharepoint_drive_id")

session = Office365Session(auth_token)
sharepoint_drive = session.get_drive(sharepoint_drive_id)

item = sharepoint_drive.get_item(sharepoint_path)
results = []

results = process_folder(sharepoint_path)

odf = pandas.DataFrame(results)
output = file_security_datasets[0]
output.write_with_schema(odf)
