from dataiku.fsprovider import FSProvider
from office365_client import Office365Session
from office365_commons import get_credentials_from_config, format_date, get_rel_path, get_lnt_path
from safe_logger import SafeLogger
from dss_constants import DSSConstants
import os
import shutil

try:
    from BytesIO import BytesIO  # for Python 2
except ImportError:
    from io import BytesIO  # for Python 3


logger = SafeLogger("office-365 plugin", DSSConstants.SECRET_PARAMETERS_KEYS)


class Office365FSProvider(FSProvider):
    def __init__(self, root, config, plugin_config):
        """
        :param root: the root path for this provider
        :param config: the dict of the configuration of the object
        :param plugin_config: contains the plugin settings
        """
        logger.info("Office-365 plugin v{} fs-provider".format(DSSConstants.PLUGIN_VERSION))
        logger.info("config={}".format(logger.filter_secrets(config)))
        if len(root) > 0 and root[0] == '/':
            root = root[1:]
        self.root = root
        self.provider_root = "/"
        auth_token = get_credentials_from_config(config)
        self.session = Office365Session(auth_token)
        self.sharepoint_drive_id = config.get("sharepoint_drive_id")
        if self.sharepoint_drive_id == "dku_manual_select":
            self.sharepoint_site_id = config.get("sharepoint_site_id")
            if self.sharepoint_site_id == "dku_manual_select":
                sharepoint_site_overwrite = config.get("sharepoint_site_overwrite")
                self.sharepoint_site_id = self.session.get_site_id(sharepoint_site_overwrite)
            site = self.session.get_site(self.sharepoint_site_id)
            sharepoint_root_overwrite = config.get("sharepoint_root_overwrite")
            self.sharepoint_drive_id = site.get_drive_id(sharepoint_root_overwrite)
        self.sharepoint_drive = self.session.get_drive(self.sharepoint_drive_id)

    def get_full_path(self, path):
        path_elts = [self.provider_root, get_rel_path(self.root), get_rel_path(path)]
        path_elts = [e for e in path_elts if len(e) > 0]
        return os.path.join(*path_elts)

    def close(self):
        """
        Perform any necessary cleanup
        """
        logger.info('close')

    def stat(self, path):
        """
        Get the info about the object at the given path inside the provider's root, or None
        if the object doesn't exist
        """
        full_path = self.get_full_path(path)
        logger.info("stat:path={}, full_path={}".format(path, full_path))
        target_path = full_path if len(full_path) < 2 else full_path.strip("/")
        item = self.sharepoint_drive.get_item(target_path)
        if not item:
            logger.info("stat:Item {} not found".format(path))
            return None
        if "folder" in item:
            return {
                'path': get_lnt_path(path),
                'size': 0,
                'lastModified': int(format_date(item.get("lastModifiedDateTime"))) if item.get("lastModifiedDateTime") else None,
                'isDirectory': True
            }
        else:
            return {
                'path': get_lnt_path(path),
                'size': item.get("size"),
                'lastModified': int(format_date(item.get("lastModifiedDateTime"))) if item.get("lastModifiedDateTime") else None,
                'isDirectory': False
            }

    def set_last_modified(self, path, last_modified):
        """
        Set the modification time on the object denoted by path. Return False if not possible
        """
        return False

    def browse(self, path):
        path = get_rel_path(path)
        full_path = get_lnt_path(self.get_full_path(path))
        logger.info("browse:path={}, full_path={}".format(path, full_path))

        item = self.sharepoint_drive.get_item(full_path)

        if not item:
            logger.info("Item {} not found".format(path))
            return {
                'fullPath': None,
                'exists': False
            }
        if "folder" not in item:
            return {
                'fullPath': get_lnt_path(path),
                'exists': True,
                'directory': False,
                'size': item.get("size"),
                'lastModified': int(format_date(item.get("lastModifiedDateTime"))) if item.get("lastModifiedDateTime") else None
            }
        children = []
        for item in self.sharepoint_drive.get_next_child(full_path):
            if "folder" in item:
                children.append(
                    {
                        'fullPath': get_lnt_path(os.path.join(path, item.get("name"))), 'exists': True, 'directory': True, 'size': 0
                    }
                )
            else:
                children.append(
                    {
                        'fullPath': get_lnt_path(os.path.join(path, item.get("name"))),
                        'exists': True,
                        'directory': False,
                        'size': item.get("size"),
                        'lastModified': int(format_date(item.get("lastModifiedDateTime"))) if item.get("lastModifiedDateTime") else None
                    }
                )
        return {'fullPath': get_lnt_path(path), 'exists': True, 'directory': True, 'children': children}

    def enumerate(self, path, first_non_empty):
        """
        Enumerate files recursively from prefix. If first_non_empty, stop at the first non-empty file.

        If the prefix doesn't denote a file or folder, return None
        """
        full_path = self.get_full_path(path)
        logger.info("enumerate:path={}, full_path={}".format(path, full_path))
        item = self.sharepoint_drive.get_item(full_path)
        if not item:
            logger.info("Item {} not found".format(path))
            return None
        if "folder" not in item:
            return [{
                'path': get_lnt_path(path)
            }]
        folder_id = item.get("id")
        ret = self.list_recursive(path, full_path, folder_id, first_non_empty)
        return ret

    def list_recursive(self, path, full_path, folder_id, first_non_empty):
        paths = []
        for item in self.sharepoint_drive.get_next_child_by_id(folder_id):
            if "folder" in item:
                paths.extend(
                    self.list_recursive(
                        get_lnt_path(os.path.join(path, item.get("name"))),
                        get_lnt_path(os.path.join(full_path, item.get("name"))),
                        item.get("id"),
                        first_non_empty
                    )
                )
            else:
                paths.append({
                    "path": get_lnt_path(os.path.join(path, item.get("name"))),
                    "lastModified": int(format_date(item.get("lastModifiedDateTime"))) if item.get("lastModifiedDateTime") else None,
                    "size": item.get("size")
                })
                if first_non_empty:
                    return paths
        return paths

    def delete_recursive(self, path):
        """
        Delete recursively from path. Return the number of deleted files (optional)
        """
        full_path = self.get_full_path(path)
        logger.info("delete_recursive:path={}, full_path={}".format(path, full_path))

        item = self.sharepoint_drive.get_item(full_path)
        item_id = item.get("id")
        number_deleted_items = 1 if "folder" not in item else int(item.get("folder", {}).get("childCount", 1))+1
        self.sharepoint_drive.delete_item_by_id(item_id)
        return number_deleted_items

    def move(self, from_path, to_path):
        """
        Move a file or folder to a new path inside the provider's root. Return false if the moved file didn't exist
        """
        full_from_file_path = self.get_full_path(from_path)
        full_to_file_path = self.get_full_path(to_path)

        logger.info("move:full_from_path={}, full_to_path={}".format(full_from_file_path, full_to_file_path))
        json_response = self.sharepoint_drive.move_item(full_from_file_path, full_to_file_path)
        if "id" in json_response:
            return True
        return False

    def read(self, path, stream, limit):
        """
        Read the object denoted by path into the stream. Limit is an optional bound on the number of bytes to send
        """
        full_path = self.get_full_path(path)
        logger.info("read:full_path={}".format(full_path))
        target_path = full_path if len(full_path) < 2 else full_path.strip("/")
        item = self.sharepoint_drive.get_item(target_path)
        download_url = item.get("@microsoft.graph.downloadUrl")
        response = self.session.get(url=download_url)
        bio = BytesIO(response.content)
        shutil.copyfileobj(bio, stream)

    def write(self, path, stream):
        """
        Write the stream to the object denoted by path into the stream
        """
        full_path = self.get_full_path(path)
        full_path_parent = os.path.dirname(full_path)
        logger.info("write:path={}, full_path={}".format(path, full_path))

        parent_item = self.sharepoint_drive.get_item(full_path_parent)
        parent_id = parent_item.get("id")

        json_response = self.sharepoint_drive.create_empty_item(parent_id, path)
        item_id = json_response.get("id")

        json_response = self.sharepoint_drive.create_upload_session(item_id)
        upload_url = json_response.get("uploadUrl")
        bio = BytesIO()
        shutil.copyfileobj(stream, bio)
        bio.seek(0)
        data = bio.read()

        self.sharepoint_drive.write_chunked_file_content(upload_url, data)
