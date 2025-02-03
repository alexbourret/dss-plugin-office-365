from office365_commons import assert_response_ok
from sharepoint_constants import SharePointConstants
from dss_constants import DSSConstants


class Office365Drive(object):
    def __init__(self, parent, drive_id):
        self.session = parent
        self.drive_id = drive_id

    def get_item(self, item_path):
        item = self.session.get_item(
            url=self.get_item_url(item_path)
        )
        return item

    def get_permission_list(self, item_id):
        url = self.get_item_by_id_url(item_id) + "/permissions"
        list = self.session.get_item(
            url=url
        )
        return list

    def get_group(self, group_id):
        # This requires Group.Read.All scope which requires admin consent
        url = "/".join([
            self.session.get_endpoint_url(),
            "groups",
            group_id
        ])
        group = self.session.get_item(url=url)
        return group

    def get_next_child(self, folder_path):
        url = self.get_children_url(folder_path)
        for child in self.session.get_next_item(url=url):
            yield child

    def get_next_child_by_id(self, folder_id):
        for item in self.session.get_next_item(
            url=self.get_item_by_id_children_url(folder_id)
        ):
            yield item

    def delete_item_by_id(self, item_id):
        self.session.request(
            method="DELETE",
            url=self.get_item_by_id_url(item_id)
        )

    def move_item(self, path_from, path_to):
        item_from = self.get_item(path_from)
        item_from_id = item_from.get("id")
        full_to_path, file_name = split_file_path(path_to)
        item_to = self.get_item(full_to_path)
        item_to_id = item_to.get("id")
        return self.move_item_with_id(item_from_id, item_to_id, file_name)

    def move_item_with_id(self, item_from_id, item_to_id, file_name):
        data = {
            "parentReference": {
                "id": item_to_id
            },
            "name": file_name
        }
        url = self.get_item_by_id_url(item_from_id)
        headers = DSSConstants.JSON_HEADERS
        response = self.session.request(
            method="PATCH",
            url=url,
            headers=headers,
            json=data
        )
        json_response = response.json()
        return json_response

    def create_empty_item(self, parent_id, path):
        response = self.session.request(
            method="PUT",
            url=self.get_content_url(parent_id, path)
        )
        json_response = response.json()
        return json_response

    def create_upload_session(self, item_id):
        response = self.session.request(
            method="POST",
            url=self.get_create_upload_session_url(item_id)
        )
        json_response = response.json()
        return json_response

    def write_chunked_file_content(self, upload_url, data, chunk_size=SharePointConstants.FILE_UPLOAD_CHUNK_SIZE):
        file_size = len(data)
        save_upload_offset = 0
        while save_upload_offset < file_size:
            next_save_upload_offset = save_upload_offset + chunk_size
            if next_save_upload_offset > file_size:
                next_save_upload_offset = file_size
            headers = DSSConstants.JSON_HEADERS
            headers.update(
                {
                    "Content-Range": "bytes {}-{}/{}".format(save_upload_offset, next_save_upload_offset-1, file_size)
                }
            )
            response = self.session.request(
                method="PUT",
                url=upload_url,
                headers=headers,
                data=data[save_upload_offset:next_save_upload_offset]
            )
            assert_response_ok(response)
            save_upload_offset = next_save_upload_offset

    def get_children_url(self, folder_path):
        if (not folder_path) or (folder_path == "/"):
            url = self.get_item_url(folder_path) + "/children"
        else:
            url = self.get_item_url(folder_path) + ":/children"
        return url

    def get_item_by_id_children_url(self, item_id):
        url = "/".join(
            [
                self.get_item_by_id_url(item_id),
                "children"
            ]
        )
        return url

    def get_content_url(self, content_parent_id, content_path):
        url = "/".join(
            [
                self.get_item_by_id_url("{}:".format(content_parent_id)),
                "{}:".format(content_path),
                "content"
            ]
        )
        return url

    def get_create_upload_session_url(self, item_id):
        url = "/".join(
            [
                self.get_item_by_id_url(item_id),
                "createUploadSession"
            ]
        )
        return url

    def get_item_by_id_url(self, item_id):
        url = "/".join(
            [
                self.get_drives_url(),
                "items",
                item_id
            ]
        )
        return url

    def get_item_url(self, item_path):
        if (not item_path) or (item_path == "/"):
            url = "/".join(
                [
                    self.get_drives_url(),
                    "root/"
                ]
            )
        else:
            url = "/".join(
                [
                    self.get_drives_url(),
                    "root:/{}".format(item_path)
                ]
            )
        return url

    def get_drives_url(self):
        return "/".join(
            [
                self.session.get_endpoint_url(),
                "drives",
                "{}".format(self.drive_id)
            ]
        )


def split_file_path(file_path):
    file_path_tokens = file_path.split("/")
    path_to_file = "/".join(file_path_tokens[:-1])
    file_name = file_path_tokens[-1:][0]
    return path_to_file, file_name
