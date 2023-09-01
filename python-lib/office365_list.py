from office365_commons import get_sharepoint_type_descriptor


class Office365List(object):
    def __init__(self, parent, list_id):
        self.session = parent.session
        self.list_id = list_id
        self.parent = parent

    def get_columns(self):
        url = self.get_column_url()
        return self.session.get_all_items(url=url)

    def get_column_url(self):
        url = "/".join(
            [
                self.parent.get_site_url(), "lists/{}/columns".format(
                    self.list_id
                )
            ]
        )
        return url

    def get_next_row(self):
        url = self.get_next_list_row_url()
        for row in self.session.get_next_item(
            url=url,
            params={"expand": "field"},
            force_no_batch=True
        ):
            yield row

    def get_next_list_row_url(self):
        url = "/".join(
            [
                self.parent.get_site_url(),
                "lists/{}/items".format(
                    self.list_id
                )
            ]
        )
        return url

    def add_column(self, name, type, description=None):
        description = "" or description
        data = {
            "description": description,
            "enforceUniqueValues": False,
            "hidden": False,
            "indexed": False,
            "name": name,
        }
        data.update(
            get_sharepoint_type_descriptor(type)
        )
        url = self.get_column_url()
        self.session.requests(
            method="POST",
            url=url,
            headers={"Content-Type": "application/json"},
            json=data,
            raise_on={403: "Check that your Azure app has Sites.Manage.All scope enabled"}
        )

    def write_row(self, row):
        headers = {
            "Content-Type": "application/json"
        }
        data = {
            "fields": row,
        }
        self.session.requests(
            method="POST",
            url=self.get_next_list_row_url(),
            headers=headers,
            json=data
        )

    def delete_row(self, row_id):
        self.session.requests(
            method="DELETE",
            url=self.get_list_row_id_url(row_id)
        )

    def get_list_row_id_url(self, row_id):
        url = "/".join(
            [
                self.parent.get_site_url(),
                "lists/{}/items/{}".format(
                    self.list_id,
                    row_id
                )
            ]
        )
        return url

    def delete_all_rows(self):
        self.session.start_batch_mode(batch_size=100)
        for row in self.get_next_row():
            row_id = row.get("id")
            self.delete_row(row_id)
        self.session.close()
