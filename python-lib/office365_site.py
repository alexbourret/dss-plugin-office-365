from office365_list import Office365List
import urllib.parse


class Office365Site(object):

    def __init__(self, parent, site_id):
        self.session = parent
        self.site_id = site_id

    def get_list(self, list_id):
        return Office365List(self, list_id)

    def get_list_id(self, list_name):
        for list in self.get_next_list():
            list_web_url = urllib.parse.unquote(list.get("webUrl"))
            if list_web_url.endswith(list_name):
                return list.get("id")
        return None

    def get_drive_id(self, drive_name):
        for drive in self.get_next_drive():
            drive_web_url = urllib.parse.unquote(drive.get("webUrl"))
            if drive_web_url.endswith(drive_name):
                return drive.get("id")
        return None

    def get_next_list(self):
        url = "/".join(
            [
                self.get_site_url(),
                "lists"
            ]
        )
        for list in self.session.get_next_item(
            url=url
        ):
            yield list

    def get_next_drive(self):
        url = "/".join(
            [
                self.get_site_url(),
                "drives"
            ]
        )
        for drive in self.session.get_next_item(
            url=url
        ):
            yield drive

    def get_site_url(self):
        return "/".join(
            [
                self.session.get_endpoint_url(),
                "sites",
                "{}".format(self.site_id)
            ]
        )
