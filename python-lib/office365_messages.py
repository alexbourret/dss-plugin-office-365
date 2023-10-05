class Office365Messages(object):
    def __init__(self, parent, search_space=None):
        self.session = parent
        self.search_space = "user" or search_space

    def get_next_message(self, user_principal_name=None, folder_id=None):
        if self.search_space == "user":
            url = "/".join(
                [
                    self.session.get_endpoint_url(),
                    "users",
                    user_principal_name,
                    "messages"
                ]
            )
        if self.search_space == "me":
            url = "/".join(
                [
                    self.session.get_endpoint_url(),
                    "me",
                    "messages"
                ]
            )
        if self.search_space == "folder":
            url = "/".join(
                [
                    self.session.get_endpoint_url(),
                    "me",
                    "mailFolders",
                    folder_id,
                    "messages"
                ]
            )
        if self.search_space == "user-folder":
            url = "/".join(
                [
                    self.session.get_endpoint_url(),
                    "users",
                    user_principal_name,
                    "mailFolders",
                    folder_id,
                    "messages"
                ]
            )
        for message in self.session.get_next_item(
            url=url
        ):
            yield message
