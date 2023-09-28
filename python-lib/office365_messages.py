class Office365Messages(object):
    def __init__(self, parent):
        self.session = parent

    def get_next_message(self, user_principal_name):
        url = "/".join(
            [
                self.session.get_endpoint_url(),
                "users",
                user_principal_name,
                "messages"
            ]
        )
        for message in self.session.get_next_item(
            url=url
        ):
            yield message
