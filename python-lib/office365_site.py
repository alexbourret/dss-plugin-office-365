from office365_list import Office365List


class Office365Site(object):

    def __init__(self, parent, site_id):
        self.session = parent
        self.site_id = site_id

    def get_list(self, list_id):
        return Office365List(self, list_id)

    def get_site_url(self):
        return "https://graph.microsoft.com/v1.0/sites/{}".format(self.site_id)
