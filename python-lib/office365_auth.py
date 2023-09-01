import requests


class Office365Auth(requests.auth.AuthBase):
    def __init__(self, access_token=None):
        self.access_token = access_token

    def __call__(self, request):
        request.headers["Authorization"] = "Bearer {}".format(
            self.access_token
        )
        return request
