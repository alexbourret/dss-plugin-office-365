import requests
from safe_logger import SafeLogger
from office365_site import Office365Site
from office365_drive import Office365Drive
from office365_messages import Office365Messages
from office365_auth import Office365Auth
from office365_commons import get_next_page_url, get_error, prepare_row
from dss_constants import DSSConstants


logger = SafeLogger("office-365 plugin", [])


class Office365Session():
    def __init__(self, access_token=None):
        self.session = requests.Session()
        self.session.auth = Office365Auth(access_token=access_token)
        self.is_batch_mode = False
        self.requests_buffer = []
        self.batch_size = 0

    def request(self, **kwargs):
        raise_on = kwargs.pop("raise_on", {})
        cannot_raise = kwargs.pop("cannot_raise", False)
        force_no_batch = kwargs.pop("force_no_batch", False)

        if self.is_batch_mode and not force_no_batch:
            self.requests_buffer.append(kwargs)
            if len(self.requests_buffer) >= self.batch_size:
                self.flush()
            return

        response = self.session.request(**kwargs)
        error_message = get_error(response)
        if raise_on:
            status_code = response.status_code
            error_message = raise_on.get(status_code)
            if error_message:
                raise Exception(error_message)
        if error_message and not cannot_raise:
            raise Exception(error_message)
        return response

    def get(self, **kwargs):
        kwargs["method"] = "GET"
        response = self.request(**kwargs)
        error_message = get_error(response)
        if error_message and not kwargs.get("cannot_raise"):
            raise Exception(error_message)
        return response

    def get_headers(self):
        headers = DSSConstants.JSON_HEADERS
        return headers

    def get_item(self, **kwargs):
        kwargs["headers"] = kwargs.get("headers", {})
        kwargs["headers"].update(DSSConstants.JSON_HEADERS)
        kwargs["headers"].update(DSSConstants.GZIP_HEADERS)
        kwargs["cannot_raise"] = True
        response = self.get(
            **kwargs
        )
        status_code = response.status_code
        if status_code == 404:
            return {}
        json_response = response.json()
        return json_response

    def get_next_item(self, **kwargs):
        kwargs["headers"] = kwargs.get("headers", {})
        kwargs["headers"].update(DSSConstants.JSON_HEADERS)
        kwargs["headers"].update(DSSConstants.GZIP_HEADERS)
        is_first_get = True
        next_page_url = None
        while next_page_url or is_first_get:
            kwargs["url"] = next_page_url or kwargs["url"]
            if next_page_url:
                # As next_page_url already contains query params
                kwargs["params"] = None
            response = self.get(
                **kwargs
            )
            is_first_get = False
            json_response = response.json()
            next_page_url = get_next_page_url(json_response)
            items = json_response.get("value", [])
            for item in items:
                yield item

    def get_next_site(self):
        for site in self.get_next_item(
                url=self.get_sites_url(),
                params={"search": "*"}
        ):
            yield site

    def get_all_items(self, **kwargs):
        items = []
        for item in self.get_next_item(**kwargs):
            items.append(item)
        return items

    def start_batch_mode(self, batch_size=None):
        batch_size = batch_size or DSSConstants.DEFAULT_BATCH_SIZE
        self.is_batch_mode = True
        self.batch_size = batch_size
        self.requests_buffer = []

    def close(self):
        self.flush()
        self.is_batch_mode = False

    def flush(self):
        responses = self.process_batch(self.requests_buffer)
        self.requests_buffer = []
        assert_responses_ok(responses)

    def get_site(self, site_id):
        return Office365Site(self, site_id)

    def get_messages(self, search_space=None):
        return Office365Messages(self, search_space=search_space)

    def get_site_id(self, site_name):
        search_by_web_url = True if "/" in site_name else False
        for site in self.get_next_site():
            if search_by_web_url:
                full_site_path = "/".join(site.get("webUrl").split("/")[3:])
                if full_site_path == site_name:
                    return site.get("id")
            else:
                if site.get("name") == site_name:
                    return site.get("id")
        return None

    def get_drive(self, drive_id):
        return Office365Drive(self, drive_id)

    def process_batch(self, requests_buffer):
        if not requests_buffer:
            return {}
        data = {}
        requests = []
        counter = 1
        for request_kwargs in requests_buffer:
            request = {
                "id": "{}".format(counter),
                "method": request_kwargs.get("method"),
                "url": self.get_relative_url(request_kwargs.get("url")),
            }
            if request_kwargs.get("headers"):
                request["headers"] = request_kwargs.get("headers")
            if request_kwargs.get("json"):
                request["body"] = request_kwargs.get("json")
            if request_kwargs.get("data"):
                request["data"] = request_kwargs.get("data")
            requests.append(
                request
            )
            counter += 1
        data["requests"] = requests
        response = self.session.request(
            method="POST",
            url=self.get_batch_url(),
            headers=DSSConstants.JSON_HEADERS,
            json=data
        )
        status_code = response.status_code
        if status_code >= 400:
            error_message = "Batch error {}".format(status_code)
            try:
                json_response = response.json()
                error_message += ". {}".format(json_response.get("error").get("message"))
            except Exception as sub_error_message:
                logger.debug("Could not enrich error message {}".format(sub_error_message))
            logger.error("Error {}, dumping content: {}".format(status_code, response.content))
            raise Exception("Error {}".format(status_code))
        json_response = response.json()
        return json_response.get("responses", {})

    def get_batch_url(self):
        return "/".join(
            [
                self.get_endpoint_url(),
                "$batch"
            ]
        )

    def get_sites_url(self):
        return "/".join(
            [
                self.get_endpoint_url(),
                "sites"
            ]
        )

    def get_relative_url(self, full_url):
        url_base = self.get_endpoint_url()
        relative_url = full_url
        if full_url.startswith(url_base):
            relative_url = full_url.replace(url_base, "")
        return relative_url

    def get_endpoint_url(self):
        return "https://graph.microsoft.com/v1.0"


def get_relative_url(url_base, full_url):
    relative_url = full_url
    if full_url.startswith(url_base):
        relative_url = full_url.replace(url_base, "")
    return relative_url


def assert_responses_ok(responses):
    for response in responses:
        if int(response.get("status", 200)) >= 400:
            logger.error("Error during batch, dumping responses: {}".format(responses))
            raise Exception("Batch id {} failed with error {}. {}".format(
                response.get("id"),
                response.get("status"),
                response.get("body")
            ))
    return True


class Office365ListWriter(object):
    def __init__(self, list, dataset_schema, batch_size=None):
        self.list = list
        self.list.session.start_batch_mode(batch_size=batch_size)
        self.columns = dataset_schema.get("columns")

    def write_row(self, row):
        self.list.write_row(prepare_row(row, self.columns))

    def close(self):
        self.list.session.close()
