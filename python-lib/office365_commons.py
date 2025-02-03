from datetime import datetime
from safe_logger import SafeLogger
import requests


logger = SafeLogger("office-365 plugin", [])


class RecordsLimit():
    def __init__(self, records_limit=-1):
        self.has_no_limit = (records_limit == -1)
        self.records_limit = records_limit
        self.counter = 0

    def is_reached(self):
        if self.has_no_limit:
            return False
        self.counter += 1
        return self.counter > self.records_limit


def get_credentials_from_config(config):
    auth_token = config.get("sharepoint_oauth", {}).get("sharepoint_oauth")
    return auth_token


def format_date(date):
    if date is not None:
        utc_time = datetime.strptime(date, "%Y-%m-%dT%H:%M:%SZ")
        epoch_time = (utc_time - datetime(1970, 1, 1)).total_seconds()
        return int(epoch_time) * 1000
    else:
        return None


def get_rel_path(path):
    if len(path) > 0 and path[0] == '/':
        path = path[1:]
    return path


def get_lnt_path(path):
    if len(path) == 0 or path == '/':
        return '/'
    elts = path.split('/')
    elts = [e for e in elts if len(e) > 0]
    return '/' + '/'.join(elts)


def get_sharepoint_type_descriptor(dss_type):
    sharepoint_type_descriptor = {
        "text": {
            "allowMultipleLines": False,
            "appendChangesToExistingText": False,
            "linesForEditing": 0,
            "maxLength": 255
        }
    }
    if dss_type == "int":
        sharepoint_type_descriptor = {
            'number': {
                'decimalPlaces': 'automatic',
                'displayAs': 'number',
                'maximum': 1.7976931348623157e+308,
                'minimum': -1.7976931348623157e+308
            }
        }
    if dss_type == "float":
        sharepoint_type_descriptor = {
            'number': {
                'decimalPlaces': 'automatic',
                'displayAs': 'number',
                'maximum': 1.7976931348623157e+308,
                'minimum': -1.7976931348623157e+308
            }
        }
    return sharepoint_type_descriptor


def get_next_page_url(json_response):
    return json_response.get("@odata.nextLink", None)


def get_error(response):
    error_message = None
    if type(response) != requests.Response:
        error_message = "Incorrect response type"
    else:
        status_code = response.status_code
        if status_code >= 400:
            error_message = "Error {} while accessing {}".format(status_code, response.url)
        try:
            json_response = response.json()
            enriched_error_message = json_response.get("error", "").get("message", "")
            error_message += ". {}".format(enriched_error_message)
        except Exception as sub_error_message:
            logger.debug("Could not decode json: {}".format(sub_error_message))
    if error_message:
        logger.error(error_message)
        logger.error("Dumping content: {}".format(response.content))
    return error_message


def is_throttling(response):
    return response.status_code == 429


def get_retry_after_value(response):
    retry_after_value = response.headers.get("Retry-After")
    if retry_after_value:
        return int(retry_after_value)
    return 30


def prepare_row(row, columns):
    prepared_row = {}
    for item, column in zip(row, columns):
        column_type = column.get("type")
        if column_type == "int":
            prepared_row[column.get("name")] = int(item)
        elif column_type == "float":
            prepared_row[column.get("name")] = float(item)
        else:
            prepared_row[column.get("name")] = str(item)
    return prepared_row


def assert_response_ok(response, context=None):
    error_message = get_error(response)
    if error_message and context:
        error_message += "({})".format(context)
    if error_message:
        raise Exception(error_message)


MANUAL_SELECT_ENTRY = {"label": "✍️ Enter manually", "value": "dku_manual_select"}
COLUMN_SELECT_ENTRY = {"label": "🏛️ Get from column", "value": "dku_column_select"}


class DSSSelectorChoices(object):

    def __init__(self):
        self.choices = []

    def append(self, label, value):
        self.choices.append(
            {
                "label": label,
                "value": value
            }
        )

    def append_alphabetically(self, new_label, new_value):
        index = 0
        new_choice = {
            "label": new_label,
            "value": new_value
        }
        for choice in self.choices:
            choice_label = choice.get("label")
            if choice_label < new_label:
                index += 1
            else:
                break
        self.choices.insert(index, new_choice)

    def append_manual_select(self):
        self.choices.append(MANUAL_SELECT_ENTRY)

    def start_with_manual_select(self):
        self.choices.insert(0, MANUAL_SELECT_ENTRY)

    def _build_select_choices(self, choices=None):
        if not choices:
            return {"choices": []}
        if isinstance(choices, str):
            return {"choices": [{"label": "{}".format(choices)}]}
        if isinstance(choices, list):
            return {"choices": choices}
        if isinstance(choices, dict):
            returned_choices = []
            for choice_key in choices:
                returned_choices.append({
                    "label": choice_key,
                    "value": choices.get(choice_key)
                })
            return returned_choices

    def text_message(self, text_message):
        return self._build_select_choices(text_message)

    def to_dss(self):
        return self._build_select_choices(self.choices)
