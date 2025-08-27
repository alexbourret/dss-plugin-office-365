class DSSConstants(object):
    APPLICATION_JSON = "application/json"
    APPLICATION_JSON_NOMETADATA = "application/json;odata=nometadata"
    AUTH_LOGIN = "login"
    AUTH_OAUTH = "oauth"
    AUTH_SITE_APP = "site-app-permissions"
    CHILDREN = 'children'
    DATE_FORMAT = "%Y-%m-%dT%H:%M:%S.%fZ"
    DEFAULT_BATCH_SIZE = 19
    DIRECTORY = 'directory'
    EXISTS = 'exists'
    FALLBACK_TYPE = "string"
    FULL_PATH = 'fullPath'
    GZIP_HEADERS = {
        "Content-Encoding": "gzip",
        "Accept-Encoding": "gzip"
    }
    IS_DIRECTORY = 'isDirectory'
    JSON_HEADERS = {
        "Content-Type": APPLICATION_JSON,
        "Accept": APPLICATION_JSON
    }
    LAST_MODIFIED = 'lastModified'
    LOGIN_DETAILS = {
        "sharepoint_tenant": "The tenant name is missing",
        "sharepoint_site": "The site name is missing",
        "sharepoint_username": "The account's username is missing",
        "sharepoint_password": "The account's password is missing"
    }
    OAUTH_DETAILS = {
        "sharepoint_tenant": "The tenant name is missing",
        "sharepoint_site": "The site name is missing",
        "sharepoint_oauth": "The access token is missing"
    }
    PATH = 'path'
    PLUGIN_VERSION = "0.0.2"
    SECRET_PARAMETERS_KEYS = ["Authorization", "sharepoint_username", "sharepoint_password", "client_secret", "sharepoint_oauth"]
    SITE_APP_DETAILS = {
        "sharepoint_tenant": "The tenant name is missing",
        "sharepoint_site": "The site name is missing",
        "tenant_id": "The tenant ID is missing. See documentation on how to obtain this information",
        "client_id": "The client ID is missing",
        "client_secret": "The client secret is missing"
    }
    SIZE = 'size'
    TYPES = {
        "string": "Text",
        "map": "Note",
        "array": "Note",
        "object": "Note",
        "double": "Number",
        "float": "Number",
        "int": "Integer",
        "bigint": "Integer",
        "smallint": "Integer",
        "tinyint": "Integer",
        "date": "DateTime"
    }
