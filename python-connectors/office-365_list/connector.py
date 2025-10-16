from dataiku.connector import Connector
from office365_commons import RecordsLimit, get_credentials_from_config, LookupList
from office365_client import Office365Session, Office365ListWriter
from safe_logger import SafeLogger
from dss_constants import DSSConstants


logger = SafeLogger("office-365 plugin", DSSConstants.SECRET_PARAMETERS_KEYS)


class Office365ListConnector(Connector):

    def __init__(self, config, plugin_config):
        Connector.__init__(self, config, plugin_config)

        logger.info("Office-365 plugin v{} list connector".format(DSSConstants.PLUGIN_VERSION))
        logger.info("config={}".format(logger.filter_secrets(config)))

        self.auth_token = get_credentials_from_config(config)
        self.sharepoint_site_id = config.get("sharepoint_site_id")
        self.sharepoint_list_id = config.get("sharepoint_list_id")
        if not self.sharepoint_site_id:
            raise Exception("A SharePoint site must be selected")
        session = Office365Session(access_token=self.auth_token)

        if self.sharepoint_site_id == "dku_manual_select":
            sharepoint_site_overwrite = config.get("sharepoint_site_overwrite")
            self.sharepoint_site_id = session.get_site_id(sharepoint_site_overwrite)

        site = session.get_site(self.sharepoint_site_id)

        if self.sharepoint_list_id == "dku_manual_select":
            sharepoint_list_title = config.get("sharepoint_list_title")
            self.sharepoint_list_id = site.get_list_id(sharepoint_list_title)

        if not self.sharepoint_list_id:
            raise Exception("A SharePoint list must be selected")

        self.list = site.get_list(self.sharepoint_list_id)
        self.must_see_columns = config.get("must_see_columns", [])

    def get_read_schema(self):
        """
        Returns the schema that this connector generates when returning rows.

        The returned schema may be None if the schema is not known in advance.
        In that case, the dataset schema will be infered from the first rows.

        If you do provide a schema here, all columns defined in the schema
        will always be present in the output (with None value),
        even if you don't provide a value in generate_rows

        The schema must be a dict, with a single key: "columns", containing an array of
        {'name':name, 'type' : type}.

        Example:
            return {"columns" : [ {"name": "col1", "type" : "string"}, {"name" :"col2", "type" : "float"}]}

        Supported types are: string, int, bigint, float, double, date, boolean
        """

        # In this example, we don't specify a schema here, so DSS will infer the schema
        # from the columns actually returned by the generate_rows method
        return None

    def generate_rows(self, dataset_schema=None, dataset_partitioning=None,
                      partition_id=None, records_limit=-1):
        limit = RecordsLimit(records_limit)
        column_display_name = {}
        lookup_list = LookupList(must_see_columns=self.must_see_columns)
        for column in self.list.get_columns():
            column_display_name[column.get("name")] = column.get("displayName")
            lookup_list.append(column)

        for row in self.list.get_next_row(
            select_list=lookup_list.get_select(),
        ):
            fields = row.get("fields", {})
            row_with_real_names = {}
            for name in fields:
                selected_column_display_name = column_display_name.get(name)
                if selected_column_display_name:
                    row_with_real_names[selected_column_display_name] = fields.get(name)
            yield row_with_real_names
            if limit.is_reached():
                return

    def get_writer(self, dataset_schema=None, dataset_partitioning=None,
                   partition_id=None):
        self.list.delete_all_rows()
        sharepoint_columns = []
        for sharepoint_column in self.list.get_columns():
            sharepoint_columns.append(
                {
                    "name": sharepoint_column.get("name"),
                    "display_name": sharepoint_column.get("displayName"),
                    "type": sharepoint_to_dss_type(sharepoint_column)
                }
            )
        missing_sharepoint_columns = compute_missing_sharepoint_columns(
            dataset_schema.get("columns"),
            sharepoint_columns
        )
        for missing_sharepoint_column in missing_sharepoint_columns:
            logger.info("Adding column '{}' of type {}".format(
                    missing_sharepoint_column.get("name"), missing_sharepoint_column.get("type")
                )
            )
            self.list.add_column(
                missing_sharepoint_column.get("name"),
                missing_sharepoint_column.get("type"),
                description="Created by DSS Office-365 plugin"
            )
        return Office365ListWriter(
            self.list, dataset_schema
        )

    def get_partitioning(self):
        """
        Return the partitioning schema that the connector defines.
        """
        raise NotImplementedError

    def list_partitions(self, partitioning):
        """Return the list of partitions for the partitioning scheme
        passed as parameter"""
        return []

    def partition_exists(self, partitioning, partition_id):
        """Return whether the partition passed as parameter exists

        Implementation is only required if the corresponding flag is set to True
        in the connector definition
        """
        raise NotImplementedError

    def get_records_count(self, partitioning=None, partition_id=None):
        """
        Returns the count of records for the dataset (or a partition).

        Implementation is only required if the corresponding flag is set to True
        in the connector definition
        """
        raise NotImplementedError


def sharepoint_to_dss_type(sharepoint_column):
    if "text" in sharepoint_column:
        return "string"
    if "number" in sharepoint_column:
        return "float"
    return "string"


def compute_missing_sharepoint_columns(dss_columns, sharepoint_columns):
    missing_sharepoint_columns = []
    for dss_column in dss_columns:
        dss_column_name = dss_column.get("name")
        dss_column_type = dss_column.get("type")
        column_found = False
        for sharepoint_column in sharepoint_columns:
            sharepoint_column_name = sharepoint_column.get("name")
            if sharepoint_column_name == dss_column_name:
                column_found = True
        if column_found:
            logger.info("Column '{}' found on SharePoint, so skipping creation".format(dss_column_name))
            continue
        else:
            missing_sharepoint_columns.append(
                {
                    "name": dss_column_name,
                    "type": dss_column_type
                }
            )
    return missing_sharepoint_columns
