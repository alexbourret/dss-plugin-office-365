from office365_commons import get_credentials_from_config
from office365_client import Office365Session


def build_select_choices(choices=None):
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


def do(payload, config, plugin_config, inputs):
    sharepoint_oauth = config.get("sharepoint_oauth")
    if not sharepoint_oauth:
        return build_select_choices("Select a preset")

    auth_token = get_credentials_from_config(config)
    parameter_name = payload.get('parameterName')
    sharepoint_site_id = config.get("sharepoint_site_id")
    search_space = config.get("search_space")
    choices = []
    if parameter_name == "sharepoint_site_id":
        session = Office365Session(access_token=auth_token)
        for sharepoint_site_id in session.get_next_item(
                url="https://graph.microsoft.com/v1.0/sites",
                params={"search": "*"}
        ):
            choices.append(
                {
                    "label": sharepoint_site_id.get("displayName"),
                    "value": sharepoint_site_id.get("id")
                }
            )

    if parameter_name == "sharepoint_list_id":
        if not sharepoint_site_id:
            return build_select_choices("Select a site")
        session = Office365Session(access_token=auth_token)
        for sharepoint_list in session.get_next_item(
            url="https://graph.microsoft.com/v1.0/sites/{}/lists".format(
                sharepoint_site_id
            )
        ):
            choices.append(
                {
                    "label": sharepoint_list.get("displayName"),
                    "value": sharepoint_list.get("id")
                }
            )
    if parameter_name == "sharepoint_drive_id":
        if search_space == "site" and not sharepoint_site_id:
            return build_select_choices("Select a site")
        session = Office365Session(access_token=auth_token)
        for sharepoint_list in session.get_next_item(
            url="https://graph.microsoft.com/v1.0/sites/{}/drives".format(
                sharepoint_site_id
            )
        ):
            choices.append(
                {
                    "label": sharepoint_list.get("name"),
                    "value": sharepoint_list.get("id")
                }
            )
    return build_select_choices(choices)
