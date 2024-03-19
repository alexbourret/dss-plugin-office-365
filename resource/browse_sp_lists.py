from office365_commons import get_credentials_from_config, DSSSelectorChoices
from office365_client import Office365Session


def do(payload, config, plugin_config, inputs):
    choices = DSSSelectorChoices()
    sharepoint_oauth = config.get("sharepoint_oauth")
    if not sharepoint_oauth:
        return choices.text_message("⚠ Select a preset")

    auth_token = get_credentials_from_config(config)
    parameter_name = payload.get('parameterName')
    sharepoint_site_id = config.get("sharepoint_site_id")
    sharepoint_site_overwrite = None
    if sharepoint_site_id == "dku_manual_select":
        sharepoint_site_overwrite = config.get("sharepoint_site_overwrite")

    search_space = config.get("search_space")

    if parameter_name == "sharepoint_site_id":
        session = Office365Session(access_token=auth_token)
        for sharepoint_site_id in session.get_next_site():
            choices.append(sharepoint_site_id.get("displayName"), sharepoint_site_id.get("id"))
        choices.append_manual_select()

    if parameter_name == "sharepoint_list_id":
        if (not sharepoint_site_id) or (sharepoint_site_id == "dku_manual_select" and not sharepoint_site_overwrite):
            return choices.text_message("⚠ Select a site")

        session = Office365Session(access_token=auth_token)
        if sharepoint_site_id == "dku_manual_select" and sharepoint_site_overwrite:
            sharepoint_site_id = session.get_site_id(sharepoint_site_overwrite)
            if not sharepoint_site_id:
                return choices.text_message("⚠ Site not found")

        site = session.get_site(sharepoint_site_id)
        for sharepoint_list in site.get_next_list():
            choices.append(sharepoint_list.get("displayName"), sharepoint_list.get("id"))
        choices.append_manual_select()

    if parameter_name == "sharepoint_drive_id":
        if search_space == "site" and (not sharepoint_site_id or (sharepoint_site_id == "dku_manual_select" and not sharepoint_site_overwrite)):
            return choices.text_message("⚠ Select a site")

        session = Office365Session(access_token=auth_token)
        if sharepoint_site_id == "dku_manual_select" and sharepoint_site_overwrite:
            sharepoint_site_id = session.get_site_id(sharepoint_site_overwrite)
            if not sharepoint_site_id:
                return choices.text_message("⚠ Site not found")
        site = session.get_site(sharepoint_site_id)
        for sharepoint_drive in site.get_next_drive():
            choices.append(sharepoint_drive.get("name"), sharepoint_drive.get("id"))
        choices.append_manual_select()

    if parameter_name == "plan_id":
        group_id = config.get("group_id", None)
        if not group_id:
            return choices.text_message("⚠ Set the group ID")

        session = Office365Session(access_token=auth_token)
        for plan in session.get_next_plan(group_id):
            choices.append(plan.get("title"), plan.get("id"))

    return choices.to_dss()
