# Plugin installation

- Using GIT:
    - Navigate to > Plugins > Add Plugin > Fetch from git repository, set Repository URL to `https://github.com/alexbourret/dss-plugin-office-365.git` 
- Using the zip file
    - Download the zip file from https://github.com/alexbourret/dss-plugin-office-365/releases/tag/v0.0.2
    - Navigate to > Plugins > Add Plugin > Upload > Browse and point to the downloaded file

# Testing the fix

## With a SharePoint list

- In a DSS flow, add a SharePoint list by clicking on +Dataset > Office 365 > SharePoint Lists
- In **Type of authentication** select *DSS connection*
- In **DSS connection** select a valid **SharePoint** connection
- Check that **Site** now
    - indicates *- Nothing selected -*
    - contains a list of accessible sites
- Select the "*✍️ Enter manually*" option
- In the **Site name** box, paste the full URL of the site you want to access. For instance `https://my-corp.sharepoint.com/sites/mysite` or `https://my-corp.sharepoint.com/sites/mysite/Shared%20Documents/Forms/AllItems.aspx`. Note that for this test to be conclusive, **this site should not be visible using the usual native connection**. 
- Check that the **List** selector is now filled in all the lists available on the selected site.

## With a SharePoint shared folder

- In a DSS flow, add a folder SharePoint list by clicking on +Dataset > Folder and setting **Store into** to *plugin office-365 > Access Office 365 files*
- Click on Settings
- In **Type of authentication** select *DSS connection*
- In **DSS connection** select a valid **SharePoint** connection
- Check that **Site** now
    - indicates *- Nothing selected -*
    - contains a list of accessible sites
- Select the "*✍️ Enter manually*" option
- In the **Site name** box, paste the full URL of the site you want to access. For instance `https://my-corp.sharepoint.com/sites/mysite` or `https://my-corp.sharepoint.com/sites/mysite/Shared%20Documents/Forms/AllItems.aspx`. Note that for this test to be conclusive, **this site should not be visible using the usual native connection**. 
- Check that the **Drive** selector is now filled in all the drives available on the selected site.

# Note

This is a test plugin. It is functionnal but not reviewed, so it is not advise to use it on production flows.
