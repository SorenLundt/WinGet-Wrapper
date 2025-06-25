# Azure AD Application Configuration for WinGet-Wrapper

## Background

The WinGet-Wrapper tool uses Microsoft Graph API to interact with Intune. By default, it uses a built-in application ID, but due to recent Microsoft infrastructure changes and security policies, it's recommended to create your own Azure AD application registration. This ensures:

1. Better security control over the application
2. Avoidance of potential throttling issues
3. Clear audit trails in your Azure environment
4. Prevention of authentication issues related to Microsoft's first-party app verification changes

## Creating Your Azure AD Application

### Step 1: Create the Application Registration
1. Log in to the Azure Portal (portal.azure.com)
2. Navigate to Azure Active Directory → App registrations
3. Click "New registration"
4. Configure the following:
   - Name: "WinGet-Wrapper-App" (or your preferred name)
   - Supported account types: "Accounts in this organizational directory only"
   - Click "Register"
5. After creation, note down the "Application (client) ID" - you'll need this later

### Step 2: Configure Authentication
1. In your app registration, go to "Authentication" in the left menu
2. Click "Add a platform"
3. Select "Mobile and desktop applications"
4. Check the box for "https://login.microsoftonline.com/common/oauth2/nativeclient"
5. Click "Configure"

This configuration is crucial because the PowerShell scripts use interactive authentication, which requires a proper redirect URI.

### Step 3: Configure API Permissions
1. Go to "API permissions" in the left menu
2. Click "Add a permission"
3. Select "Microsoft Graph"
4. Choose "Application permissions"
5. Search for and select "DeviceManagementApps.ReadWrite.All"
6. Click "Add permissions"
7. Click "Grant admin consent" and confirm

## Updating the Scripts

You need to update the ClientID in the following files:

### Option 1: Modify the Script Directly
Update `WinGet-WrapperImportFromCSV.ps1`:
```powershell
#ClientID to connect to MSGraph/InTune with Connect-MSIntuneGraph
[Parameter(Mandatory = $False)]
[string]$ClientID = "your-application-id-here"
```

### Option 2: Pass ClientID as Parameter
Run the script with your ClientID:
```powershell
.\WinGet-WrapperImportFromCSV.ps1 -TenantID "yourtenant.onmicrosoft.com" -ClientID "your-application-id" -csvFile "your-csv-file.csv"
```

## Troubleshooting

### Common Issues

1. "No reply address is registered for the application"
   - **Cause**: Missing redirect URI configuration
   - **Solution**: Follow Step 2 in the configuration process

2. "Application is not authorized to perform this operation"
   - **Cause**: Missing or unauthorized API permissions
   - **Solution**: Ensure Step 3 is completed and admin consent is granted

3. "AADSTS700016" or "AADSTS90099"
   - **Cause**: Application not properly authorized in tenant
   - **Solution**: Ensure admin consent is granted and the account has proper roles (Global Admin or Intune Administrator)

## Best Practices

1. **Security**:
   - Regularly review and audit application permissions
   - Use separate applications for development and production
   - Follow the principle of least privilege when assigning permissions

2. **Maintenance**:
   - Document your application ID and configuration
   - Regularly review and update permissions as needed
   - Monitor application usage through Azure AD audit logs

## Required Azure AD Roles

The user account running the scripts needs one of these roles:
- Intune Administrator
- Global Administrator

## Additional Resources

- [Microsoft Graph Permissions Reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
- [Azure AD Application Registration](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Microsoft Graph PowerShell Authentication](https://docs.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands) 



see https://github.com/SorenLundt/WinGet-Wrapper/issues/21
see https://github.com/SorenLundt/WinGet-Wrapper/pull/23