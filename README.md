# M365 Inactive User Management

A PowerShell script for automated management of inactive Microsoft 365 user accounts across multiple tenants using Microsoft Graph API.

## Features

- Automatically identifies and disables inactive user accounts based on configurable thresholds
- Supports multi-tenant environments through delegated administration
- Excludes specified users through group membership
- Generates detailed activity reports
- Sends email notifications with execution results
- Handles pagination for large user sets
- Includes comprehensive error handling and logging

## Prerequisites

- Powershell version > 7.2 
- Microsoft 365 environment with delegated administration rights and app registrations for the main app
- Azure AD Application with appropriate permissions

## Configuration
A full writeup will follow on how to set up the app registrations.

## Usage

Run the script with required parameters (or use a key ault which is MUCH safer)

.\DisableInactiveUsers.ps1 `
    -ApplicationID "your-app-id" `
    -ApplicationSecret "your-app-secret" `
    -MainTenant "yourtenant.onmicrosoft.com" `
    -ExceptionGroupName "InactiveUserExceptions"

## Customization

The script includes several configurable thresholds:

- Inactivity period (default: 90 days)
- New account grace period (default: 30 days)
- Email recipients for reports
- Specific tenants to be ignored

Modify these values in the script according to your organizational requirements.

## Security Considerations

- Store credentials securely (consider using Azure Key Vault)
- Review and adjust permissions as needed
- Monitor script execution logs
- Regularly review exception group membership
- Test in a non-production environment first

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

## Acknowledgments

- Microsoft Graph API documentation
- PowerShell community
- Kelvin Tegelaar for his work and insight into the multitenant management shared on his site https://cyberdrain.com

## Changelog

### Version 1.0 (2025-01-01)
- Initial release
- Basic functionality for identifying and disabling inactive users
- Email reporting feature
- Multi-tenant support
