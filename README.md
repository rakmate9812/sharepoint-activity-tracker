# sharepoint-activity-tracker
Sharepoint activity tracker console app in .NET 7.0

This console application was made for tracking all kinds of sharepoint activities made in a Microsoft tenant.

Usage:

First you have to create an AAD application inside your tenant, create a client secret for your app, and give them the necessary permissions (Office 365 MAnagement API -> ActivityFeed.Read). 
Then get your Client ID, Tenant ID, and Client Secret.

Open the Visual Studio solution, download the necessities from NuGet.

Paste your credentials in the config.cs, run the application. A JSON file containing the logs from the last day should be saved in the solution folder. You can configure the code as you wish,
the startTime, endTime, output folder might not satisfy your needs. Keep in mind, that the ManagementAPI that I'm using can get the logs not older than a week, and only can
retrieve a 24 hour session at a time.

If you have any question let me know, I'm pretty new in the world of .NET and Visual Studio. This is a self-made project for code practicing purposes.
 

