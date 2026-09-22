Squad Health Check
==================

Squad Health Check is an [open-source editor add-on for Google Sheets™](https://github.com/tmwong-com/squad-health-check) for running and tracking team retrospective surveys based on the [Spotify Squad Health Check model](https://engineering.atspotify.com/2014/09/squad-health-check-model). The add-on automates the generation of Google Forms™ surveys, collection of survey responses, and visualization of team sentiment over time within the Google Workspace™ ecosystem.

The Squad Health Check evaluates team sentiment across eleven dimensions covering areas such as software quality, speed, teamwork, and mission clarity. Teams can customize the survey questions and parameters to better reflect their own processes and organization.

<a href="https://workspace.google.com/marketplace/app/squad_health_check/746334686635?pann=b">
  <img
    alt="Get it from the Google Workspace Marketplace"
    src="https://workspace.google.com/static/img/marketplace/en/gwmBadge.svg"
    height="68">
</a>

## Key features

* Generates Google Forms™ surveys directly from Google Sheets™.
* Automatically collects and aggregates survey responses in the spreadsheet.
* Provides charts for tracking changes in team sentiment over time.
* Includes eleven Squad Health Check dimensions covering software quality, speed, teamwork, and mission clarity.
* Allows teams to customize survey questions and parameters.

## Installation

* Install the [Squad Health Check add-on](https://workspace.google.com/marketplace/app/squad_health_check/746334686635?pann=b) from the Google Workspace Marketplace™.
* Open a Google Sheets™ spreadsheet. We recommend using a new worksheet for hosting the templates and survey responses.
* Initialize the Squad Health Check survey template and compute sheets with “Extensions → Squad Health Check → Install templates…”. The add-on creates “Survey template” and “Compute” template sheets and protects them against accidental editing.
* Optionally, initialize chart sheets with “Extensions → Squad Health Check → Install charts”. The graphs in the chart sheets capture changes in perception and trend sentiments across different Squad Health Check dates.

[Sign up for a Google™ account](https://accounts.google.com/signup) to use the tool if you do not already have an account. Ensure that you have access to Google Sheets™ and permission to install add-ons from the Google Workspace™ Marketplace. If you are using a work or school account, ask your administrator to enable access if necessary.

### Google OAuth permissions

When installing the add-on, grant the following permissions to the plug-in:

* View and manage your forms in Google Drive™: This permission allows the plug-in to create the new form.
* View and manage spreadsheets that this application has been installed in: This permission allows the add-on to update the spreadsheet to collect survey responses.  
* Connect to an external service: This permission allows the add-on to pull the icons from the survey template sheet and embed them in the form.

## Survey form generation

* Generate a new Squad Health Check survey form with “Extensions → Squad Health Check → Generate survey form…”. In the dialog box, specify a date in yyyy-MM-dd format (such as the date of a meeting at which you plan to discuss the results) and select “OK”.
* Wait for the add-on to finish running. Once the add-on finishes, you will have a survey form at the root of your drive (“My Drive”) named “Squad Health Check [yyyy-MM-dd]” and a new survey response sheet in your copy of the spreadsheet named “Squad Health Check [yyyy-MM-dd]”.
* Open the new form and click the “Publish” button.

## Survey response collection

The response sheet updates automatically as team members complete the survey form. View the average sentiment scores (higher average numbers are “better”) in the “Compute” sheet and graphs of perception and trend over time in the chart sheets (if you installed the chart sheets).

## Survey response sheet removal

If you rename or remove a survey response sheet (i.e., the sheet linked to a Squad Health Check survey form), use “Extensions → Squad Health Check → Update compute” to remove the outdated references to the response sheet from the compute sheet.

## Additional help

For further questions about the Squad Health Check add-on or our experiences with running surveys, contact us by e-mail at [tmwong@tmwong.com](mailto:tmwong@tmwong.com).

If you discover issues with the tool, report them through the [Squad Health Check "Issues" page](https://github.com/tmwong-com/squad-health-check/issues) on GitHub.

## Acknowledgments

* Google, Google Sheets, Google Forms, Google Drive, and Google Workspace are trademarks of Google LLC.
* Icons made by [Magnific](https://www.flaticon.com/authors/magnific) from [www.flaticon.com](https://www.flaticon.com/)
* Squad Health Check concept from Spotify in a [2014 blog post](https://engineering.atspotify.com/2014/09/squad-health-check-model) and a [follow-up 2023 blog post](https://engineering.atspotify.com/2023/03/getting-more-from-your-team-health-checks)
* Squad Health Check questions from the TeamRetro [Spotify process](https://www.teamretro.com/health-checks/squad-health-check) used under a Creative Commons Attribution-ShareAlike license
