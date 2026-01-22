Squad Health Check
==================

The Squad Health Check is a way for teams to gauge their perception of productivity, performance, and purpose. Originally developed at Spotify, the check enables teams to identify ways to improve their processes, individual skills, and overall work quality-of-life across eleven dimensions of team sentiment. Spotify's engineers have refined the check over the years, and we recommend reading their [original 2014 blog post](https://engineering.atspotify.com/2014/09/squad-health-check-model) and [follow-up 2023 blog post](https://engineering.atspotify.com/2023/03/getting-more-from-your-team-health-checks) prior to using the tool. Here, we have adapted the questions from a [distillation of the Spotify process by TeamRetro](https://www.teamretro.com/health-checks/squad-health-check) used under a Creative Commons Attribution-ShareAlike license.

The Squad Health Check Google Sheets™ editor add-on automates the generation of a survey for teams to measure _perception_ and _trend_ sentiments in each dimension. The add-on adds sheets into an existing Google Sheets™ spreadsheet to provide a template sheet for generating a Google Forms™ survey and compute and chart sheets for aggregating survey responses from team members.

## Installation

You will need to [sign up for a Google™ account](https://accounts.google.com/signup) to use the tool if you do not already have an account. You will also need access to Google Sheets™ and permissions to install add-ons from the Google Workspace™ Marketplace; if you are using a work or school account, you may need to ask your administrator to enable access.

* Install the [Squad Health Check add-on](https://workspace.google.com/marketplace/app/squad_health_check/746334686635) from the Google Workspace Marketplace™.

* Open a Google Sheets™ spreadsheet. We recommend using a new worksheet for hosting the templates and survey responses.

* Initialize the spreadsheet by opening “Extensions → Squad Health Check → Install templates…”. The add-on will create “Survey template” and “Compute” template sheets, and protect them against accidental editing.

* Optionally, initialize the chart sheets by opening “Extensions → Squad Health Check → Install charts…”. The graphs in the chart sheets allow you to observe changes in perception and trend sentiments across different Squad Health Check dates.

### Google OAuth permissions

When installing the add-on, you will see a dialog asking you to grant the following permissions to the plug-in:

* View and manage your forms in Google Drive™: This permission allows the plug-in to create the new form.  
* View and manage spreadsheets that this application has been installed in: This permission allows the add-on to update the spreadsheet to collect survey responses.  
* Connect to an external service: This permission allows the add-on to pull the icons from the survey template sheet and embed them in the form. The add-on cannot retrieve images from a sheet cell directly but must instead retrieve them using icon URL shown in the “Icon URL” column of the survey template sheet.

## Survey form generation

* Run the Editor add-on by opening “Extensions → Squad Health Check → Generate survey form…”. A dialog box will appear asking you to specify a date in yyyy-MM-dd format (such as the date of a meeting at which you plan to discuss the results). Fill in the date and select “OK”.

* Wait for the add-on to finish running. Once the add-on finishes, you will have a survey form at the root of your drive (“My Drive”) named “Squad Health Check [yyyy-MM-dd]” and a new survey response sheet in your copy of the spreadsheet named “Squad Health Check [yyyy-MM-dd]”.

* Open the new form, and click the “Publish” button.

## Survey response collection

The response sheet will update automatically as team members complete the survey form. You can view the average sentiment scores (higher average numbers are “better”) in the “Compute” sheet, and look at graphs over time of perception and trend in the chart sheets (if you installed the chart sheets).

## Acknowledgments

* Squad Health Check questions and dimension icons from the TeamRetro [Spotify process](https://www.teamretro.com/health-checks/squad-health-check) used under a Creative Commons Attribution-ShareAlike license

* Stethoscope icon designed by [Freepik at Flaticon](https://www.flaticon.com/authors/freepik)
