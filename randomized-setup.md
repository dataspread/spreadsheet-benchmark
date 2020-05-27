# Running Randomized Scripts

For one of the methods we used on Google Sheets, we decided to run each trial of an experiment using a fresh copy of the spreadsheet and each trial would be on a random dataset size. Running the same experiment multiple trials in a row or running datasets in the same order could lead to caching and inconsistent results.

This is difficult to do with Google Scripts because each script execution has a timeout of 6 or 30 minutes depending on your G Suite tier. As the solution, we used the Google Scripts API in combination with a Python script to run all the trials we needed for an experiment in one script without running into timeouts. These are the instructions for this method.

# Creating a Google Apps Script API

## Create a Google Sheets Script

For the desired experiment, create a new project in the Google Apps Script Home ([https://script.google.com/home](https://script.google.com/home)) and copy the contents from the file that ends in `_randomized.gs` into the script. Fill in the script with the details specific to your dataset.

## Linking your Project with Google Cloud Platform

1. Go to your Google Cloud Platform ([https://console.cloud.google.com/home](https://console.cloud.google.com/home)) and create a new project. Each project will correspond to a single script, so you can give it the same name as your Apps Script. 
2. Go to the dashboard of the project you created.
3. Take note of the project number under **Project Info**.
4. Go to the Apps Script we are working with and click **Resources** > **Cloud Platform Project...** Input the project number from step 3 into the input under **Change Project** and click **Set Project**. If you are presented with `You cannot switch to a project without a configured OAuth consent screen...` , continue to step 5. Otherwise we're done with switching the project.
5. Click on the link to configure an OAuth consent screen. On the new screen, select **Internal** User type and click the **Create** button. Fill in the application name, it can just be the same as your Cloud Platform Project name, and then you can leave everything else and finish creation.
6. After completing the OAuth consent screen configuration, search for **Apps Script API** in the search bar. Once on the API Overview page, enable the API and go to the **Credentials** tab.
7. Click on **+ Create Credentials > Help me choose.** This will take you to this page. Under **Which API are you using?**, select **Apps Script API**. Under **Where will you be calling the API from?**, select **Other UI.** Click **What credentials do I need?**. Use any name for OAuth Client ID and continue to the next section. Download the credentials and keep it organized, as it will be used specifically for the project/script you downloaded it from.
8. Go back to the script page and click on **Set Project** again. This time it will ask you to confirm changes. Once confirmed, our Apps Script Project is finally switched to a user-managed project!

> The end goal for our script is to be able to execute it using the Apps Script API. At the moment, if you click on **Publish** > **Deploy as API Executable**, you will likely see the message: `You're using an Apps Script-managed Cloud Platform project. In order to publish, you'll need to switch to a user-managed Cloud Platform project for this script.` We need to switch our project from being Apps script-managed to user-managed.

## Deploying the Script as an API Executable

1. On your Apps Script, click **Publish** > **Deploy as API Executable.** Be sure to complete the previous part before working on this part.
2. On this screen, you'll be able to create new versions of the API or select the version that is executed when the API is called. Each version acts like a snapshot of the script whenever the version is created. The name of each version should be descriptive or noted so that it it's easy to remember the differences between the API versions.
3. Once you've created a version, your API executable is live and you can call it through a script. Take note of the **Current API ID** that pops up as you'll need it to call the API. This API ID is independent of the API executable version.

# Setting up the Python Script

## Create the Python script

1. Copy the python script template we provided at **TODO.** and replace it with the details specific to your project. The file structure we used was to have a separate directory for each project.
2. In the script, replace the value of **API_ID** with the **Current API ID** from the previous section.
3. Move the downloaded credentials for the project to the same directory as the python script. Our file was named **client_id.json**. Whatever the credentials file name is, be sure the flow variable uses the file path that matches the credentials file.
4. In the request object, replace the value corresponding to the function keyword with the name of the function you want to call in your Apps Script. 

## Run the Python Script

1. After completing all the previous sections, you should be able to run the python script and call the API corresponding to your Apps Script by calling **python script.py** in your command line i.e. Terminal.
2. If any errors pop up, they will often show up on the Apps Script page. If the project you are running requires more scopes than we have specified, you will need to replace the SCOPES variable in the python file with what