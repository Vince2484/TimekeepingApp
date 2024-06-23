This is just the base framework for a timekeeping app.
In order to make use of this, the user needs to take a few steps first.

You need to go to googleapi and create a service account, then create a key.
When saving the key, save it as a json file named credential.
It is then necessary to create a google drive folder that will contain the timekeeping records.
Within that folder, create folders for each office that you wish to implement this and name them
as you wish.
Within each of those folders you will need to create another folder to hold the images that will be taken every time someone uses the app.

Now it is time to edit the python file.
At the line where folderdest is initialized, you will need to create an entry for every folder that you created. This entry should
contain the folderid of the folder, which is the series of characters at the end of the google drive folder address.
You will need to do the same thing for the image folders on the line where imagedests is initialized.

Next you will have to edit where self.box.addItem() to add each option for every office by adding more addItem() lines
and changing the string to something that would make sense to your employees.
After you edit the switch cases so that it matches the folderdest and imagedest with the selected option.

After all this is done, the TimekeeperApp is ready to run. Put the credential.json file in the same folder as the TimekeeperApp.py file,
run it and you will now have spreadsheets and images of your employees everytime they checkin and checkout of the office.
