O365APIPhone
============

Demo of O365 _api using Web Authentication Broker, does not use O365 client APIs yet due to lack of support

Couple of things to note here:

Follow these steps to get the app running:

1. Configure all of the class variables in the following classes:

  a. MainPage.xaml.cs
  
  b. SharedProject/O365APIS/O365APISites
  
2. replace tenant with your O365 domain you are using such as https://<tenant>.onmicrosoft.com or https://<tenant>.sharepoint.com

3. see the nuget packages I have installed and make sure they are installed and available in your environment, (in packages folder)

4. When running you will need to click on the refresh button in the bottom app bar to activate the data retrieval

5. you will select the document you want approved in the approve pivot view

6. and click on the accept button in the bottom app bar

Couple of warnings/bugs:
1. Although I put the trick to go from single to multi-select when you select an item, it does not do batch yet, only singular or the first on in the array you select.

2. After the first approve if you try to approve another document you get an auth error, something wrong with the token reuse/cache,need to figure that out, I think it is a simple issue.

3. Need to implement a batch call, trying to figure out the best way to do that now.

Enjoy and feel free to reach out to me for questions/issues.

twitter- @jasonmcnutt

blog - http://jmmcnj.blogspot.com/

