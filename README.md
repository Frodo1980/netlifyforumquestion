# Repository in addition to this [question](https://answers.netlify.com/t/express-server-with-netlify-functions/55408) on Netlify-support-forum.

Prerequisites: install [netlify-CLI](https://docs.netlify.com/cli/get-started/) for local Dev-Server

For reproducing follow these steps:

1. clone branch "Express" for the local running Express-server
2. clone branch "function" for the wrapped server
3. in each cloned repository run `npm run start`

## Test Express-Server for serving the customized file

in your local browser call http://localhost:3000/.netlify/functions/api?A13=TextForCellA13&B15=ThisGoesToB15

expected behavior: you get a xlsx.-file with all of the content your given xlsx.-file had (here: file_example.xlsx), additionally the values given as URL-parameters are included in the specified cells.

It works! 😀

## Test lambda-function which should do the same

Calling the `npm run start` runs a development-server. After starting the server you see in the console the adress where you can reach the server.

Call this adress and put the adress to call the function (/.netlify/functions/api?A2=test&A13=Hello) after the server-adress:

You will get a corrupted xlsx.-file which cannot be opened.

But why? Where is the difference between the express-server alone and the lambda-function which should do the same?

