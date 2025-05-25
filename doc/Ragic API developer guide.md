-   [1. How to use an Ragic HTTP API](https://www.ragic.com/intl/en/doc-api?onepage#onepage0)
    1.  [1.1. Ragic REST Web Service Interface](https://www.ragic.com/intl/en/doc-api?onepage#onepage0)
        
    2.  [1.2. Using cURL to access HTTP API](https://www.ragic.com/intl/en/doc-api?onepage#onepage1)
        
    3.  [1.3. Using Postman to access HTTP API](https://www.ragic.com/intl/en/doc-api?onepage#onepage35)
        
    4.  [1.4. Finding the field id for a field](https://www.ragic.com/intl/en/doc-api?onepage#onepage18)
        
    5.  [1.5. Finding API endpoints](https://www.ragic.com/intl/en/doc-api?onepage#onepage7)
        
    6.  [1.6. API Limits](https://www.ragic.com/intl/en/doc-api?onepage#onepage22)
        
-   [2. Working with Ragic API](https://www.ragic.com/intl/en/doc-api?onepage#onepage24)
    1.  [2.1. Authentication](https://www.ragic.com/intl/en/doc-api?onepage#onepage24)
        1.  [2.1.1. HTTP Basic authentication with Ragic API Key](https://www.ragic.com/intl/en/doc-api?onepage#onepage24)
        2.  [2.1.2. Password authentication](https://www.ragic.com/intl/en/doc-api?onepage#onepage5)
    2.  [2.2. Reading](https://www.ragic.com/intl/en/doc-api?onepage#onepage8)
        1.  [2.2.1. Returned data JSON format](https://www.ragic.com/intl/en/doc-api?onepage#onepage8)
        2.  [2.2.2. Filter Conditions](https://www.ragic.com/intl/en/doc-api?onepage#onepage9)
        3.  [2.2.3. Limiting Entry Number / Paging](https://www.ragic.com/intl/en/doc-api?onepage#onepage10)
        4.  [2.2.4. Sorting and ordering](https://www.ragic.com/intl/en/doc-api?onepage#onepage11)
        5.  [2.2.5. Callback Function (JSONP)](https://www.ragic.com/intl/en/doc-api?onepage#onepage12)
        6.  [2.2.6. Field Naming](https://www.ragic.com/intl/en/doc-api?onepage#onepage13)
        7.  [2.2.7. Other GET parameters](https://www.ragic.com/intl/en/doc-api?onepage#onepage25)
        8.  [2.2.8. Retrieving uploaded files, images and e-mail attachments](https://www.ragic.com/intl/en/doc-api?onepage#onepage28)
        9.  [2.2.9. Retrieving HTML, PDF, Excel, Mail Merge and Custom Print Report of a record](https://www.ragic.com/intl/en/doc-api?onepage#onepage34)
    3.  [2.3. Writing](https://www.ragic.com/intl/en/doc-api?onepage#onepage19)
        1.  [2.3.1. Creating entries with custom HTML forms](https://www.ragic.com/intl/en/doc-api?onepage#onepage19)
        2.  [2.3.2. Creating a New Entry](https://www.ragic.com/intl/en/doc-api?onepage#onepage15)
        3.  [2.3.3. Modifying an Entry](https://www.ragic.com/intl/en/doc-api?onepage#onepage16)
        4.  [2.3.4. Create / Update Parameters](https://www.ragic.com/intl/en/doc-api?onepage#onepage26)
        5.  [2.3.5. Deleting an entry](https://www.ragic.com/intl/en/doc-api?onepage#onepage20)
        6.  [2.3.6. Upload files and images](https://www.ragic.com/intl/en/doc-api?onepage#onepage29)
        7.  [2.3.7. Comment](https://www.ragic.com/intl/en/doc-api?onepage#onepage40)
    4.  [2.4. API Request Error Response](https://www.ragic.com/intl/en/doc-api?onepage#onepage17)
        
    5.  [2.5. Sample code](https://www.ragic.com/intl/en/doc-api?onepage#onepage21)
        
    6.  [2.6. Common Questions](https://www.ragic.com/intl/en/doc-api?onepage#onepage36)
        1.  [2.6.1. Common Q&A](https://www.ragic.com/intl/en/doc-api?onepage#onepage36)
        2.  [2.6.2. Common Questions For API Parameter](https://www.ragic.com/intl/en/doc-api?onepage#onepage38)
    7.  [2.7. Mass Operation](https://www.ragic.com/intl/en/doc-api?onepage#onepage39)
        
-   [3. Using Webhook to notify you about changes](https://www.ragic.com/intl/en/doc-api?onepage#onepage32)
    1.  [3.1. What is a webhook](https://www.ragic.com/intl/en/doc-api?onepage#onepage32)
        
    2.  [3.2. Webhook on Ragic](https://www.ragic.com/intl/en/doc-api?onepage#onepage33)
        

## 1.1    Ragic REST Web Service Interface

The Ragic [REST](http://en.wikipedia.org/wiki/Representational_State_Transfer) API allows you to query any data that you have on Ragic, or to execute create / update / delete operations programmatically to integrate with your own applications.

Since the API is based on REST principles, it's very easy to write and test applications. You can use your browser to access URLs, and you can use pretty much any HTTP client in any programming language to interact with the API.

HTTP requests can be issued with two content types, form data and JSON. **However, for file uploads, form data is the only option.**

## 1.2    Using cURL to access HTTP API

You can test most GET method APIs easily by entering API endpoint URLs in your browser. For example, you can try the following URL to access customer account information on the Ragic demo:

```
<p>https://www.ragic.com/demo/sales/1?api
</p>
```

Note that you may need to modify **www** to **na3**, **ap5**, or **eu2** in your URL based on your Ragic account URL.

But it's not as easy to create POST method requests on your browser, so we recommend a tool called cURL for you to create all types of HTTP requests you want to test our API. Our document will also be using cURL commands as samples API calls, but you can also use any tool that you're familiar with to create HTTP requests and parse responses.

You can download cURL for your platform at [http://curl.haxx.se/download.html](http://curl.haxx.se/download.html) , and you can also read its full documentation at [http://curl.haxx.se/docs/manpage.html](http://curl.haxx.se/download.html) . But don't worry, we'll tell you about the necessary usages of cURL as we explain the Ragic HTTP API.

After you have downloaded cURL, you can use the following command to access the same endpoint we described above, and you should see the same output as you would in a browser.

```
<p>curl https://www.ragic.com/demo/sales/1?api
</p>
```

Please note that when using CURL, -d does not encode the content as URL string. If your string content has characters like % or & please use the option --data-urlencode instead of -d.

## 1.3    Using Postman to access HTTP API

Postman is a tool that helps to send HTTP requests without having to code. You can see Postman’s documentation [here](https://learning.postman.com/docs/getting-started/introduction/), and Postman can be downloaded [here](https://www.postman.com/downloads/).

Postman is recommended for users with little technical experience, and the fundamentals will be introduced in this tutorial.

The user interface for Postman looks like the image below:

![](Ragic%20API%20developer%20guide/file.jsp)

-   1\. The dropdown menu that lists all HTTP methods to select from
-   2\. The destination URL that the HTTP request is sent to
-   3\. The action button that sends out the HTTP request
-   4\. Attributes associated with the HTTP request

For the image above, the HTTP request is a GET request, and it is being sent to https://www.ragic.com/demo/sales/1 with a query parameter "api" that has no value attached to it.

The image below shows the list of HTTP methods from the dropdown menu.

![](Ragic%20API%20developer%20guide/file.1.jsp)

## 1.4    Finding the field id for a field

In a lot of cases you need to find a **field id** for a field. The field id is an unique number that Ragic assigns to each field that you created. It gives your program an unambiguous way to refer to a field. With this design, fields can even have the same name under a sheet and Ragic will still be able to distinguish between them.

Go to the design mode, and click on the field you want to reference to. On the left sidebar, you're going to see the **Field Name**. The **Field Id** is right under the field name.

![](Ragic%20API%20developer%20guide/file.2.jsp)

## 1.5    Finding API endpoints

You can find the API endpoint for your Ragic form or entry by passing **api** as a query string parameter to specify this is an API request.

```
<p>https://www.ragic.com/{ACCOUNT_NAME}&gt;/{TAB_FOLDER}&gt;/{SHEET_INDEX}/{RECORD_ID}?v=3&amp;api
</p>
```

Please Note:

-   It is required to modify **www** to **na3**, **ap5**, or **eu2** in the API URL based on your Ragic database account URL.
-   The API version is specified by the URL parameter, **v**. e.g. **v=3** specifies version 3 of the API.

It's a good practice to specify the API version every time you send your request to ensure a correct version. If the version is not specified, the latest version will be used, but unexpected changes to the API may cause problems in your application.

For example if you usually access your Ragic form using the following URL:

```
<p>https://www.ragic.com/demo/sales/1
</p>
```

Its HTTP API endpoint URL would be:

```
<p>https://www.ragic.com/demo/sales/1?v=3&amp;api
</p>
```

Or for a single entry in your form, you would include an id (for example record id of 1) to specify an entry:

```
<p>https://www.ragic.com/demo/sales/1/1
</p>
```

Its corresponding HTTP API endpoint URL is simply:

```
<p>https://www.ragic.com/demo/sales/1/1?v=3&amp;api
</p>
```

## 1.6    API Limits

We don't have a limit on our API calls as long as it's under reasonable use. If you're not sure about your use, feel free to just drop us a message at [support@ragic.com](mailto:support@ragic.com) to check with us.

While there is no limit on API usage, there is a queue mechanism for each database account. The maximum size of the queue is 20, and it stops accepting API requests when it is full. Therefore, we encourage waiting for the response of one request before sending another. Note that this limit currently only applies to GET API requests.

This is only to ensure reasonable use of our system, so that there will not be disproportionately heavy use in smaller accounts.

A manual review process will be triggered, if the usage exceeds 5 requests per second, to determine whether throttling will be applied.

## 2.1.1    HTTP Basic authentication with Ragic API Key

You authenticate to the Ragic API by providing one of your API keys in the request. Your API keys carry many privileges, so be sure to keep them secret!

Because when your code accesses Ragic via an API key, it will basically log in as the user of the API key and execute read write as this user. **We highly recommend creating a separate user for API key access.** This way the API access will not be mixed with a organizational user, which will make the system audit trail much clearer and debugging of your API program much easier.

Authentication to the API occurs via [HTTP Basic Auth](http://en.wikipedia.org/wiki/Basic_access_authentication). Provide your API key as the basic auth username. You do not need to provide a password.

All API requests must be made over [HTTPS](http://en.wikipedia.org/wiki/HTTP_Secure). Calls made over plain HTTP will fail. **You must authenticate for all requests**.

```
<p>curl https://www.ragic.com/demo/sales/1\
</p><p>   --get -d api \
</p><p>   -H "Authorization:Basic YOUR_API_KEY_GOES_HERE"
</p>
```

![](Ragic%20API%20developer%20guide/file.3.jsp)

Note that the HTTP header name is Authorization, and the value is your API key preceded with "Basic ", Basic with a space at the end, and you may need to modify **www** to **na3**, **ap5**, or **eu2** in your URL based on your Ragic account URL.

You can generate your API key in [Personal Settings](https://www.ragic.com/intl/en/doc-user/20/personal-settings#4).

![](Ragic%20API%20developer%20guide/file.4.jsp)

Most HTTP clients (including web-browsers) present a dialog or prompt for you to provide a username and password (empty) for HTTP basic auth. Most clients will also allow you to provide credentials in the URL itself.

If for some reason that you are not able to send the API key as HTTP header or basic authorization, you can send the API key as a parameter with the name **APIKey**. You will need to add this parameter for every single request you send.

```
<p>curl https://www.ragic.com/demo/sales/1\
</p><p>   --get -d api \
</p><p>   -d "APIKey=YOUR_API_KEY_GOES_HERE"
</p>
```

![](Ragic%20API%20developer%20guide/file.5.jsp)

## 2.1.2    Password authentication

Sometimes if your platform does not support HTTP Basic authentication, you can pass the user's e-mail and password as log in credentials to authenticate your program.

**If you registered through Sign in with Google, make sure to register a Ragic password before proceeding with this tutorial.**

**Please only use this when you can not authenticate with HTTP Basic Authentication.**

You send a request for a session id with a valid e-mail and password. You can issue a HTTP request using the -d argument containing the id and password. The -c parameter will store sessionId in the cookie jar file specified:

```
<p>curl --get -d "u=jeff@ragic.com" \
</p><p> --data-urlencode "p=123456" \
</p><p> -d "login_type=sessionId" \
</p><p>-d api \
</p><p> -c cookie.txt \
</p><p> -k \
</p><p> https://www.ragic.com/AUTH
</p>
```

![](Ragic%20API%20developer%20guide/file.6.jsp)

If authentication failed, server will return **\-1**. If authenticated, you will receive a session id in the response like this:

```
<p>2z5u940y2tnkes4zm49t2d4
</p>
```

**Note that this authentication method is session based, and session is server dependent. You may need to modify the url based on the location of the account you wish to access. For example, https://ap8.ragic.com/AUTH for accounts that reside on server https://ap8.ragic.com.**

**To use the returned sessionId in future requests to remain authenticated, please include the sessionId in url parameter as sid= For example, https://www.ragic.com/demo/sales/1?sid=**

**The use of Ragic API will be covered in later chapters, just remember to include your session ID in this manner to remain authenticated.**

If you would like to retrieve detailed info on the log in user, you can also provide an additional **json=1** parameter so that Ragic will return a json object containing the details of the user.

```
<p>curl --get -d "u=jeff@ragic.com" \
</p><p> --data-urlencode "p=123456" \
</p><p> -d "login_type=sessionId" \
</p><p> -d "json=1" \
</p><p>-d api \
</p><p> -c cookie.txt \
</p><p> -k \
</p><p> https://www.ragic.com/AUTH
</p>
```

![](Ragic%20API%20developer%20guide/file.7.jsp)

The returned format will look something like this:

```
<p>{
</p><p>"sid":"8xkz874fdftl116vkd3wgjq0t",
</p><p>"email":"jeff@ragic.com",
</p><p>"accounts":
</p><p>  {
</p><p>    "account":"demo",
</p><p>    "ragicId":25,
</p><p>    "external":false,
</p><p>    "groups":["EVERYONE","SYSADMIN"]
</p><p>  }
</p><p>}
</p>
```

## 2.2.1    Returned data JSON format

For example, this is the API response for **https://www.ragic.com/demo/sales/1?v=3&api**. Remember to specify your cookie jar file in the call like "-b cookie.txt"

```
<p>{
</p><p>"1":{
</p><p>  "_ragicId": 1,
</p><p>  "_star": false,
</p><p>  "Account Name": "Alphabet Inc.",
</p><p>  "_index_title_": "Alphabet Inc.",
</p><p>  "Short Name": "",
</p><p>  "Account ID": "C-00002",
</p><p>  "EIN / VAT Number": "",
</p><p>...
</p><p>...
</p><p>},
</p><p>"0":{
</p><p>  "_ragicId": 0,
</p><p>  "_star": false,
</p><p>  "Account Name": "Ragic, Inc.",
</p><p>  "_index_title_": "Ragic, Inc.",
</p><p>  "Short Name": "Ragic",
</p><p>  "Account ID": "C-00001",
</p><p>  "EIN / VAT Number": "",
</p><p>...
</p><p>...
</p>
```

The listing is defaulted to 1000 entries. You can add paging parameter **limit** according to the following sections to get more entries.

The data in the subtables are displayed as the sample below. The attribute name is **\_subtable\_** followed by the subtable ID to identify which subtable it is. The content of the subtable is presented the same way as the fields in a form.

```
<p>"_subtable_2000154": {
</p><p>    "0": {
</p><p>      "_ragicId": 0,
</p><p>      "Contact Name": "Jeff Kuo",
</p><p>      "Title": "Technical Manager",
</p><p>      "Phone": "886-668-037",
</p><p>    },
</p><p>    "1": {
</p><p>      "Contact Name": "Amy Tsai",
</p><p>      "Title": "Marketing",
</p><p>      "Phone": "",
</p><p>    },
</p>
```

If your application does not need the data in the subtables, you can add the parameter **subtables=0** to turn off the fetching of subtable data.

The comments of the entries is retrieved as subtables also. It will be retrieved in a subtable with the field id of 61, but you will need to add parameter **comment=true** when retrieving data for comment data to be returned.

## 2.2.2    Filter Conditions

Very often your database contains a large amount of entries, so it's better to apply filters when you retrieve data. Ragic API filters are in a special format

You can use the parameter "where" to add a filter condition to a query as below:

```
<p>curl --get -d "where=2000123,eq,Alphabet Inc." \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>-d api \
</p><p>https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.8.jsp)

The parameter is a "," comma delimited format, with at least 3 arguments.

-   1\. Field id of the field that you would like to filter.
-   2\. Operand in the form of integer to specify your filter operation. The list of operands are listed below.
-   3\. The value that you would like to filter the field with. Remember if your value might include a "," comma character, please URL encode it or just use a %2C instead to avoid collision.

You can supply a query with multiple filter conditions as below:

```
<p>curl --get -d "where=2000123,eq,Alphabet Inc." \
</p><p>-d "where=2000127,eq,Jeff Kuo" \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>-d api \
</p><p>https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.9.jsp)

Here's the list of operands that you can use:

| Operand Name | Operand Value |
| --- | --- |
| Equals | eq |
| Regular Expression | regex |
| Greater or equals | gte |
| Less or equals | lte |
| Greater | gt |
| Less | lt |
| Contains | like |
| Equals a node id | eqeq |

Please note that:

1.When you filter by date or date time, they will need to be in the following format: **yyyy/MM/dd** or **yyyy/MM/dd HH:mm:ss**.

2.You don't need to fill the third argument if you want to filter empty values, for example, **"where=2000127,eq," \\**.

3\. **OR** filtering for the SAME field can be achieved by supplying multiple where queries. For example, to retrieve records where field id 1000001 is either Ratshotel **OR** Claflin, **"where=1000001,eq,Ratshotel&where=1000001,eq,Claflin"** .

There are some system fields that has special field ids that you can use in your query. Common system fields listed below:

| System field name | Field id |
| --- | --- |
| Create Date | 105 |
| Entry Manager | 106 |
| Create User | 108 |
| Last Update Date | 109 |
| Notify User | 110 |
| If Locked | 111 |
| If Starred | 112 |

You can also use a **full text search** as a query filter. Just provide your query term in the parameter **fts** and the matched result will be returned.

```
<p>curl --get -d "fts=Alphabet" \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>-d api \
</p><p>https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.10.jsp)

You can also apply [Shared View](https://www.ragic.com/intl/en/doc/57/Saving-frequent-searches-as-views). Just set the id as below.

```
<p>curl --get -d "filterId=YOUR_SHARED_VIEW_ID" \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>-d api \
</p><p>https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.11.jsp)

You could get the id by clicking the Shared View URL.

![](Ragic%20API%20developer%20guide/file.12.jsp)

## 2.2.3    Limiting Entry Number / Paging

Very often you do not want to fetch all entries with one request, you can use the **limit** and **offset** parameters to specify how many entries that you would like to retrieve, and how many entries that you would like to skip at the start, so clients can implement pages for viewing entries.

The usage of limit and offset parameters is similiar to SQL limit parameter. Offset is how many entries that you would like to skip in the beginning, and limit is how many entries that you would like to be returned.

**Returned data is defaulted to 1000 entries**, you will need to provide limit parameters if you want your response to have more than 1000 entries.

The format is as follows:

```
<p>limit=<limit>&amp;offset=<offset>
</offset></limit></p>
```

The **offset** parameter is the number of entries that should be skipped before returning, used for going through pages that has been retrieved. The **limit** parameter is the max number of records that should be returned per call.

For example the below call will skip the first **5** entries, and return 6 ~ 13, a total of **8** entries.

```
<p>curl --get -d "limit=8" \
</p><p>-d "offset=5" \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>-d api \
</p><p>https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.13.jsp)

## 2.2.4    Sorting and ordering

Without any ordering, the data is by default ordered by creation date and time, from oldest to latest. If you would like to have the latest result first, you can specify **reverse=true** like this:

```
<p>curl --get -d "reverse=true" \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>-d api \
</p><p>https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.14.jsp)

You can also specify how the entries are ordered by adding the **order** parameter. It's also kind of similar to the ORDER BY clause in SQL, its value is a comma separated string with two arguments. The first one is the field id of the domain that you would like the entries to be sorted according to, and the second one is the order: either ASC for ascending order, or DESC for descending order.

```
<p>curl --get -d "order=800236,DESC" \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>-d api \
</p><p>https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.15.jsp)

## 2.2.5    Callback Function (JSONP)

Adding a callback function parameter will enclose the returned JSON object as a argument in a call to the callback function specified by you. This is especially useful when you're accessing our API using Javascript, you may need this to do cross domain ajax calls.

For example, adding callback=testFunc as below:

```
<p>curl --get -d "callback=testFunc" \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>-d api \
</p><p>https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.16.jsp)

Will enclose the returned JSON object in a function call so you can process the returned data in your callback function:

```
<p>testFunc({"17": {
</p><p>"Account Name": "Dunder Mifflin",
</p><p>"Account Owner": "Jim Halpert",
</p><p>"Phone": "1-267-922-5599",
</p><p> ...
</p><p> ...
</p><p>});
</p>
```

## 2.2.6    Field Naming

The JSON data format uses the field name string as the attribute name, you can use a "naming" parameter to specify that you would like to use field id as the attribute name to identify a field.

Possible values include: EID (Field Id), FNAME (Field Name). This type of attribute will only need to be set once, and the configuration will be applied to all subsequent requests in this session.

```
<p>curl --get -d "naming=EID" \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>-d api \
</p><p>https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.17.jsp)

## 2.2.7    Other GET parameters

There are several other useful parameters that you can apply to your HTTP GET request to change the content of the response:

| Parameter Name | Description |
| --- | --- |
| subtables | Specifying **subtables=0** tells Ragic API to not include subtable information in the response. |
| listing | Specifying **listing=true** tells Ragic API to only include fields in the Listing Page. |
| reverse | Specifying **reverse=true** tells Ragic API to reverse the default ordering of the listing page response. |
| info | Adding the **info=true** parameter will add "Create Date", "Create User" information to the response |
| conversation | Adding the **conversation=true** parameter will add the email conversation information related to this record to the response |
| approval | Adding the **approval=true** parameter will add the approval information related to this record to the response |
| comment | Adding the **comment=true** parameter will add the comment thread related to this record to the response |
| bbcode | Adding the **bbcode=true** parameter will retrieve the raw BBCode value saved to the field instead of being translated to HTML |
| history | 
Adding the **history=true** parameter will add the edit history related to this record to the response.

Edit Histories are in the form of JSON, which contains information about the time, type, sheet, user, detail.

```
<p>{
</p><p>  "Time": "" // utc timestamp,
</p><p>  "Type": "" // type of edit,
</p><p>  "Sheet": "" // sheet the record belongs to,
</p><p>  "User": "" // the user who made the edit,
</p><p>  "Detail": "" // edit details
</p><p>}
</p>
```

 |
| ignoreMask | When **ignoreMask=true** is given, the field value of "Masked text" will be unmasked if you are in the viewable groups. Click [here](https://www.ragic.com/intl/en/doc/17/Field-Types#24) for more information about "Masked text" field. |
| ignoreFixedFilter | When **ignoreFixedFilter=true** is given, the fixed filter on this sheet will be ignored. But note that this will only work when the API call API key user has the SYSAdmin privilege. |

## 2.2.8    Retrieving uploaded files, images and e-mail attachments

On the JSON returned by your HTTP API call, you will see something like this for file upload field or image upload fields:

```
<p>"1000537": "Ni92W2luv@My_Picture.jpg",
</p>
```

You will be able to download the file using a separate call like this (assuming your account name is "demo" and API call url being https://www.ragic.com/demo/sales/1?v=3&api) :

```
<p>https://www.ragic.com/sims/file.jsp?a=demo&amp;f=Ni92W2luv@My_Picture.jpg
</p>
```

The format is:

```
<p>https://www.ragic.com/sims/file.jsp?a=&lt; account name &gt;&amp;f=&lt; file name &gt;
</p>
```

Remember to encode your file name when you send it as an URL. Your actual file name will start after the @ character, this is to avoid file name collision.

## 2.2.9    Retrieving HTML, PDF, Excel, Mail Merge and Custom Print Report of a record

For example, if you have an URL to the record as follows

```
<p>https://www.ragic.com/demo/sales/1/41
</p>
```

### 1\. Printer Friendly

You can retrieve an **HTML printer friendly** version like this, by adding a **.xhtml** to the end:

```
<p>https://www.ragic.com/demo/sales/1/41.xhtml
</p>
```

### 2\. PDF or Excel Version

You can add a **.pdf** for a **PDF** version, **.xlsx** for an **Excel** version:

```
<p>https://www.ragic.com/demo/sales/1/41.pdf
</p><p>https://www.ragic.com/demo/sales/1/41.xlsx
</p>
```

### 3\. Mail Merge

You can retrieve a **[Mail Merge](https://www.ragic.com/intl/en/doc/37/mail-merge)** of a record, by adding **.custom?** and a Mail Merge **CID** to specify which template (which is 1 for the mail merge template used in this case):

```
<p>https://www.ragic.com/demo/sales/1/41.custom?cid=1
</p>
```

To obtain the cid of a mail merge, you can manually download a mail merge from the Ragic user interface first, and pay attention to the cid parameter in the download url.

```
<p>https://www.ragic.com/demo/sales/1/41.custom?rn=41&amp;<b>cid=1</b>
</p>
```

### 4\. Custom Print Report

You can retrieve a **[Custom Print Report](https://www.ragic.com/intl/en/doc/148/custom-print-report)** of a record by appending **.carbone?** to the URL, along with the following parameters (joined with **&**).

```
<p>https://www.ragic.com/demo/sales/1/41.carbone?fileFormat=pdf&amp;customPrintTemplateId=1&amp;fileNameRefDomainId=1001000
</p>
```

(1) **File format** syntax: fileFormat="file format" (e.g., pdf, png, docx).

Example: **fileFormat=pdf**

(2) **Custom template ID** syntax: customPrintTemplateId="Template ID".

Example: **customPrintTemplateId=1**

To obtain the "Template ID", first manually download the Custom Print Report and find the Template ID parameter in the download URL:

```
<p>https://www.ragic.com/demo/sales/1/41.carbone?fileFormat=pdf&amp;fileNameRefDomainId=-1&amp;<b>customPrintTemplateId=2</b>
</p>
```

(3) **File name referenced field** syntax (Optional)：fileNameRefDomainId="[Field ID](https://www.ragic.com/intl/en/doc-kb/299/What-is-Field-ID%253F)".

Example: **fileNameRefDomainId=1001000**

## 2.3.1    Creating entries with custom HTML forms

If you just need to create your own HTML form to save data to Ragic, you don't need to write any API code to do that. Ragic has a very simple way for your form data to be posted to Ragic.

Suppose you want to create a form that saves entry in this [sample pet store merchandise form](https://www.ragic.com/start/petstore/1).

1\. Find the **Field Id** for each field that you would like to save in the HTML form. You can find them in design mode when you focus on a field.

2\. Create a form like this HTML sample, Your HTML form saves data to the same form URL, with the query string parameter "api". You put down the field id as the parameter name for each field to map them to a field on the Ragic form:

```
<form action="https://www.ragic.com/start/petstore/1?html&amp;api" method="POST">
<p>Item Id: </p>
<p>Item Name: </p>
<p>Item Price: </p>
<p>Item Category:
</p><p>       Dog
</p><p>       Cat
</p>
</form>
```

3\. Make sure that the user have access right to enter data on this form.

**Note that to upload files through our API, JSON data format is not supported, and form data is the only option.**

## 2.3.2    Creating a New Entry

If you are not using simple HTML forms to create entries on Ragic, you will need to create API requests to create the entries. The endpoints for writing to a form is the same as reading them, **but you issue a POST request instead of a GET request.**

**The API now supports JSON data, and it is the recommended way to make HTTP requests.**

To POST JSON data, you need to change Body settings to raw JSON, as the image below.

![](Ragic%20API%20developer%20guide/file.18.jsp)

What you need to do is use the field ids of the fields as name, and the values that you want to insert as parameter values.

Please note that your user will need write access to the form for this to work.

```
<p>curl -F "2000123=Dunder Mifflin" \
</p><p> -F "2000125=1-267-922-5599" \
</p><p> -F "2000127=Jeff Kuo" \
</p><p>-F "api=" \
</p><p> -H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p> https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.19.jsp)

The equivalent JSON format is as below,

```
<p>{
</p><p>    "2000123": "Dunder Mifflin",
</p><p>    "2000125": "1-267-922-5599",
</p><p>    "2000127": "Jeff Kuo",
</p><p>}
</p>
```

If the field is a **multiple selection** that can contain multiple values, you can have multiple parameters with the same field id as names. If the field is a **date field** the value will need to be in the format of **yyyy/MM/dd** or **yyyy/MM/dd HH:mm:ss** if there's a time part. So a request would look like this:

```
<p>curl -F "2000123=Dunder Mifflin" \
</p><p> -F "2000125=1-267-922-5599" \
</p><p> -F "2000127=Jeff Kuo" \
</p><p> -F "1000001=Customer" \
</p><p> -F "1000001=Reseller" \
</p><p> -F "2000133=2018/12/25 23:30:00" \
</p><p> -F "api=" \
</p><p> -H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p> https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.20.jsp)

The equivalent JSON format is as below,

the squre brackets for the field ID 1000001 allows for multiple values in one statement.

```
<p>{
</p><p>    "2000123": "Dunder Mifflin", 
</p><p>    "2000125": "1-267-922-5599", 
</p><p>    "2000127": "Jim Halpert",
</p><p>    "1000001": ["Customer", "Reseller"],
</p><p>    "2000133": "2018/12/25 23:30:00" 
</p><p>}
</p>
```

If you would like to insert data into the **subtables** at the same time, you will need a slightly different format for the fields in the subtables because Ragic needs a way to determine if the field values belong to the same entry in a subtable.

If the field values are in the same subtable row, assign them with the same negative row id with each other. It can be any negative integer. It's only a way to determine that they are in the same row.

```
<p>2000147_-1=Bill
</p><p>2000148_-1=Manager
</p><p>2000149_-1=billg@microsoft.com
</p><p>2000147_-2=Satya
</p><p>2000148_-2=VP
</p><p>2000149_-2=satyan@microsoft.com
</p>
```

The whole request would look like this:

```
<p>curl -F "2000123=Dunder Mifflin" \
</p><p> -F "2000125=1-267-922-5599" \
</p><p> -F "2000127=Jeff Kuo" \
</p><p> -F "1000001=Customer" \
</p><p> -F "1000001=Reseller" \
</p><p> -F "2000133=2018/12/25 23:30:00" \
</p><p> -F "2000147_-1=Bill" \
</p><p> -F "2000148_-1=Manager" \
</p><p> -F "2000149_-1=billg@microsoft.com" \
</p><p> -F "2000147_-2=Satya" \
</p><p> -F "2000148_-2=VP" \
</p><p> -F "2000149_-2=satyan@microsoft.com" \
</p><p>-F 'api=' \
</p><p> -H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p> https://www.ragic.com/demo/sales/1
</p>
```

![](Ragic%20API%20developer%20guide/file.21.jsp)

The equivalent JSON format is as below,

```
<p>{
</p><p>    "2000123": "Dunder Mifflin", 
</p><p>    "2000125": "1-267-922-5599", 
</p><p>    "2000127": "Jim Halpert",
</p><p>    "1000001": ["Customer", "Reseller"],
</p><p>    "2000133": "2018/12/25 23:30:00" 
</p><p>    "_subtable_2000154": {
</p><p>        "-1": {
</p><p>            "2000147": "Bill",
</p><p>            "2000148": "Manager",
</p><p>            "2000149": "billg@microsoft.com"
</p><p>        },
</p><p>       "-2": {
</p><p>            "2000147": "Satya",
</p><p>            "2000148": "VP",
</p><p>            "2000149": "satyan@microsoft.com"
</p><p>        }
</p><p>    }
</p><p>}
</p>
```

If you would like to populate a **file upload field**, just make sure that the request encoding type is a **multipart/form-data**. The HTML equivalent would be setting enctype='multipart/form-data'

With a multipart request, you can put the file in your request, and just put the file name as the field value.

```
<p>1000088=test.jpg
</p>
```

## 2.3.3    Modifying an Entry

Ragic supports using POST, PUT, and PATCH method to modify an entry.

The endpoint for modifying an entry is the same as reading an existing entry. Notice that when you create an entry, the endpoint points to a Ragic sheet, but **when you edit an entry, your endpoint will need an extra record id to point to the exact record**.

```
<p>https://www.ragic.com/<account>/<tab folder="">/<sheet index=""><b>/<record id=""></record></b>?api
</sheet></tab></account></p>
```

All you need to provide is the field ids of the fields that you would like to modify to. If the field is a **date field** the value will need to be in the format of **yyyy/MM/dd** or **yyyy/MM/dd HH:mm:ss** if there's a time part.

```
<p>curl -F "2000123=Dunder Mifflin" \
</p><p> -F "2000127=Jim Halpert" \
</p><p> -F "api=" \
</p><p> -H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p> https://www.ragic.com/demo/sales/1<b>/3</b>
</p>
```

![](Ragic%20API%20developer%20guide/file.22.jsp)

The equivalent JSON format is as below,

```
<p>{
</p><p>    "2000123": "Dunder Mifflin", 
</p><p>    "2000127": "Jim Halpert" 
</p><p>}
</p>
```

For subtables, it's a bit more tricky. Because Ragic will need to know which row that you're editing. So you will need to find the row id of the row that you're editing. This information can be found from an API call.

As we mentioned in earlier chapter, the returned format for an entry with subtables look like this:

```
<p>"_subtable_2000154": {
</p><p>    "0": {
</p><p>      "Contact Name": "Jeff Kuo",
</p><p>      "Title": "Technical Manager",
</p><p>      "Phone": "886-668-037",
</p><p>      "E-mail": "jeff@ragic.com",
</p><p>...
</p><p>...
</p><p>    },
</p><p>    "1": {
</p><p>      "Contact Name": "Amy Tsai",
</p><p>      "Title": "Marketing",
</p><p>      "Phone": "",
</p><p>...
</p><p>...
</p><p>    },
</p><p>    "2": {
</p><p>      "Contact Name": "Allie Lin",
</p><p>      "Title": "Purchasing",
</p><p>...
</p><p>...
</p>
```

In the subtable, **1** is the row id for the contact Amy Tsai, and **2** is the row id for the contact Allie Lin. With this row id, you can modify data in the subtable pretty much like how you create an entry with subtable data.

You use the row id as the identifier following the field id. You only need to put in the fields that you want to modify:

```
<p>2000147_1=Ms. Amy Tsai
</p><p>2000148_1=Senior Specialist
</p><p>2000148_2=Senior Manager
</p>
```

The whole request would be like this:

```
<p>curl -F "2000123=Dunder Mifflin" \
</p><p> -F "2000127=Jim Halpert" \
</p><p> -F "2000147_1=Ms. Amy Tsai" \
</p><p> -F "2000148_1=Senior Specialist" \
</p><p> -F "2000148_2=Senior Manager" \
</p><p> -F "api=" \
</p><p> -H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p> https://www.ragic.com/demo/sales/1<b>/3</b>
</p>
```

![](Ragic%20API%20developer%20guide/file.23.jsp)

The equivalent JSON format is as below,

```
<p>{
</p><p>    "2000123": "Dwight Schrute", 
</p><p>    "2000127": "Jim Halpert" ,
</p><p>    "_subtable_2000154": {
</p><p>        "29" :{
</p><p>            "2000147": "Ms. Amy Tsai",
</p><p>            "2000148": "Senior Specialist"
</p><p>        },
</p><p>        "30" :{
</p><p>            "2000148": "Senior Manager"
</p><p>        } 
</p><p>    }
</p><p>}
</p>
```

If you want to delete a subtable row, you can create a request like:

```
<p>DELSUB_<subtable key="">=<subtable row="" id="">
</subtable></subtable></p>
```

The equivalent JSON format is as below,

```
<p>_DELSUB_<subtable key="">=[<subtable row="" id="">,<subtable row="" id="">,...,<subtable row="" id="">];
</subtable></subtable></subtable></subtable></p>
```

For example, if you want to delete the contact Arden Jacobs, the whole request would be like this:

```
<p>curl -F "DELSUB_2000154=3" \
</p><p> -F "api=" \
</p><p> -H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p> https://www.ragic.com/demo/sales/1<b>/3</b>
</p>
```

![](Ragic%20API%20developer%20guide/file.24.jsp)

The equivalent JSON format is as below,

```
<p>{
</p><p>    "_DELSUB_2000154": [3]
</p><p>}
</p>
```

Using the JSON format for subtable row deletion allows you to specify multiple rows in a simple manner,

```
<p>{
</p><p>    "_DELSUB_subtable key": [<subtable row="" id="">,..., <subtable row="" id="">]
</subtable></subtable></p><p>}
</p>
```

## 2.3.4    Create / Update Parameters

There are many useful parameters that you can use when creating or updating records on Ragic to save your time writing duplicate code for what can be done on Ragic.

| Parameter Name | Description |
| --- | --- |
| doFormula | Specifying doFormula=true tells Ragic API to recalculate all formulas first when a record is created or updated. Do note that if this is set to true, the workflow scripts that you configured on the sheet will not run to avoid infinite loops. |
| doDefaultValue | Specifying doDefaultValue=true tells Ragic API to load all default values when a record is created or updated. |
| doLinkLoad | Specifying doLinkLoad=true tells Ragic API to recalculate all formulas first and then load all link and load loaded values when a record is created or updated.
Specifying doLinkLoad=first tells Ragic API to load all link and load loaded values first when a record is created or updated and then recalculate all formulas.

 |
| doWorkflow | Specifying doWorkflow=true tells Ragic API to execute the workflow script associated with this API call. |
| notification | Specifying notification=true tells Ragic API to send out notifications to relevant users, or false, otherwise. The default is true when not specified. |
| checkLock | Specifying checkLock=true tells Ragic API to check if the record is locked before an update, and not edit the record if the record is locked. |

## 2.3.5    Deleting an entry

Deleting an entry is very similar to reading an entry. All you need to do is change the request method from GET to **DELETE**. When it's a DELETE request, the entry at the API endpoint will be deleted.

```
<p>curl -X DELETE \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>-d "api"  \
</p><p>https://www.ragic.com/demo/sales/1/3
</p>
```

![](Ragic%20API%20developer%20guide/file.25.jsp)

## 2.3.6    Upload files and images

Make sure your content-type is **multipart/form-data**, and you will be able to upload files.

```
<p>curl -F "1000088=@/your/file/path" \
</p><p>-F "api=" \
</p><p>-F "v=3" \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>https://www.ragic.com/demo/sales/1
</p>
```

After a success uploading, you will receive your file in this format.

```
<p>"1000088": "Ni92W2luv@test.jpg",
</p>
```

![](Ragic%20API%20developer%20guide/file.26.jsp)

Hover your mouse over the boxed area to expose the drop-down list, and select the "File" option to allow file uploads.

![](Ragic%20API%20developer%20guide/file.27.jsp)

If you want to get files just uploaded, follow instruction [here](https://www.ragic.com/intl/en/doc-api/28).

If you want to upload files via link, download it first and then follow instruction above.

```
<p>curl -o __TEMP_FILE__ YOUR_IMAGE_LINK
</p><p>curl -F "1000002=@__TEMP_FILE__" -F "api=" -F "v=3" -H "Authorization:Basic YOUR_API_KEY" https://www.ragic.com/demo/sales/1
</p>
```

Remember to replace **YOUR\_IMAGE\_LINK** and **YOUR\_API\_KEY**

Your API key could obtain [here](https://www.ragic.com/sims/reg/getAPIKey.jsp).

## 2.3.7    Comment

Make sure your content-type is **multipart/form-data**.

```
<p>curl -F "at=@/your/file/path" \
</p><p>-F "c=yourComment"
</p><p>-F "api=" \
</p><p>-H "Authorization:Basic YOUR_API_KEY_GOES_HERE" \
</p><p>https://www.ragic.com/demo/ragicsales-order-management/10001/2
</p>
```

![](Ragic%20API%20developer%20guide/file.28.jsp)

In the request body, the value of the parameter **c** is the comment(required), and the value of the parameter **at** is the attachment(optional).

Your API key could obtain [here](https://www.ragic.com/sims/reg/getAPIKey.jsp).

## 2.4    API Request Error Response

If the request returned in error, it will look something like this:

```
<p>{
</p><p>"status": "ERROR",
</p><p>"code": 303,
</p><p>"msg": "Your account has expired."
</p><p>}
</p>
```

Here is a list of the HTTP status code that might be returned on a Ragic request:

| HTTP Status | Description |
| --- | --- |
| 200 | OK - Everything worked as expected. |
| 400 | Bad Request - Often missing a required parameter. |
| 401 | Unauthorized - No valid API key provided. |
| 402 | Request Failed - Parameters were valid but request failed. |
| 404 | Not Found - The requested item doesn't exist. |
| 500, 502, 503, 504 | Server errors - something went wrong on Ragic's end. |

If there was an error processing your request, it will generally contain an error code and a description. Here is the list of error codes that you might receive as a response to a Ragic request:

| Error Code Id | Error Description |
| --- | --- |
| 101 | Invalid Account Name: {Account\_Name} |
| 102 | Invalid Path: {Path} |
| 103 | Invalid Form Index: {Form\_Index} |
| 104 | Cannot POST Data To A Custom Form |
| 105 | Authentication Required Before Using API |
| 106 | No Access Right |
| 107 | Resource Bundle Not Found |
| 108 | Error Loading Requested Form |
| 109 | Cannot Create More Records |
| 201 | Error Processing Request Parameters |
| 202 | Error Executing Request |
| 203 | POST Request Did Not Finish |
| 204 | Request Frequency Too High |
| 301 | Sid Parameter / Session Has Timed Out |
| 303 | Account Expired |
| 304 | Secret Key Is Invalid |
| 402 | Record Locked |
| 404 | Record Not Found |

## 2.5    Sample code

You can find sample codes in different programming languages on [Ragic's GitHub page](https://github.com/ragic/public/tree/master/HTTP%20API%20Sample)

If you're having trouble understanding how to work with RESTful APIs, you can also check out some [REST samples](http://rest.elkstein.org/2008/02/rest-examples-in-different-languages.html) here.

For JSON parsing, you can find tons of info on finding tools to parse JSON on [json.org](http://www.json.org/)

## 2.6.1    Common Q&A

## Does Ragic API support fetching metadata of a sheet?

Unfortunately, this feature is currently not supported. However, we do have plans supporting this feature. We also welcome suggestions and ideas, please contact us at support@ragic.com.

## Can Ragic API return all data entries in a sheet at once?

Currently, the only way to achieve this is by setting the correct limit parameter.

For more details, please refer to this [documentation](https://www.ragic.com/intl/en/doc-api/10/Limiting-Entry-Number-%2F-Paging).

Note that reading a large number of entries in a single request can significantly slow down the response rate of Ragic API.

## Why does Ragic only return 1000 entries?

GET API by default only returns 1000 entries. If you wish to modify this behavior, please use the limit and offset parameters.

For more details, please refer to this [documentation](https://www.ragic.com/intl/en/doc-api/10/Limiting-Entry-Number-%2F-Paging).

Note that reading a large number of entries in a single request can significantly slow down the response rate of Ragic API.

## Does Ragic API support creating/updating multiple data entries in one request?

This feature is currently not supported. However, we do have plans supporting bulk APIs. We also welcome bulk API suggestions and ideas, please contact us at support@ragic.com.

## Does Ragic API impose a usage limit?

We do not impose any limit on reasonable use. We will only send notification and apply throttling if we detect unfair usage from a specific IP or user.

For more details, please refer to this [documentation](https://www.ragic.com/intl/en/doc-api/22/API-Limits).

## Why does it sometimes take a while to receive a HTTP response?

There is a queue mechanism implemented for Ragic API, and each database account has an independent queue. All requests will be delayed until the ones in front have completed. In the event of serious delays, the queued requests will eventually timeout. It is recommended to send one request at a time, and to only send a new request, after the current request has received a response.

## 2.6.2    Common Questions For API Parameter

## READ

| Parameter | Common Question |
| --- | --- |
| where=,, | 
Some fields can contain special symbols such as spaces and newlines, and these symbols should be taken into account when filtering.

It is recommended to first fetch the data with an API request to inspect the field values, before constructing your filter query parameters.

 |

## WRITING

| Parameter | Common Question |
| --- | --- |
| doLinkLoad=true | 
If the specified linked field value does not exist in the source sheet, the linked field and all relevant loaded fields values will be cleared.

Link and Load operation will overwrite the specified field value within the HTTP request.

 |
| doFormula=true | 

If the formula referred fields do not contain sensible values (i.e. empty field values or non-numeric values for arithmetic formulae), the formula field may have its value cleared.

Formula calculation will overwrite the specified field value within the HTTP request.

 |
| doWorkflow=true | 

Workflow can contain complex operations, it is indeed difficult to debug. We suggest comparing the result with the web version.

 |
| doDefaultValue=true | 

There are two types of default values:

(1) “create xxx”, with $ as the first symbol

(2) “last modified xxx”, with # as the first symbol

$ starting default values are only triggered on creation of a data entry, while # starting default values are triggered every update.

For example, $DATETIME is only triggered on data entry creation, while #DATETIME is triggered every update.

 |

## 2.7    Mass Operation

Mass operation APIs are designed to perform the same set of operations for multiple records on a sheet in one single request.

There are two ways of specifying the records to be updated:

-   [where](https://www.ragic.com/intl/en/doc-api/9/Filter-Conditions) filters

```
<p>https://www.ragic.com/&lt; account &gt;/&lt; tab folder &gt;/&lt; sheet index &gt;/massOperation/&lt; Mass Operation Type &gt;?api&amp;where=&lt; Field ID &gt;,&lt; Filter Operand &gt;,&lt; Value &gt;
</p>
```

-   recordId in query string, recordId=< recordId >. e.g. recordId=1&recordId=2

```
<p>https://www.ragic.com/&lt; account &gt;/&lt; tab folder &gt;/&lt; sheet index &gt;/massOperation/&lt; Mass Operation Type &gt;?api&amp;recordId=&lt; recordId &gt;
</p>
```

## Request Format

-   Mass operation APIs are **aynschronous** operations.
-   It is required to modify www to na3, ap5, or eu2 in the API URL based on your Ragic database account URL.

```
<p>HTTP Method - POST
</p><p>URL - https://www.ragic.com/&lt; account &gt;/&lt; tab folder &gt;/&lt; sheet index &gt;/massOperation/&lt; Mass Operation Type &gt;?api
</p><p>Headers
</p><p>Authorization: Basic &lt; API Key &gt;
</p><p>Body
</p><p>{
</p><p>// JSON data that describes the operation to be performed
</p><p>}
</p><p>==========
</p><p>Response
</p><p>{
</p><p>"taskId": &lt; A UUID That Identifies The Task &gt;
</p><p>}
</p>
```

## Mass Lock

The mass lock API allows locking or unlocking multiple records at once.

[Mass Lock Document](https://www.ragic.com/intl/en/doc-user/64/batch-execute#2)

```
<p>HTTP Method - POST
</p><p>URL - https://www.ragic.com/&lt; account &gt;/&lt; tab folder &gt;/&lt; sheet index &gt;/massOperation/massLock?api
</p><p>{
</p><p>"action": &lt; lock or unlock &gt;
</p><p>}
</p><p>==========
</p><p>Response
</p><p>{
</p><p>"taskId": "6dbc796a-07d5-475b-b578-d254eb30f7d2"
</p><p>}
</p>
```

## Mass Approval

The mass approval API allows approval or rejection of multiple records at once.

[Mass Approval Document](https://www.ragic.com/intl/en/doc-user/64/batch-execute#3)

```
<p>HTTP Method - POST
</p><p>URL - https://www.ragic.com/&lt; account &gt;/&lt; tab folder &gt;/&lt; sheet index &gt;/massOperation/massApproval?api
</p><p>{
</p><p>"action": &lt; approve or reject &gt;,
</p><p>"comment": &lt; optional comment &gt; // optional
</p><p>}
</p><p>==========
</p><p>Response
</p><p>{
</p><p>"taskId": "6dbc796a-07d5-475b-b578-d254eb30f7d2"
</p><p>}
</p>
```

## Mass Action Button

The mass action button API allows the execution of an action button on multiple records at once.

[Mass Action Button Document](https://www.ragic.com/intl/en/doc-user/64/batch-execute#1)

```
<p>HTTP Method - POST
</p><p>URL - https://www.ragic.com/&lt; account &gt;/&lt; tab folder &gt;/&lt; sheet index &gt;/massOperation/massActionButton?api
</p><p>{
</p><p>"buttonId": &lt; button ID &gt;
</p><p>}
</p><p>==========
</p><p>Response
</p><p>{
</p><p>"taskId": "6dbc796a-07d5-475b-b578-d254eb30f7d2"
</p><p>}
</p>
```

To Fetch The List Of Available Action Buttons On A Sheet

```
<p>HTTP Method - GET
</p><p>URL - https://www.ragic.com/&lt; account &gt;/&lt; tab folder &gt;/&lt; sheet index &gt;/metadata/actionButton?api&amp;category=massOperation
</p><p>==========
</p><p>Response
</p><p>{
</p><p>"actionButtons": [
</p><p>{
</p><p>"id": &lt; button ID 1 &gt;,
</p><p>"name": &lt; button name 1 &gt;
</p><p>},
</p><p>.....
</p><p>,{
</p><p>"id": &lt; button ID 2 &gt;,
</p><p>"name": &lt; button name 2 &gt;
</p><p>}
</p><p>]
</p><p>}
</p>
```

## Mass Update

The mass update API allows updates of field values on multiple records at once.

[Mass Update Document](https://www.ragic.com/intl/en/doc-user/5/mass-update-records#1)

```
<p>HTTP Method - POST
</p><p>URL - https://www.ragic.com/&lt; account &gt;/&lt; tab folder &gt;/&lt; sheet index &gt;/massOperation/massUpdate?api
</p><p>{
</p><p>    "action": [
</p><p>        {
</p><p>            "field": &lt; Field ID &gt;,
</p><p>            "value": &lt; New Field Value &gt;
</p><p>        }
</p><p>    ]
</p><p>}
</p><p>==========
</p><p>Response
</p><p>{
</p><p>    "taskId": "6dbc796a-07d5-475b-b578-d254eb30f7d2"
</p><p>}
</p>
```

The mass update API also supports using on [Internal Users](https://www.ragic.com/intl/en/doc/42/internal-users) and [External Users](https://www.ragic.com/intl/en/doc/43/external-users), but there are some restrictions.

The following fields cannot be mass updated:

-   E-mail (domainId: 1)
-   Full Name (domainId: 4)
-   System Log (domainId: 10)
-   Status (domainId: 31)
-   Internal/External (domainId: 43)

Mass updating Ragic Groups (domainId: 3) has to follow the rules below:

-   The value of the key "value" has to be written in JSON Array, which has strings as its content.
-   All special characters has to be used with escape character(\\), especially " has to be used as \\".

```
<p>HTTP Method - POST
</p><p>URL - https://www.ragic.com/&lt; account &gt;/&lt; tab folder &gt;/&lt; sheet index &gt;/massOperation/massUpdate?api
</p><p>{
</p><p>    "action": [
</p><p>        {
</p><p>            "field": 3,
</p><p>            "value": "[\"SYSAdmin\"]"
</p><p>        }
</p><p>    ]
</p><p>}
</p><p>==========
</p><p>Response
</p><p>{
</p><p>    "taskId": "6dbc796a-07d5-475b-b578-d254eb30f7d2"
</p><p>}
</p>
```

-   In Internal Users, it is necessary to have at least one user who's Ragic Group is SYSAdmin.
-   In External Users, the name of users' Ragic Group has to start with "x-" or "X-".

## Mass Search And Replace

The mass search and replace API allows value replacement on multiple records at once.

[Mass Search And Replace Document](https://www.ragic.com/intl/en/doc-user/5/mass-update-records#2)

```
<p>HTTP Method - POST
</p><p>URL - https://www.ragic.com/&lt; account &gt;/&lt; tab folder &gt;/&lt; sheet index &gt;/massOperation/massSearchReplace?api
</p><p>{
</p><p>"action": [
</p><p>{
</p><p>"field": &lt; Field ID &gt;,
</p><p>"valueReplaced": &lt; Value To Be Replaced &gt;,
</p><p>"valueNew": &lt; Value To Replace With &gt;,
</p><p>}
</p><p>]
</p><p>}
</p><p>==========
</p><p>Response
</p><p>{
</p><p>"taskId": "6dbc796a-07d5-475b-b578-d254eb30f7d2"
</p><p>}
</p>
```

## Task Progress Tracking

Mass operations are asynchronous operations.

The task Id of the operation can be used to monitor its progress.

```
<p>HTTP Method - GET
</p><p>URL - https://www.ragic.com/&lt; account &gt;?api&amp;taskId=&lt; task ID &gt;
</p><p>==========
</p><p>Response
</p><p>{
</p><p>"id": &lt; task ID &gt;,
</p><p>"ap": &lt; account &gt;,
</p><p>"taskName": &lt; task name &gt;,
</p><p>"status": &lt; status &gt;
</p><p>}
</p>
```

## 3.1    What is a webhook

Webhook is a way for your external application to subscribe to changes to your Ragic application.

By subscribing to the changes, whenever a change is made on Ragic, Ragic will call the Webhook URL you provided to notify you of the change, including the ID of exact record that has been changed.

The biggest advantage of using a webhook API instead of a RESTful API, is that it is a lot more efficient in processing changes. You no longer have to poll the API every X hours to watch for latest changes.

## 3.2    Webhook on Ragic

You can find webhook for a sheet on Ragic by clicking on the Tools button at the top. You will find the webhook in the Sync section.

![](Ragic%20API%20developer%20guide/file.29.jsp)

Things to note on the webhook API:

1\. It will be triggered on create, update, and delete.

2\. It is not completely synchronized, there may be a slight delay when the loading is high.

3\. Changes will not include changes on related sheets like multiple versions.

![](Ragic%20API%20developer%20guide/file.30.jsp)

Click on the “webhook” feature and enter the URL that should receive this notification, and you’re done. You can also opt in to receive full information on the modified data by checking the checkbox "send full content on changed record". To cancel the webhook configuration, click on the x in the corner of the webhook configuration box.

The JSON format **without full content** would look like

```
<p>[1 ,2 ,4]
</p>
```

For this example, it means entries with node ID 1 & 2 & 4 were changed.

The JSON format **with full content** would look like

```
<p>{
</p><p>  "data": [
</p><p>          &lt; THE MODIFIED DATA ENTRY IN JSON &gt;
</p><p>   ]
</p><p>  "apname": "&lt; ACCOUNT NAME &gt;",
</p><p>  "path": "&lt; PATH NAME &gt;",
</p><p>  "sheetIndex": &lt; SHEET INDEX &gt;,
</p><p>  "eventType": "&lt; EVENT TYPE CREATE/UPDATE/DELETE &gt;"
</p><p>}
</p>
```