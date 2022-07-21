![EridanSharpTitleLogo](https://user-images.githubusercontent.com/57411317/177404573-ce2d4dcd-9623-4f12-af2b-40c21e1e3f36.png)

| | |
|-|-|
| Download | [![](https://img.shields.io/nuget/v/EridanSharp)](https://www.nuget.org/packages/EridanSharp) |

Firstly EridanSharp can help with authenticate in Google service.

# Documentation

Support formats: ```aac```, ```abw```, ```arc```, ```avi```, ```azw```, ```bin```, ```bmp```, ```bz```, ```bz2```, ```csh```, ```css```, ```csv```, ```doc```, ```docx```, ```eot```, ```epub```, ```gz```, ```gif```, ```html```, ```ico```, ```ics```, ```jar```, ```jpeg```, ```jpg```, ```js```, ```json```, ```jsonld```, ```mjs```, ```mp3```, ```mpeg```, ```mpkg```, ```odp```, ```ods```, ```odt```, ```oga```, ```ogv```, ```ogx```, ```opus```, ```otf```, ```png```, ```pdf```, ```php```, ```ppt```, ```pptx```, ```rar```, ```rtf```, ```sh```, ```svg```, ```swf```, ```tar```, ```tif```, ```tiff```, ```ts```, ```ttf```, ```txt```, ```vsd```, ```wav```, ```weba```, ```webm```, ```webp```, ```woff```, ```woff2```, ```xhtml```, ```xls```, ```xlsx```, ```xml```, ```xul```, ```zip```, ```3gp```, ```7z```.

If you did not find the required file format, then it will be interpreted as base binary data type.

## Examples
### Send email via gmail (OAuth2)
```csharp
static int Main(string[] args)
{
    const string clientId = "YOUR_CLIENT_ID";
    const string clientSecret = "YOUR_CLIENT_SECRET";
    const string pathSuccessPage = "success_page.html";
    const string pathUnsuccessPage = "unsuccess_page.html";
    string pathToken = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\data\\token.json";

    Gmail gmail = new Gmail(clientId, clientSecret, pathSuccessPage, pathUnsuccessPage, pathToken);

    bool existToken = gmail.CheckExistToken();

    if (!existToken)
    {
        gmail.Authentication();
    }

    Console.WriteLine("Email: " + gmail.GetProfile().EmailAddress);
    Console.WriteLine("Message total: " + gmail.GetProfile().MessagesTotal);
    Console.WriteLine("Threads total: " + gmail.GetProfile().ThreadsTotal);
    Console.WriteLine("History id: " + gmail.GetProfile().HistoryId);

    MimeMessage message = new MimeMessage();
    message.FromName = "FROM_NAME";
    message.FromEmail = "FROM_EMAIL";
    message.ToName = "TO_NAME";
    message.ToEmail = "TO_EMAIL";
    message.Subject = "SUBJECT";
    message.BodyText = "BODY_TEXT";
    message.AddAttachment("file.docx");

    gmail.Send(message);

    Console.ReadLine();
    return 0;
}
```

### Support
EridanSharp is an open-source project with a single mainteiner. If you want to solve any problem related to EridanSharp, you will have to do it yourself. Fork the repository and submit a pull request.

### License: MIT
