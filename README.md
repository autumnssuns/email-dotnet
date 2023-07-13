# .NET Console Application to send Emails using SMTP

## Installation

Run the following command in the root directory of the project to install the required packages.

```
dotnet add package Microsoft.Graph
```

Modify the `appsettings.json > "Smtp"` host and port to match the host, port and SSL of your SMTP server.

For example, for Gmail:
```
  "Smtp": {
    "host": "smtp.gmail.com",
    "port": 587,
    "ssl": true
  }
```

## Usage

1. In the `Targets.xlsx` workbook, add the email addresses of the recipients in the `Recipient Email` column, and the `Attachment Path` column if you want to attach a file to the email.
2. Edit the template email if required.
3. If more attributes are required, add them to the columns from F to I, for example, `Attribute`, and add `<Attribute>` to the template email where the attribute is required.
4. The `Subject` and `Body` columns will be computed based on the template email and the attributes.
5. To create draft emails, run the following command in the root directory of the project.

```
dotnet run <email> <password>
```

Where `<email>` is the email address of the sender, and `<password>` is the password of the sender.