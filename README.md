# Automation for Aspire Budget

Aspire Budget is a great Google Sheet available at https://aspirebudget.com/.

This collection of Google Apps Script functions automates budgeting:
- adding of transactions;
- categorization of transactions;
- adding second row for “↕️ Account Transfer” transaction with macros.

## Macros for adding second row for “↕️ Account Transfer” transaction

As Aspire Budget Getting Started Guide says:
> If you ever need to move funds between Accounts (to pay a Credit Card for example), you can select the ↕️ Account Transfer option in the Category dropdown. You'll need to create an Outflow Transaction to move money out of an account and an Inflow Transaction to move money into an account.

So, to automate this here is the script. It can be especially useful if you have imported lots of historical transactions and you need to create “↕️ Account Transfer” second rows in batch.

### Adding the macros to your Aspire Budget Google Sheet

1. Open [Code.gs](Code.gs) > Ctrl+F > "Macros"
2. Copy the functions of this section of [Code.gs](Code.gs)
3. Open your Aspire Budget Google Sheet > Tools menu > Script editor
4. Paste to the end of file
5. Save
6. In your Aspire Budget Google Sheet click Tools menu > Macros > Import
7. You’ll see AddSecondRowForAccountTransfer, click Add Function below it, and close this popup window
8. Tools menu > Macros > Manage macros > type in "1" in the text box after *Ctrl + Alt + Shift +* > Update

### Using the macros

1. Select a row or rows for which you need to create a second row(s)
2. Press Ctrl + Alt + Shift + 1
3. Do not select any other cells in the spreadsheet until all extra rows are created

## Automation of transaction adding and categorization

Almost all banks and payment services have a functionality to notify about transactions: via SMS, or e-mail, or push notifications, or webhook calls, or even with API.

So, you can use Google Apps Script to parse these messages and add all transactions to your Aspire Budget Google Sheet automatically.

Moreover, all of the new transactions are automatically categorized based on categories you have assigned to such transactions before.

### Important notice

Use of these scripts implies that you have basic knowledge of programming.

That’s because you are likely a customer of a bank or payment system which is not yet supported in [this Google Apps Script](Code.gs).

And because banks and payment systems change the format of their transaction messages time to time (in general, once in a few years), and when the format changes you would need to fix some parts of the code to adapt to the changes.

### Setting things up

1. Check what notification capabilities have your bank / payment system
2. Test them and choose what you’ll use
   - so, for example, [PayPal IPN](https://developer.paypal.com/docs/ipn/) is a great feature, but it notifies only about inflow transactions, though PayPal e-mails notify about every transaction;
   - many banks send full transaction info in e-mails, but only brief transaction info in SMS or push notifications;
   - some banks can send push notifications, but in the moments when you spend money but your device is offline they send SMS instead;
   - to the moment (2020-02-29) you can forward SMS and push notifications on Android devices, but not in iOS
3. If you choose to use e-mail notifications, set up their forwarding to your Gmail account, the same as you work with Aspire Budget Google Sheet in
4. If you choose to use SMS notifications, set up their forwarding to your Gmail account as well, with [SMS Backup+](https://play.google.com/store/apps/details?id=com.zegoggles.smssync), for example
4. **Note**: Google Apps Script has a [limit of running time](https://developers.google.com/apps-script/guides/services/quotas). That could become a problem when you have lots of messages, so an approach with starring processed messages in Gmail is implemented. So, the script processes every transaction message that is not starred, adds it to Aspire Budget Google Sheet, and makes the message starred afterwards. Therefore, if you star a transaction message in Gmail manually, it won’t be processed by the script. And vice versa, if you unstar a transaction message in Gmail manually, it will be processed one more time, and this transaction will appear in you Aspire Budget Google Sheet twice. So, a general advice is not to star or unstar transaction messages in Gmail manually to avoid that. As starred messages appear in Gmail in Primary tab, you may want to [filter](https://mail.google.com/mail/u/0/?#settings/filters) such messages: *Mark as read, Categorize as Updates*.
5. Open your Aspire Budget Google Sheet
6. On Transactions tab add one extra column and name it Comment (this one is quite useful as with automation Memo column is used for bank message, and Comment is to be used for your personal comments)
7. Then add one more extra column and name it Keywords, then you can hide it: right-click > Hide column (this column is used for keywords extraction from transaction messages to provide categorization functionality)
8. Add 2 extra tabs to your Aspire Budget Google Sheet from here: [Aspire Budget Automation Extra Tabs](https://docs.google.com/spreadsheets/d/1eoOGVff2VydL62S197weVrPZrSbDpNb0xqAvlVaM_WQ/view), you can keep both of them hidden (right click on the sheet tab > Hide sheet)
   - *Log* tab is needed to collect all of the incoming messages, so in case of getting some rare type of transaction message, or in case of message format change you would have full log of the messages saved
   - *Keywords & Categories* tab is needed for collecting your statistics of categories for every keyword
9. As soon as you have received a number of messages of different transactions (different types of inflow and outflow ones) check [Code.gs](Code.gs) and see what code to use as a basis for your bank / payment system messages parsing
10. Roll up your sleeves and adapt the code
11. In the end in Google Apps Script editor click Edit menu > Current project's triggers > Add trigger
    - Choose which function to run: process
    - Select event source: Time-driven
    - Select type of time based trigger: Minutes timer
    - Select minute interval: Every 10 minutes
    - Failure notification settings: Notify me daily

### Working with automated process

1. Open your Transactions tab time to time and check the transactions added automatically (automatically added transactions will have no icon in “Cleared” column)
2. Reassign categories if needed (e.g., you may change “Groceries” to “Gifts” if this time you have bought candies for a birthday)
3. In “Cleared” column put ✅ for these new transactions
4. Enjoy your time and clear mind you have freed from manual work, while still having full control of your budget.
