# Scripting MailMaven

This guide provides a basic introduction to MailMaven's scripting capabilities to people who are already familiar with AppleScript.
Specific vocabulary can found by examining MailMaven's script dictionary in any capable script editor.

Note that while this dictionary is compatible with using javascript to interact with Maven via Open Scriting Architecture, AppleScript is used for elaboration and examples.



## Basic Objects

If you are familiar with scripting Mail, MailMaven has the same structure for accounts, mailboxes and messages.

1. Application "MailMaven" contains accounts
2. Accounts contain mailboxes
3. Mailboxes contain messages

From the root Application level you can access accounts and from there, a specific mailbox and a mailbox's messages.

 
```applescript
tell application "MailMaven"
    tell (first account whose name is "Main")
        tell (first mailbox whose name is "inbox")
            subject of (every message whose read status is false)
        end tell
    end tell
end tell
```

<div style = "background:lightyellow; padding: 10px; margin:25px; border:1px solid palegoldenrod; border-radius:10px">
<strong>Note</strong>:  In most of sample scripts that follow, the end tell blocks are excluded for brevity.
</div>

Additionally accounts and mailboxes can be specified using their name:

```applescript
tell application "MailMaven"
    tell account "Main" 
        tell mailbox "Inbox"
```

The Local Account can be identified by the name "Maven Local"


Mailboxes also have their own heirarchy:   

```applescript
tell application "MailMaven"
    tell account "Main" 
         mailbox "foo" of mailbox "bar"     
         mailbox "parent/child" -- using the / as a path delimiter
```

As well, all mailboxes are accessible at the root level by prefixing the mailbox path with the account name

```applescript
tell application "MailMaven"
    messages of mailbox "SC/Inbox" -- SC is the account name (note that this reference is within the "Application context"
    tell account "Foo"
        messages of mailbox "SC/Parent/Child" of application "MailMaven"  -- specify the application inside other scopes.
```

<div style = "background:lightyellow; padding: 10px; margin:25px; border:1px solid palegoldenrod; border-radius:10px">
<strong>Tip</strong>:  Set a script property `property Maven : application "MailMaven"`at the top of your script to access top level objects easily
&nbsp;

```applescript
property Maven : application "MailMaven"
tell application "MailMaven"
    tell account "Foo"
        messages of Maven's mailbox "SC/Parent/Child" -- note use of Maven's to access accounts outside of "Foo"
```
</div>


#### Special Mailboxes

You can access special mailboxes using labels `inbox`,`archive mailbox`,`sent mailbox`,`trash mailbox`,`draft mailbox`,`junk mailbox`

```applescript
tell application "MailMaven"
    inbox of account "Foo"
    archive mailbox of account "Bar"
```

You can access unified mailboxes by not specifying the account 

```applescript
tell application "MailMaven"
    messages of inbox whose read status is false
    archive mailboxes  -- returns all the archive mailboxes all accounts.
```

Special mailboxes and named mailboxes can be used as destination for move and copy commands

You can access unified mailboxes by not specifying the account

```applescript
tell application "MailMaven"
    move selected messages of viewer 1 to archive mailbox -- will move each message to its appropriate mailbox
    move selected messages of viewer 1 to archive mailbox of account "foo" -- will move each message to the specific archive.
    tell selected messages of first viewer
        move to archive mailbox  -- moves to each message's respective archive mailbox.
```
#### Smart Mailboxes

You can access the smart mailboxes and access their current messages.

```applescript
tell application "MailMaven"
    messages of smart mailbox "Today's unread messages" whose subject contains "Urgent"
```

At this point in time smart mailboxes and their criteria cannot be edited via AppleScript


&nbsp;

***

## The Message Object

As the message is the item you most like want to find, examine, and manipulate, lets look at the basic properties and elements of messages

First the immutable _properties_.

* subject (the displayed subject -- using the alternate subject if set)
* original subject (the sender's subject -ie subject from headers))
* date received
* content (the text content parsed out from html as necessary)
* all headers  -- string representing the raw header data
* source  -- a string representing the full message data
* id  -- an arbitrary numeric value assigned to a message. This value is subject to change and cannot be used to universally identify a message. 
* conversation id  -- an arbitrary numeric value assigned to the conversation for message. Returns 0 if message is not part of a conversation.  This value is subject to change and cannot be used to universally identify a conversation. 
* parent message  -- the message this message is in reply to
* maven identifier -- a string identifying a message. The same message in different mailboxes of the same account will have the same maven identifier.  -- this will not change. 
* content identifier - a string identifying the contents of a message.  The same message in different accounts will have the same content identifier.  This will not change
* maven url  -- maven-message://...  used for locating a specific message in a specific account.
* message id url -- message-id://... used for locating a message regardless of account. Many apps reference messages usnng the message id url
* mailbox (the mailbox in which resides the message)

    
Addresses associated with the message are "_elements_" as they can be multiple

        Note that email addresses have 3 properties:
            * name        (eg Bob Jones)    
            * address      (eg bjones@somewhere.com)
            * contents  (the RFC address: "Bob Jones" <bjones@somewhere.com>)
    
* recipients (_addresses_)
* cc recipients
* bcc recipients
* from addresses 
* reply to addresses
* sender (may be different than the `from addresses`)

Other elements are

* headers  -- an list of header elements (which include name and header value)
* message attachments
* related messages -- the list of messages that are in the same conversation as examined message.
* child messages -- the list of messages that are replies to examined message.

Next are the mutable properties of a messages. Theses properties are modifiable both in Maven and AppleScript

* flag (no flag, red, orange, yellow, green, blue, purple, grey)
* read status  (true or false)
* junk status  (true or false)
* alternate subject
* keywords  (a list of text) (can be set using a single text object)
* project _text_
* review date  _date_)
* importance (very low, low, normal, high, urgent)
* note _text_

Read-only mutating properties

* deleted status  (true or false)
* modification date _date_ – a date timestamp of when any of the above mutable properties of a message were modified.

Messages also can be the object of the following commands,

* delete
* move to _mailbox_
* copy to _mailbox_
* forward  (see Composing and Sending Messages)
* reply (see Composing and Sending Messages)

&nbsp;

***

## Viewers

As well from the application, you can access the current message viewers, and access both their list of `messages` and `selected messages`

Note that the list of viewers includes regular viewer windows, search windows, and single message windows.  The order of the viewers are the order of the visible windows from frontmost (being viewer 1) to backmost.  If you want the currently viewed messages, use `viewer 1` or `first viewer`.

```applescript
tell application "MailMaven"
    tell first viewer
        count of (every message whose read status is false)
```

```applescript
tell application "MailMaven"
    read status of every selected message of viewer 1
```

New viewers can be created with the `make` command and closed with the `close` command.

&nbsp;

***

## Commands

### Checking new messages

You can check for new mail in all active accounts using the `check new mail` command

```applescript
tell application "MailMaven"
    check new mail
```
or check a specific account or a list of accounts

```applescript
tell application "MailMaven"
    tell account "Main" to check new mail
```

```applescript
tell application "MailMaven"
     check new mail {account "Main", account "Personal"}
```

### Saving Messages

Messages from mailboxes or queries can be saved as eml files or mbox files.

Syntax 

```applescript
save _message_ as `format` in _location_ [with replacing] [with tag data]
```

Single messages can be saved in eml or mbox format

```applescript
save first selected message of viewer 1 as eml format in "/Users/me/Documents/EmailArchive/test.eml" [with replacing] [with tag data]
save last message of sent mailbox of account "Main" as mbox format in "/Users/me/Documents/EmailArchive/test2.mbox" with tag data
```

Multiple messages must be saved in mbox format

```applescript
tell (make query with properties {logic:all criteria})
    make importance criterion with properties {level:very high, qualifier: equals to}
    save messages as mbox format in "/Users/me/Documents/EmailArchive/VeryImportant.mbox" with tag data
```

### Saving Attachments

Attachments of messages can be saved by iterating through the message attachments and saving each.

```applescript
tell application "MailMaven"
    repeat with _attachment in message attachments of first selected message of viewer 1
           save _attachment in "/Users/Me/Document/Exported Attachments/" & (name of _attachment as string) with replacing
    end repeat
end tell
```

***

### Composing and Sending Messages

You can use AppleScript to compose and send message or open them in a composer window for further modification/review.

The basic syntax follows


```applescript
compose new message [from {text|address|account}] [to {text|address}] [using template {text(template name)|new message template}] [with properties {property info}]
```

Note that any compose new message command must occur inside a "with transaction" block similar to queries.  This is to enforce a memory scope for the composition


Examples:

```applescript
with transaction
    set myMessage to compose new message from "test@smallcubed.com" to "\"my friend\" <myfriend@smallcubed.com>";
    tell myMessage
        make cc recipient with address "otherPerson@smallcubed.com"
        set subject to "This is a sample message"
        set content to "lorem ipsum"
        show composer -- or send
    end tell        
end transaction
```

```applescript
with transaction
    -- this example will send a new message to the sender of the current message from my Main account using template
    
    set selectedAddress to (first from address of (first selected message of viewer 1))
    tell (compose new message from account "My Main Account" to selectedAddress using template "Test Template")
        set subject to "This is a sample message"
        set content to "lorem ipsum"
        show composer -- or send
    end tell
end transaction
```


When you specify the from address, Maven will check the address agains your accounts and email aliases. if the address does not correspond to a valid account, it will show an error.

Currently the content is plain text only.   We are looking at rich text, html and markdown.

Multiple and differ types of recipients can be added 

```applescript
    make cc recipient with address "\"Bob Jones\" <bjones@somewhere.com>" 
    make bcc recipient with properties {name: "Bob Jones", address:"nes@somewhere.com"}
    make reply to address with address "no-reply@smallcubed.com"
```

or even removed

```applescript
    delete (every cc recipient whose address contains "smallcubed.com")
```


#### Send or Show Composer Command

The final action of composing a message is to send it, or show the composer.  

Once you show the composer you, applescript turns control over to the interface.  The composed message can no longer be altered, or even sent -- it is up to the user interacting with message to finish the job.


#### Signatures

Signatures can be added by getting the signature beforehand


```applescript
with transaction
    set theSig to first signature whose name is "Witty Quote"
    tell (compose new message from "test@smallcubed.com")
        ...
        set message signature to theSig
```

#### Templates

Templates can be specified when making the Message

```applescript
with transaction
    tell (compose new message from "test@smallcubed.com" using template (first new message template whose name is "Gentle Reminder"))
        ...
```

Note: Adding attachments is in the works.

#### Replies and Forwards

When you are viewing a message you can tell it to reply or forward.  Once you reply/forward you have a composer message to work with.

```applescript
with transaction
    tell first selected message of first viewer 
        tell (reply)
            make bcc recipient with address "secretRecipient@smallcubed.com"
            set content to "Thanks for the info"
            set signature to mySig
            send
```

The content will be placed above the source message content by default (this is controlled by the default reply template)

You can also reply all 

```applescript
with transaction
    set msg to first selected message of first viewer 
    set theReply to reply msg with reply to all 
    tell theReply
        make bcc recipient with address "secretRecipient@smallcubed.com"
        set content to "Thanks for the info"
        show composer
```

and use reply templates

```applescript
with transaction
    set msg to first selected message of first viewer 
    set theReply to reply msg with reply to all using template (first reply template whose name is "product info")
    tell theReply
        make bcc recipient with address "secretRecipient@smallcubed.com"
        set content to "Thanks for the info"
        show composer
```

And forward

```applescript
with transaction
    tell forward  (first selected message of first viewer)
    -- or tell (forward  (first selected message of first viewer) with as attachment)
        make bcc recipient with address "secretRecipient@smallcubed.com"
        set content to "Thanks for the info"
        set signature to mySig
        show composer
```

#### Encryption and Cryptographic Signatures

Set the `encrypt` property of the composer to `true` to indicate the message should be encrypted to each of the recipients.  

If you send the message via the `send` command, Maven will check all the recipient addresses for their public key and display an error if there is 1 or more more recipients for whom the message cannot be encrypted.   The message will not be sent in this case and the `send` command will return false.

When showing in composer, recipient addresses are not checked for their valid encryption keys.  The composer will highlight the problematic addresses.

Set the `digitally sign` property to `true` to digitally sign the message.  If you sent the message via the `send` command, Maven will check to verify that a digital signature can be made for the from address.  The message will not be sent in this case and the `send` command will return false.

Individual addresses can be tested ahead of type by examining the `pgp public key valid` and `pgp secret key valid` properties of the address.

```applescript
    if pgp secret key valid of address "mainAccount@me.com" then
        ...
    end if
```

#### Tags

You can set tags on the message to be delivered using the various tag properties.  (`Keywords`, `Project`, `Review Date`, `Importance`, `Notes`)

By default, the tags will be recorded only on your version of the message (in the sent mailbox). 
Setting the `deliver tags` property to `true` will embed the tags data in the headers of the message and the recipient(s) would be able to adopt tags if they are using Maven.

#### Delivery options

`delivery time` (Date)  By default, Maven set the delivery time by applying  applicable outbox rules.  If there are no rules that alter delivery time, a message sent will be delivered immediately.

Setting the delivery time to a point in the future will deliver at that time (or at first opporuntity afterwards if Maven is not running), regardless of the rules that may alter delivery time. Setting to the current time or a past time will deliver the message immediately regardless of rules that may alter delivery time.

`outbox rules` (list of outbox rules)  By default, Maven will apply all active outbox rules to a message being sent.  Set oubox rules to an empty list (`{}`) if you wish to apply no rule. Specify the outbox rule(s) by reference to apply specific rules. Applying specific rules will ignore the active state of a rule.

`archive mailbox` (mailbox) By default Maven will archive the sent message in the "Sent mailbox" of the account for the "From address" of the message.  Setting the archive mailbox will move the sent to the specified mailbox.

```applescript
with transaction
   tell (compose new message from "mainAccount@me.com" to "test@address.com")
        set subject to "This is a test message"
        set keywords to {"testKeyword"}
        set project to "test project"
        set deliver tags to true
        set delivery date to (current date) + 300 -- send it 5 minutes from now.
        set archive mailbox to Maven's mailbox "Main/test messages"
        send
    end tell

```

&nbsp;

***

## Queries

Up to now, Scripting Maven would be very similar to scripting Mail.  To find messages you iterate through accounts, mailboxes and messages and collect the messages whose properties meet certain criteria.

This can be slow and resource intensive as each message has to be loaded in memory to have its properties examined.

With MailMaven, there is a better way:  Queries

A __`query`__ is a temporary AppleScript-based smart mailbox that can are used to locate and examine/manage messages.

Using queries is _much_ faster than iterating through messages in mailboxes because it uses Maven's internal mechanism for locating messages that smart mailboxes use.

Depending on the scope and nature of an iterative search which can take seconds, if not minutes to process, queries are almost instantaneous.  Moreover, queries have the ability to sort messages and filter duplicates

Here is a sample query

```applescript
property Maven : application "MailMaven"
tell application "MailMaven"
    -- first the "old way" to show comparison
    set totalUnread to 0
    repeat with mbx in (every mailbox whose name is "inbox")
        set totalUnread to totalUnread + (count of (every message of mbx whose read status is false))
    end repeat
    
    -- and now the Query
    with transaction
        tell (make query with properties {logic:all criteria})
            make mailbox criterion with properties {mailbox:Maven's inbox}
            make read status criterion with properties {is read:false}
            set totalUnread to count of messages
        end tell
    end transaction
end tell
```

Whew!  There is a bit there so lets break it down.

What with this ```with transaction ... end transaction```?

This is a bit of an AppleScript hack.   Whenever you make an object such as a query in AppleScript, the object will stick around in memory until you deliberately delete it even after the script has finished. -- which can be easy to forget to do.   As well other scripts running may have access to that object because it is stored in a "public memory space" that all AppleScripts use.

`With transaction... end transaction` puts a scope on the the query.   When the execution of the script hits the `end transaction`, all the queries created since the start (`with transaction`) will be removed automatically. Note that AppleScript editors will automatically insert the `end transaction` and the script will not compile if it is removed.

Moreover, Maven keeps transactions separate so that different transactions don't have access to access to each others queries.

If you try to make a new query without the `with transaction`, the script editor will display an error letting you know it is needed.

```applescript
-- the following script will compile but on running there will be an error informing you with transaction is needed
tell (make query with properties {logic:all criteria})
    make read status criterion with properties {is read:false}
```

The `logic` property tell the criteria how to join all the children criteria;

Once the child criterion is set you can access the messages element of the query and use them the same as if you got those messages from a viewer, or by accessing a mailbox.

### Criterion Quick Reference

There are numerous criteria you can use for locating messages.  Criteria are combined based on the logic of the query (or parent compound criterion)

The following table provides essential information about each  criteria, including qualifiers to be used and fields that can be evaluated. (further details on criteria and qualifiers can be viewed in the script dictionary)



| Criterion | `Fields:` _types_ (default) | Qualifier: (default) <br/><span style="font-weight: normal !important">See Criterion Qualifiers section below</span>|
|---|---|---|
|**Container Criteria**|||
| `mailbox criterion` | `mailbox:` _mailbox_ \| _smart mailbox_ \| _text_| `equals qualifier` (`equals to`) |
| `account criterion` | `account:` _account_ \| _text_| `equals qualifier` (`equals to`)|
|**Message Attributes Criteria**|||
| `subject criterion` <br/>The current subject of a message| `subject:` _text_ | `criterion qualifier` (`includes`) |
| `original subject criterion`<br/>The subject the sender gave a message | `subject:` _text_| `criterion qualifier` (`includes`) |
| `alternate subject criterion`<br/>The subject you gave a received message | `subject:` _text_ | `criterion qualifier` (`equals to`)|
| `content criterion`<br/>The text body of a message | `phrase:` _text_| `criterion qualifier` (`includes`)|
| `message date criterion`<br/>The date of a message with a specified date range | `start date:` _date_ \| <br/>&emsp;`distant past` (default)<br/>`end date:` _date_ \|<br/>&emsp;`distant future` (default)| — |
| `relative message date criterion`<br/>The date of a message relative to today|  `starting at:` _integer_<br/>`unit:` _date unit_ (`day`)<br/> `for duration:` _integer_ (1)| — |
| `attachment name criterion` | `attachment name:` _text_ | `criterion qualifier` (`equals to`) |
| `attachment count criterion` | `attachment count:` _integer_ | `numeric comparison qualifier` (`equals to`) |
| `message size criterion` | `size:` _integer_ | `numeric comparison qualifier` |
| `encryption criterion` | `encryption status:` `not evaluated`(default) \|<br/>&emsp;`encrypted`\|`not encrypted`, <br/> `signature status:` `not evaluated`(default) \|<br/>&emsp;`not signed`\|`signed`\|<br/>&emsp;`valid signature`\|`invalid signature` | – |
| **Message Collection Criteria**||| 
| `message criterion` | `message:` _message_ \| _list of messages_ \| _integer_ \| _list of integer_ <br/>integers refer to message ids| `includes qualifier` (`includes`)|
| `conversation criterion` | `message:` _message_ \| _list of messages_<br/>`conversation id:` _integer_ \| _list of integer_ | `includes qualifier` (`includes`) \| `member qualifier`|
| **Address Criteria** field `address:` allow _address_ or _text_ values<br/>When an _address_ is provided, qualifier default to `equals to` otherwise `includes` |||
| `from criterion` | `address:` _address_ \| _text_ | `criterion qualifier` |
| `to recipient criterion` | `address:` _address_ \| _text_| `criterion qualifier` |
| `cc recipient criterion` | `address:` _address_ \| _text_ | `criterion qualifier` |
| `any recipient criterion` | `address:` _address_ \| _text_| `criterion qualifier` |
|**Status Criteria**|||
| `read status criterion` | `is read:` _boolean_ | — |
| `junk status criterion` | `is junk:` _boolean_ | — |
| `answered criterion` | `was answered:` _boolean_ | — |
| `forwarded criterion` | `was forwarded:` _boolean_ | — |
|**Metadata Criteria**|||
| `flag criterion` | `flag:` `no flag` \|<br/>&emsp;`red`\|`orange`\|`yellow`\|<br/>&emsp;`green`\|`blue`\|`purple`\|`grey` | `equals qualifier` (`equals to`) |
| `keyword criterion` | `keyword:` _text_ | `criterion qualifier` (`equals to`) |
| `tagged status criterion` | `is tagged:` _boolean_ | — |
| `note criterion` | `notes:` _text_| `criterion qualifier` (`includes`) |
| `project criterion` | `project:`_text_ | `criterion qualifier` (`equals to`) |
| `importance criterion` | `level:` `no importance` \|<br/>&emsp;`very low` \| `low` \| `moderate` \|<br/>&emsp;`high` \| `very high`| `numeric comparison qualifier` (`equals to`) |
| `review date criterion`<br/>Review Date with a specified date range | `start date:` _date_ \| <br/>&emsp;`distant past` (default)<br/>`end date:` _date_ \|<br/>&emsp;`distant future` (default)| — |
| `relative review date criterion`<br/>Review Date relative to today |  `starting at:` _integer_<br/>`unit:` _date unit_ (default:`day`)<br/>`for duration:` _integer_ (1) | — |
| `modification date criterion`<br/>Date message was last modified | `start date:` _date_ \| <br/>&emsp;`distant past` (default)<br/>`end date:` _date_ \|<br/>&emsp;`distant future` (default)| — |
|&nbsp;|||
| `compound criterion` | `logic:` `all criteria`\|<br/>&emsp;`any criteria`\|<br/>&emsp;`no criteria` | — |



### Criterion Qualifiers

Each criterion type uses one of four qualifier enumerations. The qualifier controls how the expression is matched.

**`criterion qualifier`** — used by most criteria (keyword, subject, from, note, project, attachment name, content etc.)

| Qualifier | Meaning |
|---|---|
| `equals to` | Exact match |  
| `not equal to` | Does not match exactly |
| `includes` | Contains the expression as a substring |
| `does not include` | Does not contain the substring |
| `has prefix` | Begins with the expression |
| `has suffix` | Ends with the expression |
| `matches regex` | Matches a regular expression |
| `is set` | Field has any value |
| `is not set` | Field has no value |

**`numeric comparison qualifier`** — used by criteria with numeric values (attachment count, importance)

| Qualifier | Meaning |
|---|---|
| `equals to` | Exact numeric match |
| `not equal to` | Not equal |
| `less` | Strictly less than |
| `greater` | Strictly greater than |
| `less or equal` | Less than or equal |
| `greater or equal` | Greater than or equal |

**`equals qualifier`** — used by criteria with fixed values (flag, encryption status)

| Qualifier | Meaning |
|---|---|
| `equals to ` | Matches the specified value |
| `not equal to` | Does not match the specified value |


**`includes qualifier`** — used by criteria with inclusion status (eg message criterion, conversation criterion)

| Qualifier | Meaning |
|---|---|
| `includes ` | includes the specified value(s) |
| `does not include` | Does not include the specified value(s) |

**`member qualifier`** — used by criteria with member status (eg conversation criterion, address group critierion)

| Qualifier | Meaning |
|---|---|
| `is member ` | is member of a conversation / is member of _group name_ |
| `is not member` | is not member... |

**Boolean criteria** (read status, junk, tagged status, answered, forwarded) — no qualifier needed; they use a boolean property directly (`is read`, `is junk`, `is tagged`, `was answered`, `was forwarded`).

**Date criteria** (date,  date offset , review date, review date offset) — no qualifier; they use date properties directly (`start date`, `end date`, `unit`, `starting at`, `for duration`).


### Nesting Criteria

Criteria can be nested using a `compound criterion` to generate complex queries involving multiple logics:

```applescript
property Maven : application "MailMaven"
- find all messages in the inboxes of all accounts that are either unread or received in the last 3 days
-- ie (any inbox AND (unread OR received in last 3 days))
tell (make query with properties {logic: all criteria})
   make mailbox criterion with properties {mailbox:Maven's inbox}
    -- create a OR compound and add subcriteria to it using a tell block
    tell (make compound criterion with properties {logic: any criteria}
        make read status criterion with properties {is read:false}
        make date offset criterion with properties {unit:day, starting at:-3, for duration:3}
    end tell
     messages  -- Accessing the messages must occur at root query level (ie outside of nested compounds; 
```


### Query Properties: `Sort Order` & `Sort Ascending`

You can use the `sort order` and `sort ascending` properties of a query to specify sort order fo the resultant messages

```applescript
property Maven : application "MailMaven"
tell (make query with properties {logic:all criteria})
   make mailbox criterion with properties {mailbox:Maven's inbox}
    make project criterion with properties {project: "Budget 2027"}
    set sort order to sort by date received
     set sort ascending to false 
    messages -- will  return all the unread messages of the inbox sorted by date from newest to oldest.
        set sort order to sort by subject
     set sort ascending to true 
    nessages -- returns same message sorted by subject from A-Z
```

#### Query Property: `Exclude duplicates`

By default a query's results will remove duplicate messages. This can be turned off with the `exclude duplicates` property

```applescript
property Maven : application "MailMaven"
tell (make query with properties {logic:all criteria})
   make mailbox criterion with properties {mailbox:Maven's inbox}
    make keyword criterion with properties {keyword: "pending"}
    count of messages --  returns 10
     set exclude duplicates to false
    count of messages --  returns 12
```

#### Query Result Stability

Because a scripter is likely to interact with results of a query over several script lines, Maven ensures that the results are stable.  That is, the messages that constitute the results will NOT change over time even those actions may take place to exclude messages or add new messages to the query.

for example, If you change the flag of messages returned by a flag criterion, the message will not be excluded from further interactions with to the query.
 
```applescript
tell (make query with properties {logic:all criteria})
   make mailbox criterion with properties {mailbox: inbox}
    make flag criterion with properties {flag: green}
     count of messages --  returns 10
    flag of item 3 of messages -> returns green
    set flag of item 3 of messages to blue
     count of messages --  still returns 10!
    flag of item 3 of messages -> returns blue!   -- message was updated but not result set for query
```

#### Refreshing results

Query results can be refreshed with the `refresh results` command sent to the query.

```applescript
tell (make query with properties {logic:all criteria})
   make mailbox criterion with properties {mailbox: Maven's inbox}
    make importance criterion with properties {level:high, qualifier: greater or equal}
     count of messages --  returns 10
    subject of item 3 of messages -> returns "Message C"
    set importance of item 3 of messages to low 
     count of messages --  still returns 10!
    refresh results
     count of messages --  now returns 9!
    subject of item 3 of messages -> returns "Message D"
```

Additionally, the results of a query are automatically refreshed whenever the criteria change or specifics of a criterion changes or the exclude duplicates property changes.  However is sort order is changed, the current results are resorted without refreshing the results.

Note that accessing messages of mailboxes and viewers __are not__ stable the same way that queries are.  
Scripts that depend on message list stability over multiple commands should use message queries.

&nbsp;

***

## Run AppleScript Rule Action

Scripts can be performed from Maven's inbox, outbox, and keystroke rules via the Run AppleScript Rule action.

To configure, aave a compiled script in the "~/Library/Application Scripts/com.smallcubed.mailmaven" folder.
Then select the script from Script menu in the Run AppleScript Rule action.  (You can add up to five differen AppleScripts to a rule.)

From the script menu you can also view the script folder and create a new script from a template.

Create a Rule Action script using the template

```applescript
using terms from application "MailMaven"  -- this is necessary to ensure that the "perform applescript action is understood when editing the script
    to perform applescript action 
        tell application "MailMaven"
            (* Vocabulary Notes:
                 messages of rule action:   gets the messages targeted by rule
                 rule of rule action:  gets the rule being performed.
                
                 Consult MailMaven's script dictionary for a full vocabulary
            *)
            repeat with _msg in messages of rule action
                -- do your magic here 
            end repeat
        end tell
    end perform applescript action
end using terms from
```

For example

```applescript
using terms from application "MailMaven"
    to perform applescript action
        tell application "MailMaven"
            repeat with _msg in messages of rule action
                if subject of _msg contains "Bitcoin" then
                    set junk status of _msg to true
                end if
            end repeat
        end tell
    end perform applescript action
end using terms from

```

Sample script


```applescript
(* 
    Messages from support system (Zendesk) always have the x-Mailer header "Zendesk Mailer".
    and have predictable text that contain the ticket number and possibly the Redmine (our backend bug tracker) issue number.
    
    The following script will find messages from Zendesk and 
    use a regular expression to extract the support ticket number and set it as a keyword.
    It will also look for any associated redmine tickets and add them as to the keywords.
    
*)
use framework "Foundation"
using terms from application "MailMaven"
    to perform applescript action
        tell application "MailMaven"
            repeat with aMsg in messages of rule action
                set xMailer to  first header of aMsg whose name is "X-Mailer"
                if (header value of xMailer) is "Zendesk Mailer" then
                    set ticketNumber to my match(content of aMsg, "\\nTicket #(\\d+)\\n", 1)
                    set kws to {}
                    if ticketNumber is not missing value then
                        set kws to kws & {"ZD " & ticketNumber}
                    end if
                    set redmineNum to my match(content of aMsg, "Assigned to Redmine issue #(\\d+)", 1)
                    if redmineNum is not missing value then
                        set kws to kws & {"🐞 " & redmineNum}
                    end if
                    set keywords of aMsg to kws
                end if
            end repeat
        end tell
    end perform applescript action
end using terms from

on match(_text, _regex, _captureGroup)
    set aString to current application's NSString's stringWithString:_text
    set {theExpr, theError} to current application's NSRegularExpression's regularExpressionWithPattern:_regex options:0 |error|:(reference)
    set theMatches to theExpr's matchesInString:aString options:0 range:{0, aString's |length|()}
    if (count of theMatches) ≥ _captureGroup then
        set aMatch to item _captureGroup of theMatches
        set theRange to (aMatch's rangeAtIndex:1)
        set theResult to (aString's substringWithRange:theRange) as text
        return theResult
    else
        return missing value
    end if
end match
```

## Using MailMaven with Large Language Models

Everything described above can also be driven by a large language model instead of hand-written scripts. **[MailMavenMCP](https://github.com/Smallcubed/MailMavenMCP)** is a Model Context Protocol (MCP) server that leverages MailMaven's AppleScript capabilities to allow AI apps such as Claude to interact with Maven using natural language instructions. With the MavenMCP you can search messages, read and organize mail, and composing new ones.

Under the hood the MCP builds on the same vocabulary and patterns covered in this guide (queries and criteria, transactions for object creation, message identifiers, and so on), so this document remains a useful reference if you want to understand what the MCP server is doing, extend it, or write your own AppleScript-based tools.

See the [MailMavenMCP repository](https://github.com/Smallcubed/MailMavenMCP) for setup instructions and the full list of available tools.

&nbsp;

***

## Acknowledgements

Many thanks to Andrew Hirst for his contributions to this guide and for assistance in testing AppleScript functionality in MailMaven.
