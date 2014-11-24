PW-Gen-Excel-VBA
================
If you need to relative securely lock and share Excel sheets. Can be used for rudimentary signing procedures.
These scripts can lock an excel worksheet with a random password, alter the password, to dilute it and generate a different string for reverse this dilution and the password itself.
Additionally the password can be splitted for different parties to obtain a part of the key, so no unconsesual changes can be made.

### Contents
  1. Implementation / Setup
  2. basic workflow
  3. extended recommended workflow

#### implementation / setup
Just import the .bas-files into Excel VBA and you are ready to go.

#### basic workflow
  1. create a random password with `Hash.GetString({#}, {#})`, for a pw # characters long
  2. lock your Excel Worksheet with this random pw
  3. add some salt and pepper to your `pw` with `Hash.SaltAndPepper({YourPW})` IMPORTANT: in the process a decode string is generated and can be saved anywhere. At the moment it is stored inside the workbook, please modify your code at the marked position!
  4. hash the salted and peppered string with a hash `x` characters long `Hash.Hash{X}({SaPStr})` and save the hash for identification procedures

To regain the pw you strap the `SaPStr` of any unused characters by using `Hash.StrapString({SaPStr}, {decodeStr})`, the decode String is generated under 3.

This workflow does not provide any security measures in traditional sense. It is critical to implement it into your sharing workflow.

Fields of application could be storing the keys in two different locations or sharing a worksheet, which's password cannot be guessed.

#### extended recommended workflow
  1. Lock a Worksheet with `Tool.Locker` (please modify any paths and number of recipients)
  2. The chosen worksheet will be locked and the key will be splitted into one part for every recipient
  3. A MS Outlook instance will open with the prepared e-mails for every recipient and two links to sign the document or deny the request. The locked document should be attached, enter e-mail addresses and such and send the e-mails.
  4. As a recipient you can click on one of the links, which opens a new e-mail with your part of the SaP-pw included, send it
  5. All signatures are collected by the issuer
  6. If all keys came back get them into the right order (maybe you mark the e-mails from the start) and reconsolidate the key. Keys can be read directly from an e-mail with `Tool.ReadKey({e-mail's path})`
  7. Unlock the Worksheet with `Tool.Unlocker({reconsolidated key}, {hash})`

This workflow provides some security measures, because the password itself is not directly stored on the issuers pc, but can hardly be altered for other purposes than gaining approval of a few recipients. Not the issuer and no recipient alone can unlock and alter the worksheet.
