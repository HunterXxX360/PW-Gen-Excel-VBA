PW-Gen-Excel-VBA
================

If you need to relative securely lock and share Excel sheets. Can be used for rudimentary signing procedures.

Create and hash passwords in module "Hash"

basic recommended workflow:
  1. create a random password with Hash.GetString(<#>, <#>), for a pw # characters long
  2. lock your Excel Worksheet with this random pw
  3. add some salt and pepper to your pw with Hash.SaltAndPepper(<YourPW>)
      a) in the process a decode string is generated and can be saved anywhere. At the moment it is stored inside the
          workbook, please modify your code here
  4. hash the salted and peppered string with Hash.Hash<X>(<SaPStr>) and save the hash for identification procedures

To regain the pw you strap the SaPStr of any unused characters by using Hash.StrapString(<SaPStr>, <decodeStr>), the decode String is generated under 2.a).

This workflow does not provide any security measures in traditional sense. It is critical to implement it into your sharing workflow.

Fields of application could be storing the keys in two different locations or sharing a worksheet, which's password cannot be guessed.

The extended workflow will be uploaded soon...
