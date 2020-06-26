# Outlook Standard E-Mail Signature

Automatically create Outlook Standard Signature through collecting attributes from Active Directory.

[![Travis](https://img.shields.io/travis/rust-lang/rust.svg)](https://github.com/donkey/systeminfo)
[![GitHub issues](https://img.shields.io/github/issues/donkey/systeminfo.svg)](https://github.com/donkey/systeminfo/issues)
[![GitHub forks](https://img.shields.io/github/forks/donkey/systeminfo.svg)](https://github.com/donkey/systeminfo/network)
[![GitHub stars](https://img.shields.io/github/stars/donkey/systeminfo.svg)](https://github.com/donkey/systeminfo/stargazers)
[![Github All Releases](https://img.shields.io/github/downloads/atom/atom/total.svg)](https://github.com/donkey/systeminfo)
[![Scrutinizer](https://img.shields.io/scrutinizer/g/filp/whoops.svg)](https://github.com/donkey/systeminfo)
[![AUR](https://img.shields.io/aur/license/yaourt.svg)](https://github.com/donkey/systeminfo)

## Preface

The purpose of this development is to deploy and distribute a uniform e-mail signature to any employees of a company, this in using Microsoft Outlook, in the sense of the corporate identity of the organization.

Within a organisation, system administrators are responsible for managing the signatures for the corporate identity. Most companies want to design the entire appearance of the signature.

Outlook does not offer any direct options for this, since the email signature in Outlook is a client-side application and users can therefore create and change their own signature. However, there is the option to prevent access to the signature options using Group Policy. But, this does not resolve the problem of first creating and generating a standard signature and making it available to users.

Where a centralized solution at the Outlook client level is preferred to distribute them via login script or GPOs, and you don't want to use an Exchange Transport Rules, or there is no Exchange Server available, this article can be used to give an approach to centralized solutions.

## VBScript Outlook signature deployment

The following VBScript creates an Outlook signature for clients, user data (attributes) are read from Active Directory and inserted into the signature, after which the script adds the signature in new e-mails and replies.

The code lines are paste into an editor, for example Notepad or vs code and save it with file extension .vbs.

## Adjustments
Be free and change your signature to match your corporate identity. Adjustments to the signature can be made in the script between **_BOF signature_** and **_EOF signature_**.

## Run the Script 
On a workstation with Outlook 2013 / 2016 / 2019 / Office 365, during execution there is no output on the screen, after hit the script, open Outlook and create a new email, the signature now appears in the message, which looks similar to the figure below.

## Screenshot
![Outlook E-Mail Standard Signature](https://think.unblog.ch/wp-content/uploads/2020/06/outlook-signature.png)

## Outlook signature client distribution

Scripts for login can be configured via GPOs. The settings for login and logout scripts can be found under User Configuration => Policies => Windows Settings => Scripts. The location for scripts that are assigned via GPOs is located under the path:
```
\\FQDN\SYSVOL\FQDN\policies\user\scripts\logon
```

## Feedback

If you have problems, questions, ideas or suggestions, please contact my by posting to a suitable [mail](http://think.unblog.ch/sty-in-touch)

## Git
```
git clone https://github.com/donkey/outlooksignature.git
```
## Addendum

This script is intentional developed in not very structured way, so it is simply to modify individual lines or omit them altogether, it should be easily customizable.

## license

donkey/outlooksignature is licensed under the MIT License.
