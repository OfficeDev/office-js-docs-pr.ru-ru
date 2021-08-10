---
title: Извлечение строк сущностей из элемента Outlook
description: Узнайте, как извлечь строки сущностей из элемента Outlook в надстройке Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: b206f76ff1b817ff64970f59cfc79752aad943b3b71a40d86a3ee3d6f6493a3f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092826"
---
# <a name="extract-entity-strings-from-an-outlook-item"></a>Извлечение строк сущностей из элемента Outlook

В этой статье рассказано, как создать надстройку Outlook **для отображения сущностей**, которая извлекает экземпляры строк поддерживаемых известных сущностей в теме и основном тексте выбранного элемента Outlook. Этим элементом может быть встреча, электронное сообщение, приглашение на собрание, ответ на такое приглашение или отказ от него.

Поддерживаемые сущности:

- **Address**. Почтовый адрес США, который содержит по крайней мере подмножество элементов, включающее номер дома, название улицы, город, штат, а также почтовый индекс.
    
- **Contact**. Контактные данные лица в контексте других сущностей, например адреса или названия организации.
    
- **Email address**. SMTP-адрес электронной почты.
    
- **Meeting suggestion**. Приглашение на собрание, например ссылка на мероприятие. Обратите внимание на то, что извлечение приглашений поддерживается только для сообщений, но не для встреч.
    
- **Phone number**. Телефонный номер Северной Америки.
    
- **Task suggestion**. Предложение задачи, которое обычно выражается фразой с действиями.
    
- **URL**.
    
Большинство из этих сущностей зависят от распознавания естественного языка, которое основывается на обработке компьютером больших объемов данных. Это распознавание недетерминированное и иногда зависит от контекста в элементе Outlook.

Outlook активирует надстройку для работы с сущностями каждый раз, когда пользователь выбирает встречу, электронное письмо, приглашение на собрание, ответ на приглашение на собрание или отказ от приглашения на собрание для просмотра. Во время инициализации в примере надстройки для работы с сущностями выполняется считывание всех экземпляров поддерживаемых сущностей из текущего элемента. 

Надстройка предоставляет кнопки, с помощью которых пользователь может выбрать тип сущности. Когда пользователь выбирает какую-либо сущность, надстройка отображает экземпляры выбранной сущности в области надстройки. В последующих разделах имеются манифест в формате XML, HTML- и JavaScript-файлы надстроек сущностей, а также выделен код, поддерживающий извлечение соответствующих сущностей.

## <a name="xml-manifest"></a>XML-манифест

Надстройка для работы с сущностями использует два правила активации, объединенных логической операцией ИЛИ. 

```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

Эти правила определяют, что Outlook должен активировать надстройку, если в области чтения или инспекторе просмотра выбрана встреча или сообщение (включая письмо или приглашение на собрание, ответ на приглашение или отмену собрания).

Ниже приведен манифест надстройки для работы с сущностями. В нем используется схема версии 1.1 для манифестов надстроек Office.

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
xsi:type="MailApp">
  <Id>6880A140-1C4F-11E1-BDDB-0800200C9A68</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>EN-US</DefaultLocale>
  <DisplayName DefaultValue="Display entities"/>
  <Description DefaultValue=
     "Display known entities on the selected item."/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify the web -->
        <!-- server where the HTML file is hosted. -->
        <SourceLocation DefaultValue=
          "http://webserver/default_entities/default_entities.html"/>
        <RequestedHeight>350</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <!-- Activate the add-in if the current item in Outlook is -->
  <!-- an email or appointment item. -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>
```


## <a name="html-implementation"></a>Реализация HTML

HTML-файл надстройки для работы с сущностями определяет кнопки, позволяющие пользователю выбрать каждый тип сущности, и одну кнопку для очистки отображаемых экземпляров сущности. В нем есть JavaScript-файл, default_entities.js, который описан в следующем разделе [Реализация JavaScript](#javascript-implementation). JavaScript-файл содержит обработчики событий для каждой кнопки.

Обратите внимание, что все надстройки Outlook должны включать файл office.js. Приведенный ниже HTML-файл включает файл office.js версии 1.1 в CDN. 

```html
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" >
    <title>standard_item_properties</title>
    <link rel="stylesheet" type="text/css" media="all" href="default_entities.css" />
    <script type="text/javascript" src="MicrosoftAjax.js"></script>
    <!-- Use the CDN reference to Office.js. -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    <script type="text/javascript"  src="default_entities.js"></script>
</head>

<body>
    <div id="container">
        <div id="button">
        <input type="button" value="clear" 
            onclick="myClearEntitiesBox();">
        <input type="button" value="Get Addresses" 
            onclick="myGetAddresses();">
        <input type="button" value="Get Contact Information" 
            onclick="myGetContacts();">
        <input type="button" value="Get Email Addresses" 
            onclick="myGetEmailAddresses();">
        <input type="button" value="Get Meeting Suggestions" 
            onclick="myGetMeetingSuggestions();">
        <input type="button" value="Get Phone Numbers" 
            onclick="myGetPhoneNumbers();">
        <input type="button" value="Get Task Suggestions" 
            onclick="myGetTaskSuggestions();">
        <input type="button" value="Get URLs" 
            onclick="myGetUrls();">
        </div>
        <div id="entities_box"></div>
    </div>
</body>
</html>
```


## <a name="style-sheet"></a>Таблица стилей


В надстройке для работы с сущностями используется дополнительный файл таблицы стилей default_entities.css, который определяет макет выходных данных. Ниже приведен листинг CSS-файла.


```CSS
*
{
    color: #FFFFFF;
    margin: 0px;
    padding: 0px;
    font-family: Arial, Sans-serif;
}
html 
{
    scrollbar-base-color: #FFFFFF;
    scrollbar-arrow-color: #ABABAB; 
    scrollbar-lightshadow-color: #ABABAB; 
    scrollbar-highlight-color: #ABABAB; 
    scrollbar-darkshadow-color: #FFFFFF; 
    scrollbar-track-color: #FFFFFF;
}
body
{
    background: #4E9258;
}
input
{
    color: #000000;
    padding: 5px;
}
span
{
    color: #FFFF00;
}
div#container
{
    height: 100%;
    padding: 2px;
    overflow: auto;
}
div#container td
{
    border-bottom: 1px solid #CCCCCC;
}
td.property-name
{
    padding: 0px 5px 0px 0px;
    border-right: 1px solid #CCCCCC;
}
div#meeting_suggestions
{
    border-top: 1px solid #CCCCCC;
}
```


## <a name="javascript-implementation"></a>Реализация JavaScript

В следующих разделах описано, как этот пример (файл default_entities.js) извлекает известные сущности из темы и текста сообщения или встречи, которую просматривает пользователь.

## <a name="extracting-entities-upon-initialization"></a>Извлечение сущностей при инициализации

Когда происходит событие [Office.initialize](/javascript/api/office#Office_initialize_reason_), надстройка для работы с сущностями вызывает метод [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) текущего элемента. Метод `getEntities` возвращает глобальной переменной `_MyEntities` массив экземпляров поддерживаемых сущностями. Ниже представлен соответствующий код JavaScript.


```js
// Global variables
var _Item;
var _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}

```


## <a name="extracting-addresses"></a>Извлечение адресов


Когда пользователь нажимает кнопку **Get Addresses** (Получить адреса), обработчик событий `myGetAddresses` получает массив адресов из свойства [addresses](/javascript/api/outlook/office.entities#addresses) объекта `_MyEntities` (если был извлечен хотя бы один адрес). Каждый извлеченный адрес хранится в массиве в виде строки. Чтобы отобразить список извлеченных URL-адресов, обработчик событий `myGetAddresses` формирует локальную HTML-строку в `htmlText`. Ниже представлен соответствующий код JavaScript.


```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-contact-information"></a>Извлечение контактных данных


Когда пользователь нажимает кнопку **Получить** контактные сведения, обработник событий получает массив контактов вместе с их сведениями из свойства контактов объекта, если таково было `myGetContacts` [](/javascript/api/outlook/office.entities#contacts) `_MyEntities` извлечено. Каждый извлеченный контакт хранится в виде объекта [Contact](/javascript/api/outlook/office.contact) в массиве. Обработчик событий `myGetContacts` получает дополнительные данные о каждом контакте. Обратите внимание, что контекст определяет, может ли Outlook извлечения контакта из элемента подпись в конце сообщения электронной почты, или хотя бы некоторые из следующих сведений должны существовать в непосредственной близости от &mdash; контакта.


- Имя контакта из свойства [Contact.personName](/javascript/api/outlook/office.contact#personName).

- Название компании, связанное с контактом, из свойства [Contact.businessName](/javascript/api/outlook/office.contact#businessName).

- Массив номеров телефонов, связанных с контактом, из свойства [Contact.phoneNumbers](/javascript/api/outlook/office.contact#phoneNumbers). Каждый номер телефона представлен объектом [PhoneNumber](/javascript/api/outlook/office.phonenumber).

- Строка, представляющая телефонный номер из свойства [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phoneString) для каждого элемента **PhoneNumber** в массиве телефонных номеров.

- Массив URL-адресов, связанных с контактом, из свойства [Contact.urls](/javascript/api/outlook/office.contact#urls). Каждый URL-адрес представлен в виде строки в элементе массива.

- Массив адресов эл. почты, связанных с контактом, из свойства [Contact.emailAddresses](/javascript/api/outlook/office.contact#emailAddresses). Каждый адрес эл. почты представлен в виде строки в элементе массива.

- Массив почтовых адресов, связанных с контактом, из свойства [Contact.addresses](/javascript/api/outlook/office.contact#addresses). Каждый почтовый адрес представлен в виде строки в элементе массива.

Чтобы отобразить данные каждого контакта, обработчик событий `myGetContacts` формирует локальную HTML-строку в `htmlText`. Ниже представлен соответствующий код JavaScript.




```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    var htmlText = "";

    // Gets an array of contacts and their information.
    var contactsArray = _MyEntities.contacts;
    for (var i = 0; i < contactsArray.length; i++)
    {
        // Gets the name of the person. The name is a string.
        htmlText += "Name : <span>" + contactsArray[i].personName +
            "</span><br/>";

        // Gets the company name associated with the contact.
        htmlText += "Business : <span>" + 
        contactsArray[i].businessName + "</span><br/>";

        // Gets an array of phone numbers associated with the 
        // contact. Each phone number is represented by a 
        // PhoneNumber object.
        var phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (var j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        var urlsArray = contactsArray[i].urls;
        for (var j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        var emailAddressesArray = contactsArray[i].emailAddresses;
        for (var j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        var addressesArray = contactsArray[i].addresses;
        for (var j = 0; j < addressesArray.length; j++)
        {
          htmlText += "Address : <span>" + addressesArray[j] + 
              "</span><br/>";
        }

        htmlText += "<hr/>";
        }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-email-addresses"></a>Извлечение электронных адресов


Когда пользователь нажимает кнопку **Get Email Addresses** (Получить электронные адреса), обработчик события `myGetEmailAddresses` получает массив SMTP-адресов электронной почты из свойства [emailAddresses](/javascript/api/outlook/office.entities#emailAddresses) объекта `_MyEntities` (если был извлечен хотя бы один адрес). Каждый извлеченный электронный адрес сохраняется в массиве в виде строки. Для отображения списка извлеченных электронных адресов обработчик событий `myGetEmailAddresses` формирует локальную HTML-строку в `htmlText`. Ниже приведен соответствующий код JavaScript.


```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses() {
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++) {
        htmlText += "E-mail Address : <span>" + emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-meeting-suggestions"></a>Извлечение приглашений на собрания


Когда пользователь нажимает кнопку **Get Meeting Suggestions** (Получить приглашения на собрания), обработчик событий `myGetMeetingSuggestions` получает массив приглашений на собрания из свойства [meetingSuggestions](/javascript/api/outlook/office.entities#meetingSuggestions) объекта `_MyEntities` (если было извлечено хотя бы одно приглашение).


 > [!NOTE]
 > Только сообщения, но не встречи поддерживают `MeetingSuggestion` тип сущности.

Каждое извлеченное приглашение на собрание хранится в виде объекта [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) в массиве. Обработчик событий `myGetMeetingSuggestions` получает дополнительные данные о каждом приглашении на собрание:


- Приглашение на собрание из свойства [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingString).

- Массив участников собрания из свойства [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees). Каждый участник представлен объектом [EmailUser](/javascript/api/outlook/office.emailuser).

- Имя из свойства [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayName) для каждого участника.

- SMTP-адрес из свойства [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailAddress) для каждого участника.

- Предлагаемое место проведения собрания из свойства [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location).

- Предлагаемая тема собрания из свойства [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject).

- Предлагаемое время начала собрания из свойства [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start).

- Предлагаемое время окончания собрания из свойства [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end).

Чтобы отобразить данные каждого приглашения на собрание, обработчик событий `myGetMeetingSuggestions` формирует локальную HTML-строку в `htmlText`. Ниже представлен соответствующий код JavaScript.




```js
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions() {
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++) {
        // Gets the string that was identified as a meeting suggestion.
        htmlText += "MeetingString : <span>" + meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++) {
            htmlText += "Attendee : ( ";

            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-phone-numbers"></a>Извлечение телефонных номеров


Когда пользователь нажимает кнопку **Get Phone Numbers** (Получить телефонные номера), обработчик событий `myGetPhoneNumbers` получает массив телефонных номеров из свойства [phoneNumbers](/javascript/api/outlook/office.entities#phoneNumbers) объекта `_MyEntities` (если был извлечен хотя бы один номер). Каждый извлеченный номер сохраняется в качестве объекта [PhoneNumber](/javascript/api/outlook/office.phonenumber) в массиве. Обработчик событий `myGetPhoneNumbers` получает дополнительные данные о каждом телефонном номере.


- Строка, представляющая тип номера телефона (например, домашний номер) из свойства [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type).

- Номер телефона из свойства [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phoneString).

- Исходный номер телефона из свойства [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalPhoneString).

Чтобы отобразить данные каждого номера телефона, обработчик событий `myGetPhoneNumbers` формирует локальную HTML-строку в `htmlText`. Ниже представлен соответствующий код JavaScript.




```js
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
           "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
            phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="extracting-task-suggestions"></a>Извлечение предложений задач


Когда пользователь нажимает кнопку **Get Task Suggestions** (Получить предложения задач), обработчик событий `myGetTaskSuggestions` получает массив предложений задач из свойства [taskSuggestions](/javascript/api/outlook/office.entities#taskSuggestions) объекта `_MyEntities` (если было извлечено хотя бы одно предложение). Каждое извлеченное предложение сохраняется в качестве объекта [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) в массиве. Обработчик событий `myGetTaskSuggestions` получает дополнительные данные о каждом предложении задачи.


- Строка, изначально определенная как предложение задачи из свойства [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskString).

- Массив уполномоченных из свойства [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees). Каждый уполномоченный представлен объектом [EmailUser](/javascript/api/outlook/office.emailuser).

- Имя из свойства [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayName) для каждого уполномоченного.

- SMTP-адрес из свойства [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailAddress) для каждого уполномоченного.

Чтобы отобразить данные каждого предложения задачи, обработчик событий `myGetTaskSuggestions` формирует локальную HTML-строку в `htmlText`. Ниже представлен соответствующий код JavaScript.




```js
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
               "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="extracting-urls"></a>Извлечение URL-адресов


Когда пользователь нажимает кнопку **Get URLs** (Получить URL-адреса), обработчик событий `myGetUrls` получает массив URL-адресов из свойства [urls](/javascript/api/outlook/office.entities#urls) объекта `_MyEntities` (если был извлечен хотя бы один URL-адрес). Каждый извлеченный адрес сохраняется в массиве в виде строки. Для отображения списка извлеченных URL-адресов обработчик событий `myGetUrls` формирует локальную HTML-строку в `htmlText`.


```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="clearing-displayed-entity-strings"></a>Очистка отображаемых строк сущностей


В заключение, надстройка для работы с сущностями указывает обработчик событий `myClearEntitiesBox`, который очищает отображаемые строки. Ниже приведен соответствующий код.


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## <a name="javascript-listing"></a>Листинг JavaScript


Ниже приведен полный листинг реализации JavaScript.


```js
// Global variables
var _Item;
var _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}


// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}

// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++)
    {
        htmlText += "E-mail Address : <span>" + 
            emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions()
{
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++)
        {
            htmlText += "Attendee : ( ";
            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + 
            meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + 
            meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + 
           meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + 
            meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
            "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
           phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="see-also"></a>См. также

- [Создание надстроек Outlook для форм чтения](read-scenario.md)
- [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md)
- [Метод item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
