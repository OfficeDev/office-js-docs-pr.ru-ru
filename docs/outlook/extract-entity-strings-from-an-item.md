---
title: Извлечение строк сущности из элемента Outlook
description: Узнайте, как извлечь строки сущностей из элемента Outlook в надстройке Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: b15ad23427f79a333ae8ae9d342acdf28e6d010c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608945"
---
# <a name="extract-entity-strings-from-an-outlook-item"></a><span data-ttu-id="67390-103">Извлечение строк сущностей из элемента Outlook</span><span class="sxs-lookup"><span data-stu-id="67390-103">Extract entity strings from an Outlook item</span></span>

<span data-ttu-id="67390-p101">В этой статье рассказано, как создать надстройку Outlook **для отображения сущностей**, которая извлекает экземпляры строк поддерживаемых известных сущностей в теме и основном тексте выбранного элемента Outlook. Этим элементом может быть встреча, электронное сообщение, приглашение на собрание, ответ на такое приглашение или отказ от него.</span><span class="sxs-lookup"><span data-stu-id="67390-p101">This article describes how to create a **Display entities** Outlook add-in that extracts string instances of supported well-known entities in the subject and body of the selected Outlook item. This item can be an appointment, email message, or meeting request, response, or cancellation.</span></span>

<span data-ttu-id="67390-106">Поддерживаемые сущности:</span><span class="sxs-lookup"><span data-stu-id="67390-106">The supported entities include:</span></span>

- <span data-ttu-id="67390-107">**Address**. Почтовый адрес США, который содержит по крайней мере подмножество элементов, включающее номер дома, название улицы, город, штат, а также почтовый индекс.</span><span class="sxs-lookup"><span data-stu-id="67390-107">**Address**: A United States postal address, that has at least a subset of the elements of a street number, street name, city, state, and zip code.</span></span>
    
- <span data-ttu-id="67390-108">**Contact**. Контактные данные лица в контексте других сущностей, например адреса или названия организации.</span><span class="sxs-lookup"><span data-stu-id="67390-108">**Contact**: A person's contact information, in the context of other entities such as an address or business name.</span></span>
    
- <span data-ttu-id="67390-109">**Email address**. SMTP-адрес электронной почты.</span><span class="sxs-lookup"><span data-stu-id="67390-109">**Email address**: An SMTP email address.</span></span>
    
- <span data-ttu-id="67390-p102">**Meeting suggestion**. Приглашение на собрание, например ссылка на мероприятие. Обратите внимание на то, что извлечение приглашений поддерживается только для сообщений, но не для встреч.</span><span class="sxs-lookup"><span data-stu-id="67390-p102">**Meeting suggestion**: A meeting suggestion, such as a reference to an event. Note that only messages but not appointments support extracting meeting suggestions.</span></span>
    
- <span data-ttu-id="67390-112">**Phone number**. Телефонный номер Северной Америки.</span><span class="sxs-lookup"><span data-stu-id="67390-112">**Phone number**: A North American phone number.</span></span>
    
- <span data-ttu-id="67390-113">**Task suggestion**. Предложение задачи, которое обычно выражается фразой с действиями.</span><span class="sxs-lookup"><span data-stu-id="67390-113">**Task suggestion**: A task suggestion, typically expressed in an actionable phrase.</span></span>
    
- <span data-ttu-id="67390-114">**URL**.</span><span class="sxs-lookup"><span data-stu-id="67390-114">**URL**</span></span>
    
<span data-ttu-id="67390-p103">Большинство из этих сущностей зависят от распознавания естественного языка, которое основывается на обработке компьютером больших объемов данных. Это распознавание недетерминированное и иногда зависит от контекста в элементе Outlook.</span><span class="sxs-lookup"><span data-stu-id="67390-p103">Most of these entities rely on natural language recognition, which is based on machine learning of large amounts of data. This recognition is nondeterministic and sometimes depends on the context in the Outlook item.</span></span>

<span data-ttu-id="67390-p104">Outlook активирует надстройку для работы с сущностями каждый раз, когда пользователь выбирает встречу, электронное письмо, приглашение на собрание, ответ на приглашение на собрание или отказ от приглашения на собрание для просмотра. Во время инициализации в примере надстройки для работы с сущностями выполняется считывание всех экземпляров поддерживаемых сущностей из текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="67390-p104">Outlook activates the entities add-in whenever the user selects an appointment, email message, or meeting request, response, or cancellation for viewing. During initialization, the sample entities add-in reads all instances of the supported entities from the current item.</span></span> 

<span data-ttu-id="67390-p105">Надстройка предоставляет кнопки, с помощью которых пользователь может выбрать тип сущности. Когда пользователь выбирает какую-либо сущность, надстройка отображает экземпляры выбранной сущности в области надстройки. В последующих разделах имеются манифест в формате XML, HTML- и JavaScript-файлы надстроек сущностей, а также выделен код, поддерживающий извлечение соответствующих сущностей.</span><span class="sxs-lookup"><span data-stu-id="67390-p105">The add-in provides buttons for the user to choose a type of entity. When the user selects an entity, the add-in displays instances of the selected entity in the add-in pane. The following sections list the XML manifest, and HTML and JavaScript files of the entities add-in, and highlight the code that supports the respective entity extraction.</span></span>

## <a name="xml-manifest"></a><span data-ttu-id="67390-122">XML-манифест</span><span class="sxs-lookup"><span data-stu-id="67390-122">XML manifest</span></span>

<span data-ttu-id="67390-123">Надстройка для работы с сущностями использует два правила активации, объединенных логической операцией ИЛИ.</span><span class="sxs-lookup"><span data-stu-id="67390-123">The entities add-in has two activation rules joined by a logical OR operation.</span></span> 

```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

<span data-ttu-id="67390-124">Эти правила определяют, что Outlook должен активировать надстройку, если в области чтения или инспекторе просмотра выбрана встреча или сообщение (включая письмо или приглашение на собрание, ответ на приглашение или отмену собрания).</span><span class="sxs-lookup"><span data-stu-id="67390-124">These rules specify that Outlook should activate this add-in when the currently selected item in the Reading Pane or read inspector is an appointment or message (including an email message, or meeting request, response, or cancellation).</span></span>

<span data-ttu-id="67390-p106">Ниже приведен манифест надстройки для работы с сущностями. В нем используется схема версии 1.1 для манифестов надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="67390-p106">The following is the manifest of the entities add-in. It uses version 1.1 of the schema for Office Add-ins manifests.</span></span>

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


## <a name="html-implementation"></a><span data-ttu-id="67390-127">Реализация HTML</span><span class="sxs-lookup"><span data-stu-id="67390-127">HTML implementation</span></span>

<span data-ttu-id="67390-p107">HTML-файл надстройки для работы с сущностями определяет кнопки, позволяющие пользователю выбрать каждый тип сущности, и одну кнопку для очистки отображаемых экземпляров сущности. В нем есть JavaScript-файл, default_entities.js, который описан в следующем разделе [Реализация JavaScript](#javascript-implementation). JavaScript-файл содержит обработчики событий для каждой кнопки.</span><span class="sxs-lookup"><span data-stu-id="67390-p107">The HTML file of the entities add-in specifies buttons for the user to select each type of entity, and another button to clear displayed instances of an entity. It includes a JavaScript file, default_entities.js, which is described in the next section under [JavaScript implementation](#javascript-implementation). The JavaScript file includes the event handlers for each of the buttons.</span></span>

<span data-ttu-id="67390-p108">Обратите внимание, что все надстройки Outlook должны включать файл office.js. Приведенный ниже HTML-файл включает файл office.js версии 1.1 в CDN.</span><span class="sxs-lookup"><span data-stu-id="67390-p108">Note that all Outlook add-ins must include office.js. The HTML file that follows includes version 1.1 of office.js on the CDN.</span></span> 

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


## <a name="style-sheet"></a><span data-ttu-id="67390-133">Таблица стилей</span><span class="sxs-lookup"><span data-stu-id="67390-133">Style sheet</span></span>


<span data-ttu-id="67390-p109">В надстройке для работы с сущностями используется дополнительный файл таблицы стилей default_entities.css, который определяет макет выходных данных. Ниже приведен листинг CSS-файла.</span><span class="sxs-lookup"><span data-stu-id="67390-p109">The entities add-in uses an optional CSS file, default_entities.css, to specify the layout of the output. The following is a listing of the CSS file.</span></span>


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


## <a name="javascript-implementation"></a><span data-ttu-id="67390-136">Реализация JavaScript</span><span class="sxs-lookup"><span data-stu-id="67390-136">JavaScript implementation</span></span>

<span data-ttu-id="67390-137">В следующих разделах описано, как этот пример (файл default_entities.js) извлекает известные сущности из темы и текста сообщения или встречи, которую просматривает пользователь.</span><span class="sxs-lookup"><span data-stu-id="67390-137">The remaining sections describe how this sample (default_entities.js file) extracts well-known entities from the subject and body of the message or appointment that the user is viewing.</span></span>

## <a name="extracting-entities-upon-initialization"></a><span data-ttu-id="67390-138">Извлечение сущностей при инициализации</span><span class="sxs-lookup"><span data-stu-id="67390-138">Extracting entities upon initialization</span></span>

<span data-ttu-id="67390-139">Когда происходит событие [Office.initialize](/javascript/api/office#office-initialize-reason-), надстройка для работы с сущностями вызывает метод [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="67390-139">Upon the [Office.initialize](/javascript/api/office#office-initialize-reason-) event, the entities add-in calls the [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method of the current item.</span></span> <span data-ttu-id="67390-140">`getEntities`Метод возвращает глобальную переменную в `_MyEntities` виде массива экземпляров поддерживаемых сущностей.</span><span class="sxs-lookup"><span data-stu-id="67390-140">The `getEntities` method returns the global variable `_MyEntities` an array of instances of supported entities.</span></span> <span data-ttu-id="67390-141">Ниже представлен соответствующий код JavaScript.</span><span class="sxs-lookup"><span data-stu-id="67390-141">The following is the related JavaScript code.</span></span>


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


## <a name="extracting-addresses"></a><span data-ttu-id="67390-142">Извлечение адресов</span><span class="sxs-lookup"><span data-stu-id="67390-142">Extracting addresses</span></span>


<span data-ttu-id="67390-143">Когда пользователь нажимает кнопку **Get Addresses** (Получить адреса), обработчик событий `myGetAddresses` получает массив адресов из свойства [addresses](/javascript/api/outlook/office.entities#addresses) объекта `_MyEntities` (если был извлечен хотя бы один адрес).</span><span class="sxs-lookup"><span data-stu-id="67390-143">When the user clicks the **Get Addresses** button, the `myGetAddresses` event handler obtains an array of addresses from the [addresses](/javascript/api/outlook/office.entities#addresses) property of the `_MyEntities` object, if any address was extracted.</span></span> <span data-ttu-id="67390-144">Каждый извлеченный адрес хранится в массиве в виде строки.</span><span class="sxs-lookup"><span data-stu-id="67390-144">Each extracted address is stored as a string in the array.</span></span> <span data-ttu-id="67390-145">Чтобы отобразить список извлеченных URL-адресов, обработчик событий `myGetAddresses` формирует локальную HTML-строку в `htmlText`.</span><span class="sxs-lookup"><span data-stu-id="67390-145">`myGetAddresses` forms a local HTML string in `htmlText` to display the list of extracted addresses.</span></span> <span data-ttu-id="67390-146">Ниже представлен соответствующий код JavaScript.</span><span class="sxs-lookup"><span data-stu-id="67390-146">The following is the related JavaScript code.</span></span>


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


## <a name="extracting-contact-information"></a><span data-ttu-id="67390-147">Извлечение контактных данных</span><span class="sxs-lookup"><span data-stu-id="67390-147">Extracting contact information</span></span>


<span data-ttu-id="67390-p112">Когда пользователь нажимает кнопку **Get Contact Information** (Получить контактные данные), обработчик событий `myGetContacts` получает массив контактов вместе с соответствующими сведениями из свойства [contacts](/javascript/api/outlook/office.entities#contacts) объекта `_MyEntities` (если был извлечен хотя бы один контакт). Каждый извлеченный контакт сохраняется в качестве объекта [Contact](/javascript/api/outlook/office.contact) в массиве. Обработчик событий `myGetContacts` получает дополнительные данные о каждом контакте. Обратите внимание на то, что контекст определяет, может ли Outlook извлекать контакт из элемента (подпись в конце электронного сообщения) или же в непосредственной близости от контакта должны присутствовать какие-либо из указанных ниже данных.</span><span class="sxs-lookup"><span data-stu-id="67390-p112">When the user clicks the **Get Contact Information** button, the `myGetContacts` event handler obtains an array of contacts together with their information from the [contacts](/javascript/api/outlook/office.entities#contacts) property of the `_MyEntities` object, if any was extracted. Each extracted contact is stored as a [Contact](/javascript/api/outlook/office.contact) object in the array. `myGetContacts` obtains further data about each contact. Note that the context determines whether Outlook can extract a contact from an item&mdash;a signature at the end of an email message, or at least some of the following information would have to exist in the vicinity of the contact:</span></span>


- <span data-ttu-id="67390-152">Строка, представляющая имя контакта из свойства [Contact.personName](/javascript/api/outlook/office.contact#personname).</span><span class="sxs-lookup"><span data-stu-id="67390-152">The string representing the contact's name from the [Contact.personName](/javascript/api/outlook/office.contact#personname) property.</span></span>

- <span data-ttu-id="67390-153">Название компании, связанное с контактом, из свойства [Contact.businessName](/javascript/api/outlook/office.contact#businessname).</span><span class="sxs-lookup"><span data-stu-id="67390-153">The string representing the company name associated with the contact from the [Contact.businessName](/javascript/api/outlook/office.contact#businessname) property.</span></span>

- <span data-ttu-id="67390-p113">Массив номеров телефонов, связанных с контактом, из свойства [Contact.phoneNumbers](/javascript/api/outlook/office.contact#phonenumbers). Каждый номер телефона представлен объектом [PhoneNumber](/javascript/api/outlook/office.phonenumber).</span><span class="sxs-lookup"><span data-stu-id="67390-p113">The array of telephone numbers associated with the contact from the [Contact.phoneNumbers](/javascript/api/outlook/office.contact#phonenumbers) property. Each telephone number is represented by a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object.</span></span>

- <span data-ttu-id="67390-156">Строка, представляющая телефонный номер из свойства [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) для каждого элемента **PhoneNumber** в массиве телефонных номеров.</span><span class="sxs-lookup"><span data-stu-id="67390-156">For each **PhoneNumber** member in the telephone numbers array, the string representing the telephone number from the [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) property.</span></span>

- <span data-ttu-id="67390-p114">Массив URL-адресов, связанных с контактом, из свойства [Contact.urls](/javascript/api/outlook/office.contact#urls). Каждый URL-адрес представлен в виде строки в элементе массива.</span><span class="sxs-lookup"><span data-stu-id="67390-p114">The array of URLs associated with the contact from the [Contact.urls](/javascript/api/outlook/office.contact#urls) property. Each URL is represented as a string in an array member.</span></span>

- <span data-ttu-id="67390-p115">Массив адресов эл. почты, связанных с контактом, из свойства [Contact.emailAddresses](/javascript/api/outlook/office.contact#emailaddresses). Каждый адрес эл. почты представлен в виде строки в элементе массива.</span><span class="sxs-lookup"><span data-stu-id="67390-p115">The array of email addresses associated with the contact from the [Contact.emailAddresses](/javascript/api/outlook/office.contact#emailaddresses) property. Each email address is represented as a string in an array member.</span></span>

- <span data-ttu-id="67390-p116">Массив почтовых адресов, связанных с контактом, из свойства [Contact.addresses](/javascript/api/outlook/office.contact#addresses). Каждый почтовый адрес представлен в виде строки в элементе массива.</span><span class="sxs-lookup"><span data-stu-id="67390-p116">The array of postal addresses associated with the contact from the [Contact.addresses](/javascript/api/outlook/office.contact#addresses) property. Each postal address is represented as a string in an array member.</span></span>

<span data-ttu-id="67390-p117">Чтобы отобразить данные каждого контакта, обработчик событий `myGetContacts` формирует локальную HTML-строку в `htmlText`. Ниже представлен соответствующий код JavaScript.</span><span class="sxs-lookup"><span data-stu-id="67390-p117">`myGetContacts` forms a local HTML string in `htmlText` to display the data for each contact. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-email-addresses"></a><span data-ttu-id="67390-165">Извлечение электронных адресов</span><span class="sxs-lookup"><span data-stu-id="67390-165">Extracting email addresses</span></span>


<span data-ttu-id="67390-p118">Когда пользователь нажимает кнопку **Get Email Addresses** (Получить электронные адреса), обработчик события `myGetEmailAddresses` получает массив SMTP-адресов электронной почты из свойства [emailAddresses](/javascript/api/outlook/office.entities#emailaddresses) объекта `_MyEntities` (если был извлечен хотя бы один адрес). Каждый извлеченный электронный адрес сохраняется в массиве в виде строки. Для отображения списка извлеченных электронных адресов обработчик событий `myGetEmailAddresses` формирует локальную HTML-строку в `htmlText`. Ниже приведен соответствующий код JavaScript.</span><span class="sxs-lookup"><span data-stu-id="67390-p118">When the user clicks the **Get Email Addresses** button, the `myGetEmailAddresses` event handler obtains an array of SMTP email addresses from the [emailAddresses](/javascript/api/outlook/office.entities#emailaddresses) property of the `_MyEntities` object, if any was extracted. Each extracted email address is stored as a string in the array. `myGetEmailAddresses` forms a local HTML string in `htmlText` to display the list of extracted email addresses. The following is the related JavaScript code.</span></span>


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


## <a name="extracting-meeting-suggestions"></a><span data-ttu-id="67390-170">Извлечение приглашений на собрания</span><span class="sxs-lookup"><span data-stu-id="67390-170">Extracting meeting suggestions</span></span>


<span data-ttu-id="67390-171">Когда пользователь нажимает кнопку **Get Meeting Suggestions** (Получить приглашения на собрания), обработчик событий `myGetMeetingSuggestions` получает массив приглашений на собрания из свойства [meetingSuggestions](/javascript/api/outlook/office.entities#meetingsuggestions) объекта `_MyEntities` (если было извлечено хотя бы одно приглашение).</span><span class="sxs-lookup"><span data-stu-id="67390-171">When the user clicks the **Get Meeting Suggestions** button, the `myGetMeetingSuggestions` event handler obtains an array of meeting suggestions from the [meetingSuggestions](/javascript/api/outlook/office.entities#meetingsuggestions) property of the `_MyEntities` object, if any was extracted.</span></span>


 > [!NOTE]
 > <span data-ttu-id="67390-172">Тип объекта поддерживается только сообщениями, но не встречами `MeetingSuggestion` .</span><span class="sxs-lookup"><span data-stu-id="67390-172">Only messages but not appointments support the `MeetingSuggestion` entity type.</span></span>

<span data-ttu-id="67390-p119">Каждое извлеченное приглашение на собрание хранится в виде объекта [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) в массиве. Обработчик событий `myGetMeetingSuggestions` получает дополнительные данные о каждом приглашении на собрание:</span><span class="sxs-lookup"><span data-stu-id="67390-p119">Each extracted meeting suggestion is stored as a [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) object in the array. `myGetMeetingSuggestions` obtains further data about each meeting suggestion:</span></span>


- <span data-ttu-id="67390-175">Приглашение на собрание из свойства [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingstring).</span><span class="sxs-lookup"><span data-stu-id="67390-175">The string that was identified as a meeting suggestion from the [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingstring) property.</span></span>

- <span data-ttu-id="67390-p120">Массив участников собрания из свойства [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees). Каждый участник представлен объектом [EmailUser](/javascript/api/outlook/office.emailuser).</span><span class="sxs-lookup"><span data-stu-id="67390-p120">The array of meeting attendees from the [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees) property. Each attendee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.</span></span>

- <span data-ttu-id="67390-178">Имя из свойства [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) для каждого участника.</span><span class="sxs-lookup"><span data-stu-id="67390-178">For each attendee, the name from the [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) property.</span></span>

- <span data-ttu-id="67390-179">SMTP-адрес из свойства [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) для каждого участника.</span><span class="sxs-lookup"><span data-stu-id="67390-179">For each attendee, the SMTP address from the [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) property.</span></span>

- <span data-ttu-id="67390-180">Предлагаемое место проведения собрания из свойства [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location).</span><span class="sxs-lookup"><span data-stu-id="67390-180">The string representing the location of the meeting suggestion from the [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location) property.</span></span>

- <span data-ttu-id="67390-181">Предлагаемая тема собрания из свойства [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject).</span><span class="sxs-lookup"><span data-stu-id="67390-181">The string representing the subject of the meeting suggestion from the [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject) property.</span></span>

- <span data-ttu-id="67390-182">Предлагаемое время начала собрания из свойства [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start).</span><span class="sxs-lookup"><span data-stu-id="67390-182">The string representing the start time of the meeting suggestion from the [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start) property.</span></span>

- <span data-ttu-id="67390-183">Предлагаемое время окончания собрания из свойства [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end).</span><span class="sxs-lookup"><span data-stu-id="67390-183">The string representing the end time of the meeting suggestion from the [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end) property.</span></span>

<span data-ttu-id="67390-p121">Чтобы отобразить данные каждого приглашения на собрание, обработчик событий `myGetMeetingSuggestions` формирует локальную HTML-строку в `htmlText`. Ниже представлен соответствующий код JavaScript.</span><span class="sxs-lookup"><span data-stu-id="67390-p121">`myGetMeetingSuggestions` forms a local HTML string in `htmlText` to display the data for each of the meeting suggestions. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-phone-numbers"></a><span data-ttu-id="67390-186">Извлечение телефонных номеров</span><span class="sxs-lookup"><span data-stu-id="67390-186">Extracting phone numbers</span></span>


<span data-ttu-id="67390-p122">Когда пользователь нажимает кнопку **Get Phone Numbers** (Получить телефонные номера), обработчик событий `myGetPhoneNumbers` получает массив телефонных номеров из свойства [phoneNumbers](/javascript/api/outlook/office.entities#phonenumbers) объекта `_MyEntities` (если был извлечен хотя бы один номер). Каждый извлеченный номер сохраняется в качестве объекта [PhoneNumber](/javascript/api/outlook/office.phonenumber) в массиве. Обработчик событий `myGetPhoneNumbers` получает дополнительные данные о каждом телефонном номере.</span><span class="sxs-lookup"><span data-stu-id="67390-p122">When the user clicks the **Get Phone Numbers** button, the `myGetPhoneNumbers` event handler obtains an array of phone numbers from the [phoneNumbers](/javascript/api/outlook/office.entities#phonenumbers) property of the `_MyEntities` object, if any was extracted. Each extracted phone number is stored as a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object in the array. `myGetPhoneNumbers` obtains further data about each phone number:</span></span>


- <span data-ttu-id="67390-190">Строка, представляющая тип номера телефона (например, домашний номер) из свойства [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type).</span><span class="sxs-lookup"><span data-stu-id="67390-190">The string representing the kind of phone number, for example, home phone number, from the [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type) property.</span></span>

- <span data-ttu-id="67390-191">Номер телефона из свойства [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring).</span><span class="sxs-lookup"><span data-stu-id="67390-191">The string representing the actual phone number from the [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) property.</span></span>

- <span data-ttu-id="67390-192">Исходный номер телефона из свойства [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalphonestring).</span><span class="sxs-lookup"><span data-stu-id="67390-192">The string that was originally identified as the phone number from the [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalphonestring) property.</span></span>

<span data-ttu-id="67390-p123">Чтобы отобразить данные каждого номера телефона, обработчик событий `myGetPhoneNumbers` формирует локальную HTML-строку в `htmlText`. Ниже представлен соответствующий код JavaScript.</span><span class="sxs-lookup"><span data-stu-id="67390-p123">`myGetPhoneNumbers` forms a local HTML string in `htmlText` to display the data for each of the phone numbers. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-task-suggestions"></a><span data-ttu-id="67390-195">Извлечение предложений задач</span><span class="sxs-lookup"><span data-stu-id="67390-195">Extracting task suggestions</span></span>


<span data-ttu-id="67390-p124">Когда пользователь нажимает кнопку **Get Task Suggestions** (Получить предложения задач), обработчик событий `myGetTaskSuggestions` получает массив предложений задач из свойства [taskSuggestions](/javascript/api/outlook/office.entities#tasksuggestions) объекта `_MyEntities` (если было извлечено хотя бы одно предложение). Каждое извлеченное предложение сохраняется в качестве объекта [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) в массиве. Обработчик событий `myGetTaskSuggestions` получает дополнительные данные о каждом предложении задачи.</span><span class="sxs-lookup"><span data-stu-id="67390-p124">When the user clicks the **Get Task Suggestions** button, the `myGetTaskSuggestions` event handler obtains an array of task suggestions from the [taskSuggestions](/javascript/api/outlook/office.entities#tasksuggestions) property of the `_MyEntities` object, if any was extracted. Each extracted task suggestion is stored as a [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) object in the array. `myGetTaskSuggestions` obtains further data about each task suggestion:</span></span>


- <span data-ttu-id="67390-199">Строка, изначально определенная как предложение задачи из свойства [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskstring).</span><span class="sxs-lookup"><span data-stu-id="67390-199">The string that was originally identified a task suggestion from the [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskstring) property.</span></span>

- <span data-ttu-id="67390-p125">Массив уполномоченных из свойства [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees). Каждый уполномоченный представлен объектом [EmailUser](/javascript/api/outlook/office.emailuser).</span><span class="sxs-lookup"><span data-stu-id="67390-p125">The array of task assignees from the [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees) property. Each assignee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.</span></span>

- <span data-ttu-id="67390-202">Имя из свойства [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) для каждого уполномоченного.</span><span class="sxs-lookup"><span data-stu-id="67390-202">For each assignee, the name from the [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) property.</span></span>

- <span data-ttu-id="67390-203">SMTP-адрес из свойства [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) для каждого уполномоченного.</span><span class="sxs-lookup"><span data-stu-id="67390-203">For each assignee, the SMTP address from the [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) property.</span></span>

<span data-ttu-id="67390-p126">Чтобы отобразить данные каждого предложения задачи, обработчик событий `myGetTaskSuggestions` формирует локальную HTML-строку в `htmlText`. Ниже представлен соответствующий код JavaScript.</span><span class="sxs-lookup"><span data-stu-id="67390-p126">`myGetTaskSuggestions` forms a local HTML string in `htmlText` to display the data for each task suggestion. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-urls"></a><span data-ttu-id="67390-206">Извлечение URL-адресов</span><span class="sxs-lookup"><span data-stu-id="67390-206">Extracting URLs</span></span>


<span data-ttu-id="67390-p127">Когда пользователь нажимает кнопку **Get URLs** (Получить URL-адреса), обработчик событий `myGetUrls` получает массив URL-адресов из свойства [urls](/javascript/api/outlook/office.entities#urls) объекта `_MyEntities` (если был извлечен хотя бы один URL-адрес). Каждый извлеченный адрес сохраняется в массиве в виде строки. Для отображения списка извлеченных URL-адресов обработчик событий `myGetUrls` формирует локальную HTML-строку в `htmlText`.</span><span class="sxs-lookup"><span data-stu-id="67390-p127">When the user clicks the **Get URLs** button, the `myGetUrls` event handler obtains an array of URLs from the [urls](/javascript/api/outlook/office.entities#urls) property of the `_MyEntities` object, if any was extracted. Each extracted URL is stored as a string in the array. `myGetUrls` forms a local HTML string in `htmlText` to display the list of extracted URLs.</span></span>


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


## <a name="clearing-displayed-entity-strings"></a><span data-ttu-id="67390-210">Очистка отображаемых строк сущностей</span><span class="sxs-lookup"><span data-stu-id="67390-210">Clearing displayed entity strings</span></span>


<span data-ttu-id="67390-p128">В заключение, надстройка для работы с сущностями указывает обработчик событий `myClearEntitiesBox`, который очищает отображаемые строки. Ниже приведен соответствующий код.</span><span class="sxs-lookup"><span data-stu-id="67390-p128">Lastly, the entities add-in specifies a  `myClearEntitiesBox` event handler which clears any displayed strings. The following is the related code.</span></span>


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## <a name="javascript-listing"></a><span data-ttu-id="67390-213">Листинг JavaScript</span><span class="sxs-lookup"><span data-stu-id="67390-213">JavaScript listing</span></span>


<span data-ttu-id="67390-214">Ниже приведен полный листинг реализации JavaScript.</span><span class="sxs-lookup"><span data-stu-id="67390-214">The following is the complete listing of the JavaScript implementation.</span></span>


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


## <a name="see-also"></a><span data-ttu-id="67390-215">См. также</span><span class="sxs-lookup"><span data-stu-id="67390-215">See also</span></span>

- [<span data-ttu-id="67390-216">Создание надстроек Outlook для форм чтения</span><span class="sxs-lookup"><span data-stu-id="67390-216">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="67390-217">Сопоставление строк в элементе Outlook как известных сущностей</span><span class="sxs-lookup"><span data-stu-id="67390-217">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
- [<span data-ttu-id="67390-218">Метод item.getEntities</span><span class="sxs-lookup"><span data-stu-id="67390-218">item.getEntities method</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
