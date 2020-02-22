---
title: Использование веб-служб Exchange (EWS) из надстройки Outlook
description: Содержит пример, в котором показано, как надстройка Outlook может запрашивать сведения из веб-службы Exchange.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4c0c97a9a796dc1f257b1a0b0ec880b3ca3d8e74
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166631"
---
# <a name="call-web-services-from-an-outlook-add-in"></a><span data-ttu-id="699fe-103">Вызов веб-служб из надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="699fe-103">Call web services from an Outlook add-in</span></span>

<span data-ttu-id="699fe-p101">Ваша надстройка может использовать веб-службы Exchange (EWS) на компьютере с Exchange Server 2013; веб-службу, доступную на сервере, предоставляющем исходное расположение для пользовательского интерфейса надстройки; или веб-службу, доступную через Интернет. В этой статье приведен пример того, как надстройка Outlook может запрашивать данные из EWS.</span><span class="sxs-lookup"><span data-stu-id="699fe-p101">Your add-in can use Exchange Web Services (EWS) from a computer that is running Exchange Server 2013, a web service that is available on the server that provides the source location for the add-in's UI, or a web service that is available on the Internet. This article provides an example that shows how an Outlook add-in can request information from EWS.</span></span>

<span data-ttu-id="699fe-p102">Способы вызова веб-службы различаются в зависимости от расположения службы. В таблице 1 приведены различные способы вызова веб-службы в зависимости от расположения.</span><span class="sxs-lookup"><span data-stu-id="699fe-p102">The way that you call a web service varies based on where the web service is located. Table 1 lists the different ways that you can call a web service based on location.</span></span>


<span data-ttu-id="699fe-108">**Таблица 1. Способы вызова веб-служб из надстройки Outlook**</span><span class="sxs-lookup"><span data-stu-id="699fe-108">**Table 1. Ways to call web services from an Outlook add-in**</span></span>

<br/>

|<span data-ttu-id="699fe-109">**Расположение веб-службы**</span><span class="sxs-lookup"><span data-stu-id="699fe-109">**Web service location**</span></span>|<span data-ttu-id="699fe-110">**Способ вызова веб-службы**</span><span class="sxs-lookup"><span data-stu-id="699fe-110">**Way to call the web service**</span></span>|
|:-----|:-----|
|<span data-ttu-id="699fe-111">Сервер Exchange, на котором размещен почтовый ящик клиента</span><span class="sxs-lookup"><span data-stu-id="699fe-111">The Exchange server that hosts the client mailbox</span></span>|<span data-ttu-id="699fe-p103">Используйте метод [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) для вызова операций EWS, поддерживаемых надстройками. Сервер Exchange Server, на котором размещен почтовый ящик, также предоставляет доступ к EWS.</span><span class="sxs-lookup"><span data-stu-id="699fe-p103">Use the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to call EWS operations that add-ins support. The Exchange server that hosts the mailbox also exposes EWS.</span></span>|
|<span data-ttu-id="699fe-114">Веб-сервер, предоставляющий исходное расположение для пользовательского интерфейса надстроек.</span><span class="sxs-lookup"><span data-stu-id="699fe-114">The web server that provides the source location for the add-in UI</span></span>|<span data-ttu-id="699fe-p104">Вызывайте веб-службу с помощью стандартных методик JavaScript. Код JavaScript в пределах пользовательского интерфейса работает в контексте веб-сервера, предоставляющего пользовательский интерфейс. Поэтому он сможет вызывать веб-службы на этом сервере, не создавая ошибки межсайтового скрипта.</span><span class="sxs-lookup"><span data-stu-id="699fe-p104">Call the web service by using standard JavaScript techniques. The JavaScript code in the UI frame runs in the context of the web server that provides the UI. Therefore, it can call web services on that server without causing a cross-site scripting error.</span></span>|
|<span data-ttu-id="699fe-118">Все другие расположения</span><span class="sxs-lookup"><span data-stu-id="699fe-118">All other locations</span></span>|<span data-ttu-id="699fe-p105">Создайте прокси для веб-службы на веб-сервере, предоставляющем исходное расположение для пользовательского интерфейса. Если не указать прокси, надстройка не запустится из-за ошибок межсайтовых сценариев. Один из способов указать такой прокси — это использовать JSON/P. Дополнительные сведения см. в статье [Конфиденциальность и безопасность надстроек для Office](../develop/privacy-and-security.md).</span><span class="sxs-lookup"><span data-stu-id="699fe-p105">Create a proxy for the web service on the web server that provides the source location for the UI. If you do not provide a proxy, cross-site scripting errors will prevent your add-in from running. One way to provide a proxy is by using JSON/P. For more information, see [Privacy and security for Office Add-ins](../develop/privacy-and-security.md).</span></span>|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a><span data-ttu-id="699fe-123">Получение доступа к операциям веб-служб Exchange с помощью метода makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="699fe-123">Using the makeEwsRequestAsync method to access EWS operations</span></span>

<span data-ttu-id="699fe-124">С помощью метода [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) вы можете отправить запрос EWS на сервер Exchange Server, на котором размещается почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="699fe-124">You can use the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to make an EWS request to the Exchange server that hosts the user's mailbox.</span></span>

<span data-ttu-id="699fe-p106">Веб-службы Exchange поддерживают различные операции на сервере Exchange. Например, операции копирования, поиска, обновления или отправки на уровне элемента, а также операции создания, получения или обновления на уровне папки. Чтобы выполнить операцию веб-служб Exchange, создайте для нее SOAP-запрос в формате XML. После завершения операции будет возвращен SOAP-ответ в формате XML с необходимыми данными. SOAP-запросы к веб-службам Exchange и их SOAP-ответы соответствуют схеме, определенной в файле Messages.xsd. Как и другие файлы схемы веб-служб Exchange, файл Message.xsd расположен в виртуальном каталоге IIS, в котором размещены веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-p106">EWS supports different operations on an Exchange server; for example, item-level operations to copy, find, update, or send an item, and folder-level operations to create, get, or update a folder. To perform an EWS operation, create an XML SOAP request for that operation. When the operation finishes, you get an XML SOAP response that contains data that is relevant to the operation. EWS SOAP requests and responses follow the schema defined in the Messages.xsd file. Like other EWS schema files, the Message.xsd file is located in the IIS virtual directory that hosts EWS.</span></span>

<span data-ttu-id="699fe-130">Чтобы использовать метод **makeEwsRequestAsync** для запуска операции веб-служб Exchange, предоставьте следующее:</span><span class="sxs-lookup"><span data-stu-id="699fe-130">To use the **makeEwsRequestAsync** method to initiate an EWS operation, provide the following:</span></span>

- <span data-ttu-id="699fe-131">XML-код SOAP-запроса для соответствующей операции EWS в качестве аргумента для параметра  _data_;</span><span class="sxs-lookup"><span data-stu-id="699fe-131">The XML for the SOAP request for that EWS operation, as an argument to the  _data_ parameter</span></span>

- <span data-ttu-id="699fe-132">метод обратного вызова (в качестве аргумента  _callback_);</span><span class="sxs-lookup"><span data-stu-id="699fe-132">A callback method (as the  _callback_ argument)</span></span>

- <span data-ttu-id="699fe-133">все необязательные входные данные для этого метода обратного вызова (в качестве аргумента  _userContext_).</span><span class="sxs-lookup"><span data-stu-id="699fe-133">Any optional input data for that callback method (as the  _userContext_ argument)</span></span>

<span data-ttu-id="699fe-p107">Когда SOAP-запрос к веб-службам Exchange выполнен, Outlook вызывает метод обратного вызова с аргументом в виде объекта [AsyncResult](/javascript/api/office/office.asyncresult). Такой метод позволяет получить доступ к двум свойствам объекта  **AsyncResult**. Вот они: свойство  **value**, содержащее SOAP-ответ в формате XML (получен при выполнении операции веб-служб Exchange), и свойство  **asyncContext** (необязательное), содержащее все данные, переданные в виде параметра **userContext**. Как правило, затем метод обратного вызова анализирует XML-код в SOAP-ответе, чтобы получить необходимые сведения и обработать их соответствующим образом.</span><span class="sxs-lookup"><span data-stu-id="699fe-p107">When the EWS SOAP request is complete, Outlook calls the callback method with one argument, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object. The callback method can access two properties of the **AsyncResult** object: the **value** property, which contains the XML SOAP response of the EWS operation, and optionally, the **asyncContext** property, which contains any data passed as the **userContext** parameter. Typically, the callback method then parses the XML in the SOAP response to get any relevant information, and processes that information accordingly.</span></span>


## <a name="tips-for-parsing-ews-responses"></a><span data-ttu-id="699fe-137">Советы по анализу ответов веб-служб Exchange</span><span class="sxs-lookup"><span data-stu-id="699fe-137">Tips for parsing EWS responses</span></span>

<span data-ttu-id="699fe-138">При анализе SOAP-ответа, полученного при выполнении операции веб-служб Exchange, обратите внимание на приведенные ниже особенности, связанные с типом браузера.</span><span class="sxs-lookup"><span data-stu-id="699fe-138">When parsing a SOAP response from an EWS operation, note the following browser-dependent issues:</span></span>


- <span data-ttu-id="699fe-139">При использовании метода DOM **getElementsByTagName** укажите префикс имени тега, чтобы включить поддержку браузера Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="699fe-139">Specify the prefix for a tag name when using the DOM method **getElementsByTagName**, to include support for Internet Explorer.</span></span>

  <span data-ttu-id="699fe-p108">Метод **getElementsByTagName** работает по-разному в зависимости от типа браузера. Например, ответ EWS может содержать следующий XML-код (отформатированный и сокращенный для наглядности):</span><span class="sxs-lookup"><span data-stu-id="699fe-p108">**getElementsByTagName** behaves differently depending on browser type. For example, an EWS response can contain the following XML (formatted and abbreviated for display purposes):</span></span>

   ```XML
        <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
        PropertyName="MyProperty" 
        PropertyType="String"/>
        <t:Value>{
        ...
        }</t:Value></t:ExtendedProperty>
   ```

   <span data-ttu-id="699fe-142">Приведенный ниже код позволит получить XML-код, заключенный в теги **ExtendedProperty**, в таком браузере, как Chrome.</span><span class="sxs-lookup"><span data-stu-id="699fe-142">Code, as in the following, would work on a browser like Chrome to get the XML enclosed by the **ExtendedProperty** tags:</span></span>

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("ExtendedProperty")
            });
   ```

   <span data-ttu-id="699fe-143">В Internet Explorer необходимо включить префикс `t:` имени тега, как показано ниже:</span><span class="sxs-lookup"><span data-stu-id="699fe-143">On Internet Explorer, you must include the `t:` prefix of the tag name, as shown below:</span></span>

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("t:ExtendedProperty")
            });
   ```

- <span data-ttu-id="699fe-144">Чтобы получить содержимое тега в ответе веб-служб Exchange, используйте свойство DOM **textContent**:</span><span class="sxs-lookup"><span data-stu-id="699fe-144">Use the DOM property **textContent** to get the contents of a tag in an EWS response, as shown below:</span></span>
    
   ```js
      content = $.parseJSON(value.textContent);
   ```

   <span data-ttu-id="699fe-145">Другие свойства, например **innerHTML** могут не работать в Internet Explorer для некоторых тегов в ответе веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-145">Other properties such as **innerHTML** may not work on Internet Explorer for some tags in an EWS response.</span></span>
    

## <a name="example"></a><span data-ttu-id="699fe-146">Пример</span><span class="sxs-lookup"><span data-stu-id="699fe-146">Example</span></span>

<span data-ttu-id="699fe-p109">Следующий пример вызывает  **makeEwsRequestAsync** для использования операции [GetItem](/exchange/client-developer/web-service-reference/getitem-operation), чтобы получить тему элемента. Этот пример содержит три следующие функции:</span><span class="sxs-lookup"><span data-stu-id="699fe-p109">The following example calls **makeEwsRequestAsync** to use the [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to get the subject of an item. This example includes the following three functions:</span></span>

-  <span data-ttu-id="699fe-149">`getSubjectRequest` &ndash; принимает в качестве входных данных идентификатор элемента и возвращает XML-код SOAP-запроса, чтобы вызвать операцию **GetItem** для заданного элемента.</span><span class="sxs-lookup"><span data-stu-id="699fe-149">`getSubjectRequest` &ndash; Takes an item ID as input, and returns the XML for the SOAP request to call **GetItem** for the specified item.</span></span>
    
-  <span data-ttu-id="699fe-150">`sendRequest` &ndash; вызывает функцию `getSubjectRequest`, чтобы получить SOAP-запрос для выбранного элемента. Затем передает этот запрос и метод обратного вызова, `callback`, в **makeEwsRequestAsync**, чтобы получить тему выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="699fe-150">`sendRequest` &ndash; Calls  `getSubjectRequest` to get the SOAP request for the selected item, then passes the SOAP request and the callback method, `callback`, to **makeEwsRequestAsync** to get the subject of the specified item.</span></span>
    
-  <span data-ttu-id="699fe-151">`callback` &ndash; обрабатывает SOAP-ответ, включающий тему и другие сведения об указанном элементе.</span><span class="sxs-lookup"><span data-stu-id="699fe-151">`callback` &ndash; Processes the SOAP response which includes any subject and other information about the specified item.</span></span>
    

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
'<?xml version="1.0" encoding="utf-8"?>' +
'<soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"' +
'               xmlns:xsd="https://www.w3.org/2001/XMLSchema"' +
'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
'  <soap:Header>' +
'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
'  </soap:Header>' +
'  <soap:Body>' +
'    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
'      <ItemShape>' +
'        <t:BaseShape>IdOnly</t:BaseShape>' +
'        <t:AdditionalProperties>' +
'            <t:FieldURI FieldURI="item:Subject"/>' +
'        </t:AdditionalProperties>' +
'      </ItemShape>' +
'      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
'    </GetItem>' +
'  </soap:Body>' +
'</soap:Envelope>';

   return result;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   var mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
}
```


## <a name="ews-operations-that-add-ins-support"></a><span data-ttu-id="699fe-152">Операции веб-служб Exchange, которые надстройки поддерживают</span><span class="sxs-lookup"><span data-stu-id="699fe-152">EWS operations that add-ins support</span></span>

<span data-ttu-id="699fe-p110">Надстройки Outlook могут получать доступ к подмножеству операций EWS с помощью метода **makeEwsRequestAsync**. Если вы не знакомы с операциями EWS и не знаете, как использовать метод **makeEwsRequestAsync** для доступа к операциям, начните с примера SOAP-запроса для настройки аргумента _data_.</span><span class="sxs-lookup"><span data-stu-id="699fe-p110">Outlook add-ins can access a subset of operations that are available in EWS via the **makeEwsRequestAsync** method. If you are unfamiliar with EWS operations and how to use the **makeEwsRequestAsync** method to access an operation, start with a SOAP request example to customize your _data_ argument.</span></span> 

<span data-ttu-id="699fe-155">В следующем примере показано, как применить метод  **makeEwsRequestAsync**:</span><span class="sxs-lookup"><span data-stu-id="699fe-155">The following describes how you can use the **makeEwsRequestAsync** method:</span></span>

1. <span data-ttu-id="699fe-156">В XML-коде замените все идентификаторы элементов и релевантные атрибуты операций EWS на соответствующие значения.</span><span class="sxs-lookup"><span data-stu-id="699fe-156">In the XML, substitute any item IDs and relevant EWS operation attributes with appropriate values.</span></span>
    
2. <span data-ttu-id="699fe-157">Включите SOAP-запрос в качестве аргумента для параметра  _data_ метода **makeEwsRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="699fe-157">Include the SOAP request as an argument for the  _data_ parameter of **makeEwsRequestAsync**.</span></span>
    
3. <span data-ttu-id="699fe-158">Укажите метод обратного вызова и вызовите **makeEwsRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="699fe-158">Specify a callback method and call **makeEwsRequestAsync**.</span></span>
    
4. <span data-ttu-id="699fe-159">В методе обратного вызова проверьте результаты операции в SOAP-ответе.</span><span class="sxs-lookup"><span data-stu-id="699fe-159">In the callback method, verify the results of the operation in the SOAP response.</span></span>
    
5. <span data-ttu-id="699fe-160">Используйте результаты операции EWS в соответствии с вашими потребностями.</span><span class="sxs-lookup"><span data-stu-id="699fe-160">Use the results of the EWS operation according to your needs.</span></span>
    
<span data-ttu-id="699fe-p111">В следующей таблице указаны операции EWS, которые надстройки поддерживают. Чтобы просмотреть примеры SOAP-запросов и SOAP-ответов, выберите ссылку для каждой операции. Дополнительные сведения об операциях EWS см. в статье [Операции EWS в Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="699fe-p111">The following table lists the EWS operations that add-ins support. To see examples of SOAP requests and responses, choose the link for each operation. For more information about EWS operations, see [EWS operations in Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).</span></span>

<span data-ttu-id="699fe-164">**Таблица 2. Поддерживаемые операции EWS**</span><span class="sxs-lookup"><span data-stu-id="699fe-164">**Table 2. Supported EWS operations**</span></span>

<br/>

|<span data-ttu-id="699fe-165">**Операция служб EWS**</span><span class="sxs-lookup"><span data-stu-id="699fe-165">**EWS operation**</span></span>|<span data-ttu-id="699fe-166">**Описание**</span><span class="sxs-lookup"><span data-stu-id="699fe-166">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="699fe-167">Операция CopyItem</span><span class="sxs-lookup"><span data-stu-id="699fe-167">CopyItem operation</span></span>](/exchange/client-developer/web-service-reference/copyitem-operation)|<span data-ttu-id="699fe-168">Копирует выбранные элементы и размещает новые элементы в выделенной папке в хранилище Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-168">Copies the specified items and puts the new items in a designated folder in the Exchange store.</span></span>|
|[<span data-ttu-id="699fe-169">Операция CreateFolder</span><span class="sxs-lookup"><span data-stu-id="699fe-169">CreateFolder operation</span></span>](/exchange/client-developer/web-service-reference/createfolder-operation)|<span data-ttu-id="699fe-170">Создает папки в выбранном расположении в хранилище Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-170">Creates folders in the specified location in the Exchange store.</span></span>|
|[<span data-ttu-id="699fe-171">Операция CreateItem</span><span class="sxs-lookup"><span data-stu-id="699fe-171">CreateItem operation</span></span>](/exchange/client-developer/web-service-reference/createitem-operation)|<span data-ttu-id="699fe-172">Создает заданные элементы в хранилище Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-172">Creates the specified items in the Exchange store.</span></span>|
|[<span data-ttu-id="699fe-173">Операция ExpandDL</span><span class="sxs-lookup"><span data-stu-id="699fe-173">ExpandDL operation</span></span>](/exchange/client-developer/web-service-reference/expanddl-operation)|<span data-ttu-id="699fe-174">Отображает полное членство списков рассылки.</span><span class="sxs-lookup"><span data-stu-id="699fe-174">Displays the full membership of distribution lists.</span></span>|
|[<span data-ttu-id="699fe-175">Операция FindConversation</span><span class="sxs-lookup"><span data-stu-id="699fe-175">FindConversation operation</span></span>](/exchange/client-developer/web-service-reference/findconversation-operation)|<span data-ttu-id="699fe-176">Перечисляет список бесед в определенной папке в хранилище Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-176">Enumerates a list of conversations in the specified folder in the Exchange store.</span></span>|
|[<span data-ttu-id="699fe-177">Операция FindFolder</span><span class="sxs-lookup"><span data-stu-id="699fe-177">FindFolder operation</span></span>](/exchange/client-developer/web-service-reference/findfolder-operation)|<span data-ttu-id="699fe-178">Ищет вложенные папки заданной папки и возвращает набор свойств, описывающих вложенные папки.</span><span class="sxs-lookup"><span data-stu-id="699fe-178">Finds subfolders of an identified folder and returns a set of properties that describe the set of subfolders.</span></span>|
|[<span data-ttu-id="699fe-179">Операция FindItem</span><span class="sxs-lookup"><span data-stu-id="699fe-179">FindItem operation</span></span>](/exchange/client-developer/web-service-reference/finditem-operation)|<span data-ttu-id="699fe-180">Определяет элементы, расположенные в определенной папке в хранилище Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-180">Identifies items that are located in a specified folder in the Exchange store.</span></span>|
|[<span data-ttu-id="699fe-181">Операция GetConversationItems</span><span class="sxs-lookup"><span data-stu-id="699fe-181">GetConversationItems operation</span></span>](/exchange/client-developer/web-service-reference/getconversationitems-operation)|<span data-ttu-id="699fe-182">Получает один или несколько наборов элементов, упорядоченных в узлы в беседе.</span><span class="sxs-lookup"><span data-stu-id="699fe-182">Gets one or more sets of items that are organized in nodes in a conversation.</span></span>|
|[<span data-ttu-id="699fe-183">Операция GetFolder</span><span class="sxs-lookup"><span data-stu-id="699fe-183">GetFolder operation</span></span>](/exchange/client-developer/web-service-reference/getfolder-operation)|<span data-ttu-id="699fe-184">Получает определенные свойства и содержимое папок из хранилища Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-184">Gets the specified properties and contents of folders from the Exchange store.</span></span>|
|[<span data-ttu-id="699fe-185">Операция GetItem</span><span class="sxs-lookup"><span data-stu-id="699fe-185">GetItem operation</span></span>](/exchange/client-developer/web-service-reference/getitem-operation)|<span data-ttu-id="699fe-186">Получает определенные свойства и содержимое элементов из хранилища Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-186">Gets the specified properties and contents of items from the Exchange store.</span></span>|
|[<span data-ttu-id="699fe-187">Операция GetUserAvailability</span><span class="sxs-lookup"><span data-stu-id="699fe-187">GetUserAvailability operation</span></span>](/exchange/client-developer/web-service-reference/getuseravailability-operation)|<span data-ttu-id="699fe-188">Предоставляет подробные сведения о доступности наборов пользователей, помещений и ресурсов в рамках определенного периода времени.</span><span class="sxs-lookup"><span data-stu-id="699fe-188">Provides detailed information about the availability of a set of users, rooms, and resources within a specified time period.</span></span>|
|[<span data-ttu-id="699fe-189">Операция MarkAsJunk</span><span class="sxs-lookup"><span data-stu-id="699fe-189">MarkAsJunk operation</span></span>](/exchange/client-developer/web-service-reference/markasjunk-operation)|<span data-ttu-id="699fe-190">Перемещает сообщения электронной почты в папку "Нежелательная почта" и соответствующим образом добавляет или удаляет отправителей сообщений в списке заблокированных отправителей.</span><span class="sxs-lookup"><span data-stu-id="699fe-190">Moves email messages to the Junk Email folder, and adds or removes senders of the messages from the blocked senders list accordingly.</span></span>|
|[<span data-ttu-id="699fe-191">Операция MoveItem</span><span class="sxs-lookup"><span data-stu-id="699fe-191">MoveItem operation</span></span>](/exchange/client-developer/web-service-reference/moveitem-operation)|<span data-ttu-id="699fe-192">Перемещает элементы в одну целевую папку в хранилище Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-192">Moves items to a single destination folder in the Exchange store.</span></span>|
|[<span data-ttu-id="699fe-193">Операция ResolveNames</span><span class="sxs-lookup"><span data-stu-id="699fe-193">ResolveNames operation</span></span>](/exchange/client-developer/web-service-reference/resolvenames-operation)|<span data-ttu-id="699fe-194">Сопоставляет неоднозначные адреса электронной почты и отображает имена.</span><span class="sxs-lookup"><span data-stu-id="699fe-194">Resolves ambiguous email addresses and display names.</span></span>|
|[<span data-ttu-id="699fe-195">Операция SendItem</span><span class="sxs-lookup"><span data-stu-id="699fe-195">SendItem operation</span></span>](/exchange/client-developer/web-service-reference/senditem-operation)|<span data-ttu-id="699fe-196">Отправляет сообщения электронной почты, расположенные в хранилище Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-196">Sends email messages that are located in the Exchange store.</span></span>|
|[<span data-ttu-id="699fe-197">Операция UpdateFolder</span><span class="sxs-lookup"><span data-stu-id="699fe-197">UpdateFolder operation</span></span>](/exchange/client-developer/web-service-reference/updatefolder-operation)|<span data-ttu-id="699fe-198">Изменяет свойства существующих папок в хранилище Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-198">Modifies the properties of existing folders in the Exchange store.</span></span>|
|[<span data-ttu-id="699fe-199">Операция UpdateItem</span><span class="sxs-lookup"><span data-stu-id="699fe-199">UpdateItem operation</span></span>](/exchange/client-developer/web-service-reference/updateitem-operation)|<span data-ttu-id="699fe-200">Изменяет свойства существующих элементов в хранилище Exchange.</span><span class="sxs-lookup"><span data-stu-id="699fe-200">Modifies the properties of existing items in the Exchange store.</span></span>|

 > [!NOTE]
 > <span data-ttu-id="699fe-201">Элементы FAI (сведения, связанные с папками) нельзя обновлять (или создавать) из надстройки.</span><span class="sxs-lookup"><span data-stu-id="699fe-201">FAI (Folder Associated Information) items cannot be updated (or created) from an add-in.</span></span> <span data-ttu-id="699fe-202">Эти скрытые сообщения находятся в папке и используются для хранения различных параметров и вспомогательных данных.</span><span class="sxs-lookup"><span data-stu-id="699fe-202">These hidden messages are stored in a folder and are used to store a variety of settings and auxiliary data.</span></span>  <span data-ttu-id="699fe-203">При попытке использовать операцию UpdateItem возникнет ошибка ErrorAccessDenied: "У расширения Office нет разрешения на обновление такого элемента".</span><span class="sxs-lookup"><span data-stu-id="699fe-203">Attempting to use the UpdateItem operation will throw an ErrorAccessDenied error: "Office extension is not allowed to update this type of item".</span></span> <span data-ttu-id="699fe-204">В качестве альтернативы можно использовать [управляемый API служб EWS](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) для обновления этих элементов в клиентском или серверном приложении для Windows.</span><span class="sxs-lookup"><span data-stu-id="699fe-204">As an alternative, you may use the [EWS Managed API](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) to update these items from a Windows client or a server application.</span></span> <span data-ttu-id="699fe-205">Рекомендуем соблюдать осторожность, так как внутренние структуры данных для служб могут меняться и сделать решение неработоспособным.</span><span class="sxs-lookup"><span data-stu-id="699fe-205">Caution is recommended as internal, service-type data structures are subject to change and could break your solution.</span></span>


## <a name="authentication-and-permission-considerations-for-makeewsrequestasync"></a><span data-ttu-id="699fe-206">Разрешения и проверка подлинности для makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="699fe-206">Authentication and permission considerations for makeEwsRequestAsync</span></span>

<span data-ttu-id="699fe-207">При использовании метода **makeEwsRequestAsync** запрос проходит проверку подлинности с помощью данных учетной записи электронной почты текущего пользователя.</span><span class="sxs-lookup"><span data-stu-id="699fe-207">When you use the **makeEwsRequestAsync** method, the request is authenticated by using the email account credentials of the current user.</span></span> <span data-ttu-id="699fe-208">Метод **makeEwsRequestAsync** управляет учетными данными, чтобы вам не нужно было предоставлять учетные данные для проверки подлинности с вашим запросом.</span><span class="sxs-lookup"><span data-stu-id="699fe-208">The **makeEwsRequestAsync** method manages the credentials for you so that you do not have to provide authentication credentials with your request.</span></span>

> [!NOTE]
> <span data-ttu-id="699fe-209">Администратор сервера должен использовать командлет [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) или [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps), чтобы установить для параметра _OAuthAuthentication_ значение **true** в каталоге EWS сервера клиентского доступа, чтобы метод **makeEwsRequestAsync** мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="699fe-209">The server administrator must use the [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) or the [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) cmldet to set the _OAuthAuthentication_ parameter to **true** on the Client Access server EWS directory in order to enable the **makeEwsRequestAsync** method to make EWS requests.</span></span>

<span data-ttu-id="699fe-210">Надстройка должна указать разрешение **ReadWriteMailbox** в своем манифесте, чтобы использовать метод **makeEwsRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="699fe-210">Your add-in must specify the **ReadWriteMailbox** permission in its add-in manifest to use the **makeEwsRequestAsync** method.</span></span> <span data-ttu-id="699fe-211">Сведения об использовании разрешения **ReadWriteMailbox** см. в разделе [Разрешение ReadWriteMailbox](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) в статье [Общие сведения о разрешениях для надстроек Outlook](understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="699fe-211">For information about using the **ReadWriteMailbox** permission, see the section [ReadWriteMailbox permission](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) in [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

> [!NOTE]
> <span data-ttu-id="699fe-212">Администратор сервера должен использовать командлет [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) или [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps), чтобы установить для параметра _OAuthAuthentication_ значение **true** в каталоге EWS сервера клиентского доступа, чтобы метод **makeEwsRequestAsync** мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="699fe-212">The server administrator must use the [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) or the [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) cmldet to set the _OAuthAuthentication_ parameter to **true** on the Client Access server EWS directory in order to enable the **makeEwsRequestAsync** method to make EWS requests.</span></span>



## <a name="see-also"></a><span data-ttu-id="699fe-213">См. также</span><span class="sxs-lookup"><span data-stu-id="699fe-213">See also</span></span>

- [<span data-ttu-id="699fe-214">Конфиденциальность и безопасность надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="699fe-214">Privacy and security for Office Add-ins</span></span>](../develop/privacy-and-security.md)   
- [<span data-ttu-id="699fe-215">Работа с ограничениями по принципу одинакового источника в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="699fe-215">Addressing same-origin policy limitations in Office Add-ins</span></span>](../develop/addressing-same-origin-policy-limitations.md)
- [<span data-ttu-id="699fe-216">Справка по службам EWS для Exchange</span><span class="sxs-lookup"><span data-stu-id="699fe-216">EWS reference for Exchange</span></span>](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)   
- [<span data-ttu-id="699fe-217">Приложения электронной почты для Outlook и служб EWS в Exchange</span><span class="sxs-lookup"><span data-stu-id="699fe-217">Mail apps for Outlook and EWS in Exchange</span></span>](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)
   
<span data-ttu-id="699fe-218">Сведения о создании внутренних служб для надстроек с помощью веб-API ASP.NET см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="699fe-218">See the following for creating backend services for add-ins using ASP.NET Web API:</span></span>

- [<span data-ttu-id="699fe-219">Создание веб-службы надстройки для Office с использованием веб-API ASP.NET</span><span class="sxs-lookup"><span data-stu-id="699fe-219">Create a web service for an Office Add-in using the ASP.NET Web API</span></span>](https://blogs.msdn.microsoft.com/officeapps/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api/)    
- [<span data-ttu-id="699fe-220">Основы создания службы HTTP с использованием веб-API ASP.NET</span><span class="sxs-lookup"><span data-stu-id="699fe-220">The basics of building an HTTP service using ASP.NET Web API</span></span>](https://www.asp.net/web-api)
    