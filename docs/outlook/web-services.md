---
title: Использование веб-служб Exchange (EWS) из надстройки Outlook
description: Содержит пример, в котором показано, как надстройка Outlook может запрашивать сведения из веб-службы Exchange.
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: 63c969355c9bae5dab6ef8603a9f3d61d8e82eec
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348456"
---
# <a name="call-web-services-from-an-outlook-add-in"></a>Вызов веб-служб из надстройки Outlook

Ваша надстройка может использовать веб-службы Exchange (EWS) на компьютере с Exchange Server 2013; веб-службу, доступную на сервере, предоставляющем исходное расположение для пользовательского интерфейса надстройки; или веб-службу, доступную через Интернет. В этой статье приведен пример того, как надстройка Outlook может запрашивать данные из EWS.

Способы вызова веб-службы различаются в зависимости от расположения службы. В таблице 1 приведены различные способы вызова веб-службы в зависимости от расположения.


**Таблица 1. Способы вызова веб-служб из надстройки Outlook**

<br/>

|**Расположение веб-службы**|**Способ вызова веб-службы**|
|:-----|:-----|
|Сервер Exchange, на котором размещен почтовый ящик клиента|Используйте метод [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) для вызова операций EWS, поддерживаемых надстройками. Сервер Exchange Server, на котором размещен почтовый ящик, также предоставляет доступ к EWS.|
|Веб-сервер, предоставляющий исходное расположение для пользовательского интерфейса надстроек.|Вызывайте веб-службу с помощью стандартных методик JavaScript. Код JavaScript в пределах пользовательского интерфейса работает в контексте веб-сервера, предоставляющего пользовательский интерфейс. Поэтому он сможет вызывать веб-службы на этом сервере, не создавая ошибки межсайтового скрипта.|
|Все другие расположения|Создайте прокси для веб-службы на веб-сервере, предоставляющем исходное расположение для пользовательского интерфейса. Если не указать прокси, надстройка не запустится из-за ошибок межсайтовых сценариев. Один из способов указать такой прокси — это использовать JSON/P. Дополнительные сведения см. в статье [Конфиденциальность и безопасность надстроек для Office](../concepts/privacy-and-security.md).|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a>Получение доступа к операциям веб-служб Exchange с помощью метода makeEwsRequestAsync

С помощью метода [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) вы можете отправить запрос EWS на сервер Exchange Server, на котором размещается почтовый ящик пользователя.

Веб-службы Exchange поддерживают различные операции на сервере Exchange. Например, операции копирования, поиска, обновления или отправки на уровне элемента, а также операции создания, получения или обновления на уровне папки. Чтобы выполнить операцию веб-служб Exchange, создайте для нее SOAP-запрос в формате XML. После завершения операции будет возвращен SOAP-ответ в формате XML с необходимыми данными. SOAP-запросы к веб-службам Exchange и их SOAP-ответы соответствуют схеме, определенной в файле Messages.xsd. Как и другие файлы схемы веб-служб Exchange, файл Message.xsd расположен в виртуальном каталоге IIS, в котором размещены веб-службы Exchange.

Чтобы использовать метод `makeEwsRequestAsync` для инициации операции EWS, ука-

- XML-код SOAP-запроса для соответствующей операции EWS в качестве аргумента для параметра  _data_;

- метод обратного вызова (в качестве аргумента  _callback_);

- все необязательные входные данные для этого метода обратного вызова (в качестве аргумента  _userContext_).

Когда запрос EWS SOAP завершен, Outlook вызывает метод вызова с одним аргументом, который является [объектом AsyncResult.](/javascript/api/office/office.asyncresult) Метод вызова может получить доступ к двум свойствам объекта: свойству, которое содержит ответ XML SOAP операции EWS, и необязательно свойство, которое содержит все данные, переданные в качестве `AsyncResult` `value` `asyncContext` `userContext` параметра. Обычно метод вызываемого вызова затем обрабатывает XML в ответе SOAP, чтобы получить любую соответствующую информацию, и обрабатывает эту информацию соответствующим образом.


## <a name="tips-for-parsing-ews-responses"></a>Советы по анализу ответов веб-служб Exchange

При разрисовыве ответа SOAP из операции EWS обратите внимание на следующие проблемы, зависящие от браузера.


- Укажите префикс для имени тега при использовании метода DOM, чтобы включить `getElementsByTagName` поддержку Internet Explorer.

  `getElementsByTagName` ведет себя по-разному в зависимости от типа браузера. Например, ответ EWS может содержать следующий XML (форматированный и сокращенный для отображения).

   ```XML
        <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
        PropertyName="MyProperty" 
        PropertyType="String"/>
        <t:Value>{
        ...
        }</t:Value></t:ExtendedProperty>
   ```

   Код, как и в следующем, будет работать в браузере, как Chrome, чтобы получить XML, закрытый `ExtendedProperty` тегами.

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("ExtendedProperty")
            });
   ```

   В Internet Explorer необходимо включить префикс имени тега `t:` следующим образом.

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("t:ExtendedProperty")
            });
   ```

- Используйте свойство DOM, чтобы получить содержимое тега в `textContent` ответе EWS следующим образом.

   ```js
      content = $.parseJSON(value.textContent);
   ```

   Другие свойства, `innerHTML` например, могут не работать в Internet Explorer для некоторых тегов в ответе EWS.


## <a name="example"></a>Пример

В следующем `makeEwsRequestAsync` примере вызывается использование [операции GetItem](/exchange/client-developer/web-service-reference/getitem-operation) для получения субъекта элемента. В этом примере содержатся следующие три функции.

- `getSubjectRequest`Принимает ID элемента в качестве ввода и возвращает XML для запроса SOAP для вызова &ndash; `GetItem` указанного элемента.

- `sendRequest`Вызывает запрос SOAP для выбранного элемента, затем передает запрос SOAP и метод вызова, чтобы получить субъект &ndash;  `getSubjectRequest` `callback` `makeEwsRequestAsync` указанного элемента.

- `callback` &ndash; обрабатывает SOAP-ответ, включающий тему и другие сведения об указанном элементе.


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


## <a name="ews-operations-that-add-ins-support"></a>Операции веб-служб Exchange, которые надстройки поддерживают

Outlook надстройки могут получать доступ к подмножество операций, доступных в EWS с помощью `makeEwsRequestAsync` метода. Если вы не знакомы с операциями EWS и как использовать метод для доступа к операции, начните с примера запроса SOAP для настройки `makeEwsRequestAsync` _аргумента данных._

Ниже описано, как можно использовать `makeEwsRequestAsync` метод.

1. В XML-коде замените все идентификаторы элементов и релевантные атрибуты операций EWS на соответствующие значения.

1. Включи запрос SOAP в качестве аргумента для  _параметра данных_ `makeEwsRequestAsync` .

1. Укажите метод вызова и вызов `makeEwsRequestAsync` .

1. В методе обратного вызова проверьте результаты операции в SOAP-ответе.

1. Используйте результаты операции EWS в соответствии с вашими потребностями.

В следующей таблице указаны операции EWS, которые надстройки поддерживают. Чтобы просмотреть примеры SOAP-запросов и SOAP-ответов, выберите ссылку для каждой операции. Дополнительные сведения об операциях EWS см. в статье [Операции EWS в Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).

**Таблица 2. Поддерживаемые операции EWS**

<br/>

|**Операция служб EWS**|**Описание**|
|:-----|:-----|
|[Операция CopyItem](/exchange/client-developer/web-service-reference/copyitem-operation)|Копирует выбранные элементы и размещает новые элементы в выделенной папке в хранилище Exchange.|
|[Операция CreateFolder](/exchange/client-developer/web-service-reference/createfolder-operation)|Создает папки в выбранном расположении в хранилище Exchange.|
|[Операция CreateItem](/exchange/client-developer/web-service-reference/createitem-operation)|Создает заданные элементы в хранилище Exchange.|
|[Операция ExpandDL](/exchange/client-developer/web-service-reference/expanddl-operation)|Отображает полное членство списков рассылки.|
|[Операция FindConversation](/exchange/client-developer/web-service-reference/findconversation-operation)|Перечисляет список бесед в определенной папке в хранилище Exchange.|
|[Операция FindFolder](/exchange/client-developer/web-service-reference/findfolder-operation)|Ищет вложенные папки заданной папки и возвращает набор свойств, описывающих вложенные папки.|
|[Операция FindItem](/exchange/client-developer/web-service-reference/finditem-operation)|Определяет элементы, расположенные в определенной папке в хранилище Exchange.|
|[Операция GetConversationItems](/exchange/client-developer/web-service-reference/getconversationitems-operation)|Получает один или несколько наборов элементов, упорядоченных в узлы в беседе.|
|[Операция GetFolder](/exchange/client-developer/web-service-reference/getfolder-operation)|Получает определенные свойства и содержимое папок из хранилища Exchange.|
|[Операция GetItem](/exchange/client-developer/web-service-reference/getitem-operation)|Получает определенные свойства и содержимое элементов из хранилища Exchange.|
|[Операция GetUserAvailability](/exchange/client-developer/web-service-reference/getuseravailability-operation)|Предоставляет подробные сведения о доступности наборов пользователей, помещений и ресурсов в рамках определенного периода времени.|
|[Операция MarkAsJunk](/exchange/client-developer/web-service-reference/markasjunk-operation)|Перемещает сообщения электронной почты в папку "Нежелательная почта" и соответствующим образом добавляет или удаляет отправителей сообщений в списке заблокированных отправителей.|
|[Операция MoveItem](/exchange/client-developer/web-service-reference/moveitem-operation)|Перемещает элементы в одну целевую папку в хранилище Exchange.|
|[Операция ResolveNames](/exchange/client-developer/web-service-reference/resolvenames-operation)|Сопоставляет неоднозначные адреса электронной почты и отображает имена.|
|[Операция SendItem](/exchange/client-developer/web-service-reference/senditem-operation)|Отправляет сообщения электронной почты, расположенные в хранилище Exchange.|
|[Операция UpdateFolder](/exchange/client-developer/web-service-reference/updatefolder-operation)|Изменяет свойства существующих папок в хранилище Exchange.|
|[Операция UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)|Изменяет свойства существующих элементов в хранилище Exchange.|

 > [!NOTE]
 > Элементы FAI (сведения, связанные с папками) нельзя обновлять (или создавать) из надстройки. Эти скрытые сообщения находятся в папке и используются для хранения различных параметров и вспомогательных данных.  При попытке использовать операцию UpdateItem возникнет ошибка ErrorAccessDenied: "У расширения Office нет разрешения на обновление такого элемента". В качестве альтернативы можно использовать [управляемый API служб EWS](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) для обновления этих элементов в клиентском или серверном приложении для Windows. Рекомендуем соблюдать осторожность, так как внутренние структуры данных для служб могут меняться и сделать решение неработоспособным.


## <a name="authentication-and-permission-considerations-for-makeewsrequestasync"></a>Разрешения и проверка подлинности для makeEwsRequestAsync

При использовании метода запрос проверки подлинности с помощью учетных данных учетной записи электронной почты `makeEwsRequestAsync` текущего пользователя. Метод управляет учетными данными для вас, чтобы не предоставлять учетные данные проверки подлинности `makeEwsRequestAsync` с вашим запросом.

> [!NOTE]
> Администратор сервера должен использовать  [new-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) или групплет [Set-WebServicesVirtualDirectory,](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) чтобы установить параметр _OAuthAuthAuthentication_ в каталоге EWS сервера клиентского доступа, чтобы включить метод для запросов `makeEwsRequestAsync` EWS.

Ваша надстройка должна указать разрешение в манифесте надстройки `ReadWriteMailbox` для `makeEwsRequestAsync` использования метода. Сведения об использовании разрешения см. в разделе `ReadWriteMailbox` [Разрешение ReadWriteMailbox](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) в разделе Outlook разрешений [надстройки.](understanding-outlook-add-in-permissions.md)

## <a name="see-also"></a>См. также

- [Конфиденциальность и безопасность надстроек для Office](../concepts/privacy-and-security.md)
- [Работа с ограничениями по принципу одинакового источника в надстройках Office](../develop/addressing-same-origin-policy-limitations.md)
- [Справка по службам EWS для Exchange](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)
- [Приложения электронной почты для Outlook и служб EWS в Exchange](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)

См. ниже ниже для создания дополнительных служб для надстройок с веб-API ASP.NET.

- [Создание веб-службы надстройки для Office с использованием веб-API ASP.NET](/archive/blogs/officeapps/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api)
- [Основы создания службы HTTP с использованием веб-API ASP.NET](https://dotnet.microsoft.com/apps/aspnet/apis)