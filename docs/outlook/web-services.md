---
title: Использование веб-служб Exchange (EWS) из надстройки Outlook
description: Содержит пример, в котором показано, как надстройка Outlook может запрашивать сведения из веб-службы Exchange.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 94fff26fc7f9c16e2e385d6c44c128e4b03f968e
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467015"
---
# <a name="call-web-services-from-an-outlook-add-in"></a>Вызов веб-служб из надстройки Outlook

Your add-in can use Exchange Web Services (EWS) from a computer that is running Exchange Server 2013, a web service that is available on the server that provides the source location for the add-in's UI, or a web service that is available on the Internet. This article provides an example that shows how an Outlook add-in can request information from EWS.

The way that you call a web service varies based on where the web service is located. Table 1 lists the different ways that you can call a web service based on location.

**Таблица 1. Способы вызова веб-служб из надстройки Outlook**

|**Расположение веб-службы**|**Способ вызова веб-службы**|
|:-----|:-----|
|Сервер Exchange, на котором размещен почтовый ящик клиента|Use the [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method to call EWS operations that add-ins support. The Exchange server that hosts the mailbox also exposes EWS.|
|Веб-сервер, предоставляющий исходное расположение для пользовательского интерфейса надстроек.|Call the web service by using standard JavaScript techniques. The JavaScript code in the UI frame runs in the context of the web server that provides the UI. Therefore, it can call web services on that server without causing a cross-site scripting error.|
|Все другие расположения|Create a proxy for the web service on the web server that provides the source location for the UI. If you do not provide a proxy, cross-site scripting errors will prevent your add-in from running. One way to provide a proxy is by using JSON/P. For more information, see [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md).|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a>Получение доступа к операциям веб-служб Exchange с помощью метода makeEwsRequestAsync

С помощью метода [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) вы можете отправить запрос EWS на сервер Exchange Server, на котором размещается почтовый ящик пользователя.

EWS supports different operations on an Exchange server; for example, item-level operations to copy, find, update, or send an item, and folder-level operations to create, get, or update a folder. To perform an EWS operation, create an XML SOAP request for that operation. When the operation finishes, you get an XML SOAP response that contains data that is relevant to the operation. EWS SOAP requests and responses follow the schema defined in the Messages.xsd file. Like other EWS schema files, the Message.xsd file is located in the IIS virtual directory that hosts EWS.

Чтобы использовать этот `makeEwsRequestAsync` метод для инициации операции EWS, укажите следующее:

- XML-код SOAP-запроса для соответствующей операции EWS в качестве аргумента для параметра  _data_;

- Функция обратного вызова (в качестве  _аргумента обратного_ вызова)

- Любые необязательные входные данные для этой функции обратного вызова (в качестве  _аргумента userContext_ )

После завершения запроса SOAP EWS Outlook вызывает функцию обратного вызова с одним аргументом, который является [объектом AsyncResult](/javascript/api/office/office.asyncresult) . Функция обратного `AsyncResult` вызова может получить доступ к двум свойствам объекта: `value` свойству, содержащее ответ XML SOAP операции EWS, и, при необходимости, `asyncContext` свойству, которое содержит любые данные, `userContext` передаваемые в качестве параметра. Как правило, функция обратного вызова анализирует XML-код в ответе SOAP, чтобы получить любую соответствующую информацию, и обрабатывает эту информацию соответствующим образом.

## <a name="tips-for-parsing-ews-responses"></a>Советы по анализу ответов веб-служб Exchange

При анализе ответа SOAP из операции EWS обратите внимание на следующие проблемы, зависящие от браузера.

- Укажите префикс для имени тега при использовании метода DOM `getElementsByTagName`, чтобы включить поддержку Internet Explorer.

  `getElementsByTagName` ведет себя по-разному в зависимости от типа браузера. Например, ответ EWS может содержать следующий XML-код (отформатированный и сокращенный для отображения).

   ```XML
   <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
   PropertyName="MyProperty" 
   PropertyType="String"/>
   <t:Value>{
   ...
   }</t:Value></t:ExtendedProperty>
   ```

   Код, как показано ниже, будет работать в браузере, например Chrome, чтобы получить XML-код, заключенный в `ExtendedProperty` теги.

   ```js
   const mailbox = Office.context.mailbox;
   mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
       const response = $.parseXML(result.value);
       const extendedProps = response.getElementsByTagName("ExtendedProperty")
   });
   ```

   В Internet Explorer необходимо `t:` включить префикс имени тега, как показано ниже.

   ```js
   const mailbox = Office.context.mailbox;
   mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
       const response = $.parseXML(result.value);
       const extendedProps = response.getElementsByTagName("t:ExtendedProperty")
   });
   ```

- Используйте свойство DOM для `textContent` получения содержимого тега в ответе EWS, как показано ниже.

   ```js
   content = $.parseJSON(value.textContent);
   ```

   Другие свойства, например `innerHTML` , могут не работать в Internet Explorer для некоторых тегов в ответе EWS.

## <a name="example"></a>Пример

В следующем примере вызывается `makeEwsRequestAsync` использование [операции GetItem](/exchange/client-developer/web-service-reference/getitem-operation) для получения темы элемента. В этом примере содержатся следующие три функции.

- `getSubjectRequest`&ndash; Принимает идентификатор элемента в качестве входных данных и возвращает XML-код запроса SOAP `GetItem` для вызова указанного элемента.

- `sendRequest`&ndash; Вызывает `getSubjectRequest` запрос SOAP для выбранного элемента, а затем передает запрос SOAP и функцию обратного вызова, `callback``makeEwsRequestAsync` чтобы получить тему указанного элемента.

- `callback` &ndash; обрабатывает SOAP-ответ, включающий тему и другие сведения об указанном элементе.

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   const result = 
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
   const mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   const result = asyncResult.value;
   const context = asyncResult.context;

   // Process the returned response here.
}
```

## <a name="ews-operations-that-add-ins-support"></a>Операции веб-служб Exchange, которые надстройки поддерживают

С помощью этого метода надстройки Outlook могут получить доступ к подмножество операций, доступных в EWS `makeEwsRequestAsync` . Если вы не знакомы с операциями EWS `makeEwsRequestAsync` и как использовать метод для доступа к операции, начните с примера запроса SOAP для настройки _аргумента_ данных.

Ниже описано, как можно использовать `makeEwsRequestAsync` метод.

1. В XML-коде замените все идентификаторы элементов и релевантные атрибуты операций EWS на соответствующие значения.

1. Включите soap-запрос в качестве аргумента для _параметра_ данных .`makeEwsRequestAsync`

1. Укажите функцию обратного вызова и вызовите ее `makeEwsRequestAsync`.

1. В функции обратного вызова проверьте результаты операции в ответе SOAP.

1. Используйте результаты операции EWS в соответствии с вашими потребностями.

The following table lists the EWS operations that add-ins support. To see examples of SOAP requests and responses, choose the link for each operation. For more information about EWS operations, see [EWS operations in Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).

**Таблица 2. Поддерживаемые операции EWS**

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

При использовании метода `makeEwsRequestAsync` проверка подлинности запроса выполняется с использованием учетных данных учетной записи электронной почты текущего пользователя. Этот `makeEwsRequestAsync` метод управляет учетными данными, чтобы вам не нужно было предоставлять учетные данные для проверки подлинности в запросе.

> [!NOTE]
> Администратор сервера должен использовать [командлет New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) или [Командлет Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) , чтобы задать для параметра _OAuthAuthentication_ `true` значение в каталоге EWS сервера клиентского доступа, `makeEwsRequestAsync` чтобы разрешить методу выполнять запросы EWS.

Чтобы использовать этот `makeEwsRequestAsync` метод, надстройка должна запросить разрешение на чтение **и** запись почтового ящика в манифесте. Разметка зависит от типа манифеста.

- **XML-манифест**: задайте **\<Permissions\>** для элемента **значение ReadWriteMailbox**.
- **Манифест Teams (** предварительная версия): задайте для свойства name объекта в массиве authorization.permissions.resourceSpecific значение Mailbox.ReadWrite.User.

Сведения об использовании разрешения на чтение и **запись** почтового ящика см. в разделе разрешений на чтение [и запись почтового ящика](understanding-outlook-add-in-permissions.md#readwrite-mailbox-permission).

## <a name="see-also"></a>См. также

- [Конфиденциальность и безопасность надстроек для Office](../concepts/privacy-and-security.md)
- [Работа с ограничениями по принципу одинакового источника в надстройках Office](../develop/addressing-same-origin-policy-limitations.md)
- [Справка по службам EWS для Exchange](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)
- [Приложения электронной почты для Outlook и служб EWS в Exchange](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)

См. следующие сведения о создании серверных служб для надстроек с помощью веб-API ASP.NET.

- [Создание веб-службы надстройки для Office с использованием веб-API ASP.NET](/archive/blogs/officeapps/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api)
- [Основы создания службы HTTP с использованием веб-API ASP.NET](https://dotnet.microsoft.com/apps/aspnet/apis)
