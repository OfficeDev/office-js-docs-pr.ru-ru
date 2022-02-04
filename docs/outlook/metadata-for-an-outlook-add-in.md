---
title: Просмотр и изменение метаданных элемента в надстройке Outlook
description: Управление пользовательскими данными в надстройке Outlook с помощью параметров перемещения или настраиваемых свойств.
ms.date: 10/31/2019
ms.localizationpriority: medium
---

# <a name="get-and-set-add-in-metadata-for-an-outlook-add-in"></a>Просмотр и изменение метаданных для надстройки Outlook

Для управления пользовательскими данными в настройке Outlook можно использовать следующее:

- параметры перемещения, которые управляют пользовательскими данными для почтового ящика пользователя;
- настраиваемые свойства, которые управляют пользовательскими данными для элемента в почтовом ящике пользователя.

Оба этих способа предоставляют доступ к пользовательским данным, доступным только надстройке Outlook, но каждый метод хранит данные отдельно от остальных. Другими словами, данные, хранящиеся с помощью параметров перемещения, недоступны настраиваемым свойствам и наоборот. Данные хранятся на сервере этого почтового ящика и доступны в последующих сеансах Outlook на всех поддерживаемых надстройкой форм-факторах.

## <a name="custom-data-per-mailbox-roaming-settings"></a>Пользовательские данные на один почтовый ящик: параметры перемещения

Вы можете указать данные, специфичные для пользователя почтового ящика Exchange, с помощью объекта [RoamingSettings](/javascript/api/outlook/office.roamingsettings). Примерами таких данных являются личные данные и предпочтения пользователя. Ваша почтовая надстройка может получить доступ к параметрам перемещения, когда перемещение происходит на любом из устройств, предназначенных для работы (настольный ПК, планшет или смартфон).

Изменения этих данных хранятся в памяти текущего сеанса Outlook. После изменения все параметры перемещения следует сохранить, чтобы они были доступны, когда пользователь откроет надстройку на том же или другом поддерживаемом устройстве в следующий раз.


### <a name="roaming-settings-format"></a>Формат параметров перемещения

Данные в объекте **RoamingSettings** хранятся в виде сериализованной строки нотации объектов JavaScript (JSON). 

Ниже приведен пример структуры для трех определенных параметров перемещения с именами `add-in_setting_name_0`, `add-in_setting_name_1`, и `add-in_setting_name_2`.


```json
{
  "add-in_setting_name_0": "add-in_setting_value_0",
  "add-in_setting_name_1": "add-in_setting_value_1",
  "add-in_setting_name_2": "add-in_setting_value_2"
}
```


### <a name="loading-roaming-settings"></a>Загрузка параметров перемещения

Надстройка почты обычно загружает параметры перемещения в обработчик событий [Office.initialize](/javascript/api/office#Office_initialize_reason_). В следующем примере кода JavaScript показано, как загрузить существующие параметры роуминга и получить значения 2 параметров, **имя** клиента и **customerBalance**.


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### <a name="creating-or-assigning-a-roaming-setting"></a>Создание или назначение параметра перемещения

Развивая предыдущий пример, следующая функция JavaScript `setAddInSetting` показывает, как использовать метод [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings) для определения заданного параметра `cookie` с указанием сегодняшнего числа, и как сохраненить данных с помощью метода [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-saveasync-member(1)), чтобы сохранить все параметры перемещения на сервере.

Метод `set` создает параметр, если параметр еще не существует, и назначает параметр указанному значению. Метод `saveAsync` сохраняет параметры роуминга асинхронно. Этот пример кода передает метод вызова, `saveMyAddInSettingsCallback``saveAsync` когда асинхронный вызов завершается, `saveMyAddInSettingsCallback` вызван с помощью одного параметра _, asyncResult_. Этот параметр является объектом [AsyncResult](/javascript/api/office/office.asyncresult), который содержит результат и все сведения об асинхронном вызове. Необязательный параметр _userContext_ можно использовать для передачи сведений о состоянии из асинхронного вызова в функцию обратного звонка.

```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### <a name="removing-a-roaming-setting"></a>Удаление параметра перемещения

Кроме того, в расширениях предыдущих примеров следующая функция JavaScript —  `removeAddInSetting` — показывает, как метод [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-remove-member(1)) используется для удаления параметра `cookie` и сохранения всех параметров перемещения обратно в Exchange Server.


```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## <a name="custom-data-per-item-in-a-mailbox-custom-properties"></a>Пользовательские данные для каждого элемента в почтовом ящике: пользовательские свойства

Вы также можете указать данные, характерные для элемента в почтовом ящике пользователя, используя объект [CustomProperties](/javascript/api/outlook/office.customproperties). Например, ваша почтовая надстройка могла бы категоризировать некоторые сообщения и отмечать категорию с помощью настраиваемого свойства `messageCategory`. Либо, если ваша почтовая надстройка создает встречи из сообщений с предложениями о собрании, вы можете использовать настраиваемое свойство, чтобы отслеживать каждую из этих встреч. Это гарантирует, что если пользователь вновь откроет сообщение, ваша почтовая надстройка не станет во второй раз предлагать создать встречу.

Аналогично параметрам перемещения, изменения настраиваемых свойств хранятся в копии контейнера свойств для текущего сеанса Outlook. Чтобы эти настраиваемые свойства были доступны при следующем сеансе, используйте [CustomProperties.saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)).

Эти настраиваемые свойства, определенные для отдельных элементов, можно получить только с помощью `CustomProperties` объекта. Эти свойства отличаются от настраиваемого свойства [userProperties](/office/vba/api/Outlook.UserProperties) на основе MAPI в объектной модели Outlook и расширенными свойствами в Exchange Web Services (EWS). Вы не можете напрямую `CustomProperties` получить доступ с Outlook объектной модели, EWS или REST. Чтобы узнать, как получить доступ к `CustomProperties` EWS или REST, см. в разделе [Get custom properties using EWS или REST](#get-custom-properties-using-ews-or-rest).

### <a name="using-custom-properties"></a>Использование настраиваемых свойств

Перед использованием настраиваемых свойств необходимо загрузить их, вызвав метод [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods). После создания контейнера свойств можно использовать методы [set](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-set-member(1)) и [get](/javascript/api/outlook/office.customproperties) для добавления и извлечения настраиваемых свойств. Чтобы сохранить любые изменения, внесенные в контейнер свойств, необходимо использовать метод [saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)).


 > [!NOTE]
 > Так как Outlook для Mac не кэширует настраиваемые свойства, в случае перебоев в работе сети пользователя почтовые надстройки в Outlook для Mac не смогут получить доступ к их настраиваемым свойствам.


### <a name="custom-properties-example"></a>Пример пользовательских свойств


Следующий пример показывает простой набор методов для надстройки Outlook, использующей настраиваемые свойства. Этот пример можно использовать в качестве отправной точки для создания надстройки, использующей настраиваемые свойства.

В этом примере содержатся следующие методы.


- [Office.initialize](/javascript/api/office#Office_initialize_reason_): инициализирует надстройку и загружает контейнер настраиваемых свойств с сервера Exchange Server.

- **customPropsCallback**: получает контейнер настраиваемых свойств, возвращенный с сервера, и сохраняет его для дальнейшего использования.

- **updateProperty**: задает или обновляет определенное свойство, а затем сохраняет изменения на сервере.

- **removeProperty**: удаляет определенное свойство из контейнера свойств, а затем сохраняет удаление на сервере.


```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = _customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```

### <a name="get-custom-properties-using-ews-or-rest"></a>Просмотр настраиваемых свойств с помощью EWS или REST

Чтобы получить объект **CustomProperties** с помощью EWS или REST, необходимо сначала определить имя его расширенного свойства, основанного на интерфейсе MAPI. Затем можно получить это свойство способом, аналогичным используемому при получении любого расширенного свойства, основанного на интерфейсе MAPI.

#### <a name="how-custom-properties-are-stored-on-an-item"></a>Способ хранения настраиваемых свойств в элементе

Настраиваемые свойства, присвоенные надстройкой, отличаются от обычных свойств, основанных на интерфейсе MAPI. API надстройки `CustomProperties` сериализируют все надстройки в качестве полезной нагрузки JSON, а затем сохраните их в одном расширенном свойстве на основе MAPI `cecp-<app-guid>` , имя которого (`<app-guid>` это ID надстройки) и набор свойств GUID `{00020329-0000-0000-C000-000000000046}`. (Дополнительные сведения об этом объекте см. в статье [MS-OXCEXT 2.2.5 Настраиваемые свойства почтового приложения](/openspecs/exchange_server_protocols/ms-oxcext/4cf1da5e-c68e-433e-a97e-c45625483481)). Затем можно использовать EWS или REST, чтобы получить это свойство, основанное на интерфейсе MAPI.

#### <a name="get-custom-properties-using-ews"></a>Просмотр настраиваемых свойств с помощью EWS

Надстройка почты может получить расширенное `CustomProperties` свойство MAPI с помощью операции EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) . Доступ `GetItem` на стороне сервера с помощью маркера вызова или с клиентской стороны с помощью метода [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) . В запросе `GetItem` укажите `CustomProperties` свойство MAPI в наборе свойств с помощью сведений, предоставленных в предыдущем разделе Как настраиваемые свойства хранятся [на элементе](#how-custom-properties-are-stored-on-an-item).

В приведенном ниже примере показано, как получить элемент и его настраиваемые свойства.

> [!IMPORTANT]
> В приведенном ниже примере замените `<app-guid>` идентификатором своей надстройки.

```typescript
let request_str =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                   'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                   'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
                   'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '<soap:Header xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"' +
                     'xmlns:wsa="http://www.w3.org/2005/08/addressing">' +
            '<t:RequestServerVersion Version="Exchange2010_SP1"/>' +
        '</soap:Header>' +
        '<soap:Body>' +
            '<m:GetItem>' +
                '<m:ItemShape>' +
                    '<t:BaseShape>AllProperties</t:BaseShape>' +
                    '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
                    '<t:AdditionalProperties>' +
                        '<t:ExtendedFieldURI ' +
                          'DistinguishedPropertySetId="PublicStrings" ' +
                          'PropertyName="cecp-<app-guid>"' +
                          'PropertyType="String" ' +
                        '/>' +
                    '</t:AdditionalProperties>' +
                '</m:ItemShape>' +
                '<m:ItemIds>' +
                    '<t:ItemId Id="' +
                      Office.context.mailbox.item.itemId +
                    '"/>' +
                '</m:ItemIds>' +
            '</m:GetItem>' +
        '</soap:Body>' +
    '</soap:Envelope>';

Office.context.mailbox.makeEwsRequestAsync(
    request_str,
    function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(asyncResult.value);
        }
        else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

Также можно получить дополнительные настраиваемые свойства, если указать их в строке запроса как другие элементы [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri).

#### <a name="get-custom-properties-using-rest"></a>Просмотр настраиваемых свойств с помощью REST

В своей надстройке можно создать запрос REST для получения сообщений и событий, уже имеющих настраиваемые свойства. В запрос нужно включить расширенное свойство на основе интерфейса MAPI **CustomProperties** и его набор свойств с помощью сведений, указанных в разделе [Способ хранения настраиваемых свойств в элементе](#how-custom-properties-are-stored-on-an-item).

В приведенном ниже примере показано, как получить все события, которые содержат любые настраиваемые свойства, присвоенные вашей надстройкой, и обеспечить наличие в отклике значения свойства, чтобы в дальнейшем можно было применить логику фильтрации.

> [!IMPORTANT]
> В приведенном ниже примере замените `<app-guid>` идентификатором своей надстройки.

```rest
GET https://outlook.office.com/api/v2.0/Me/Events?$filter=SingleValueExtendedProperties/Any
  (ep: ep/PropertyId eq 'String {00020329-0000-0000-C000-000000000046}
  Name cecp-<app-guid>' and ep/Value ne null)
  &$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String
  {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')
```

Другие примеры использования REST для получения однозначного расширенного свойства, основанного на интерфейсе MAPI, см. в статье [Получение объекта singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0&preserve-view=true).

В приведенном ниже примере показано, как получить элемент и его настраиваемые свойства. В функции обратного вызова для метода `done` объект `item.SingleValueExtendedProperties` содержит список требуемых настраиваемых свойств.

> [!IMPORTANT]
> В приведенном ниже примере замените `<app-guid>` идентификатором своей надстройки.

```typescript
Office.context.mailbox.getCallbackTokenAsync(
    {
        isRest: true
    },
    function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded
            && asyncResult.value !== "") {
            let item_rest_id = Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.v2_0);
            let rest_url = Office.context.mailbox.restUrl +
                           "/v2.0/me/messages('" +
                           item_rest_id +
                           "')";
            rest_url += "?$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')";

            let auth_token = asyncResult.value;
            $.ajax(
                {
                    url: rest_url,
                    dataType: 'json',
                    headers:
                        {
                            "Authorization":"Bearer " + auth_token
                        }
                }
                ).done(
                    function (item) {
                        console.log(JSON.stringify(item));
                    }
                ).fail(
                    function (error) {
                        console.log(JSON.stringify(error));
                    }
                );
        } else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

## <a name="see-also"></a>См. также

- [Обзор свойств MAPI](/office/client-developer/outlook/mapi/mapi-property-overview)
- [Обзор свойств Outlook](/office/vba/outlook/How-to/Navigation/properties-overview)  
- [Вызов REST API Outlook из надстройки Outlook](use-rest-api.md)
- [Вызов веб-служб из надстройки Outlook](web-services.md)
- [Свойства и расширенные свойства в веб-службах Exchange](/exchange/client-developer/exchange-web-services/properties-and-extended-properties-in-ews-in-exchange)
- [Наборы свойств и формы ответа в веб-службах Exchange](/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange)