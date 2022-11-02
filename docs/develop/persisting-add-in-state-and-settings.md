---
title: Сохранение состояния и параметров надстройки
description: Узнайте, как сохранять данные в веб-приложениях надстроек Office, работающих в среде без отслеживания состояния элемента управления браузера.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: e2018e5ecf419744257cdceac31b8b1688fa65ff
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810010"
---
# <a name="persist-add-in-state-and-settings"></a>Сохранение состояния и параметров надстройки

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location.
To do that, you can:

- Используйте члены API JavaScript для Office, которые хранят данные, как:
  - пар имя-значение в контейнере свойств, расположение которого зависит от типа надстройки;
  - пользовательского кода XML в документе.

- Использовать способы, предоставленные базовыми элементами управления браузером: cookie-файлы браузера или веб-хранилище HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) или [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).
    > [!NOTE]
    > Некоторые браузеры или параметры браузера пользователя могут блокировать методы хранения на основе браузера. Необходимо протестировать доступность, как описано в статье [Использование API веб-хранилища](https://developer.mozilla.org/docs/Web/API/Web_Storage_API/Using_the_Web_Storage_API).

В этой статье рассматривается использование API JavaScript для Office для сохранения состояния надстройки в текущем документе. Если необходимо сохранить состояние в документах, например отслеживать предпочтения пользователей в открытых документах, необходимо использовать другой подход. Например, вы можете использовать [единый вход](use-sso-to-get-office-signed-in-user-token.md) для получения удостоверения пользователя, а затем сохранить идентификатор пользователя и его параметры в оперативной базе данных.

## <a name="persist-add-in-state-and-settings-with-the-office-javascript-api"></a>Сохранение состояния и параметров надстройки с помощью API JavaScript для Office

API JavaScript для Office предоставляет объекты [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings) и [CustomProperties](/javascript/api/outlook/office.customproperties) для сохранения состояния надстройки в сеансах, как описано в следующей таблице. Во всех случаях сохраненные значения параметров связаны с [Id](/javascript/api/manifest/id) создавшей их надстройки.

|Объект|Поддерживаемый тип надстроек|Место хранения|Поддержка приложений Office|
|:-----|:-----|:-----|:-----|
|[Параметры](/javascript/api/office/office.settings)|-Содержимого<br>— область задач|Документ, электронная таблица или презентация, с которыми работает надстройка. Параметры надстроек области задач и контентных надстроек доступны создавшей их надстройке в документе, в котором они сохранены.<br/><br/>**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|-Слово<br>-Excel<br>-Powerpoint<br/><br/> **Примечание:** Надстройки области задач для Project 2013 не поддерживают API **параметров** для хранения состояния или параметров надстройки. Однако для надстроек, работающих в Project (а также в других клиентских приложениях Office), можно использовать такие методы, как файлы cookie браузера или веб-хранилище. Дополнительные сведения об этих методах см. [в разделе Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|mail;|Почтовый ящик пользователя на сервере Exchange, на котором установлена надстройка. Так как эти параметры хранятся в почтовом ящике сервера пользователя, они могут перемещаться вместе с пользователем и доступны надстройке, когда она выполняется в контексте любого поддерживаемого клиентского приложения Office или браузера, обращающейся к почтовому ящику этого пользователя.<br/><br/> Параметры перемещения надстройки Outlook доступны только создавшей их надстройке и только в том почтовом ящике, в котором она установлена.|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|mail;|The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|Надстройки области задач|The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.<br/><br/>**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|— Word (с использованием общего API JavaScript для Office)<br>— Excel (с использованием API JavaScript для конкретного приложения Для Excel)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Данные параметров обрабатываются в памяти во время выполнения.

> [!NOTE]
> В следующих двух разделах рассматриваются параметры в контексте общего API JavaScript для Office. Api JavaScript для Конкретного приложения Excel также предоставляет доступ к пользовательским параметрам. Интерфейсы API Excel и шаблоны программирования слегка отличаются. Дополнительные сведения см. в статье [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).

Внутренние данные в контейнере свойств, доступ к которым осуществляется с помощью `Settings`объектов , `CustomProperties`или `RoamingSettings` , хранятся в виде сериализованного объекта Нотации объектов JavaScript (JSON), содержащего пары "имя-значение". Имя (ключ) для каждого значения должно быть `string`, а хранимое значение может быть JavaScript `string`, `number`, `date`, или `object`, но не **функцией**.

Пример структуры контейнера свойств, содержащего три определенных **строковых** значения с именами `firstName`, `location` и `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

После сохранения контейнера свойств параметров во время предыдущего сеанса надстройки он может быть загружен при инициализации надстройки или в любое время после этого в течение текущего сеанса надстройки. Во время сеанса управление параметрами осуществляется полностью в памяти с помощью `get`методов , `set`и `remove` объекта , соответствующего типу создаваемых параметров (**Settings**, **CustomProperties** или **RoamingSettings**).

> [!IMPORTANT]
> Чтобы сохранить все добавления, обновления или удаления, сделанные во время текущего сеанса надстройки, в хранилище, необходимо вызвать `saveAsync` метод соответствующего объекта, используемого для работы с параметрами такого типа. Методы `get`, `set`и `remove` работают только с копией в памяти контейнера свойств settings. Если надстройка закрыта без вызова `saveAsync`, все изменения, внесенные в параметры во время этого сеанса, будут потеряны.

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Сохранение состояния надстройки и параметров документа для надстроек области задач и контентных надстроек

Чтобы сохранить состояние или пользовательские параметры в контентной надстройке или надстройке области задач в Word, Excel или PowerPoint, следует использовать объект [Settings](/javascript/api/office/office.settings) и его методы. Контейнер свойств, созданный с помощью методов `Settings` объекта, доступен только экземпляру создавшего его содержимого или надстройки области задач и только из документа, в котором он сохранен.

Объект `Settings` автоматически загружается как часть объекта [Document](/javascript/api/office/office.document) и становится доступным при активации области задач или контентной надстройки. После создания экземпляра `Document` объекта можно получить доступ к объекту `Settings` с помощью свойства `Document` [settings](/javascript/api/office/office.document#office-office-document-settings-member) объекта . Во время существования сеанса можно просто использовать `Settings.get`методы , `Settings.set`и `Settings.remove` для чтения, записи или удаления сохраненных параметров и состояния надстройки из копии контейнера свойств в памяти.

Поскольку методы "set" и "remove" работают только в копии контейнера свойств параметров, содержащейся в памяти, для сохранения новых или измененных параметров документа, с которым сопоставлена надстройка, необходимо вызвать метод [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)).

### <a name="creating-or-updating-a-setting-value"></a>Создание или обновление значения параметра

The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#office-office-settings-set-member(1)) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.

```js
Office.context.document.settings.set('themeColor', 'green');
```

 Создается параметр с указанным именем, если таковой еще не существует или обновляется значение, если параметр существует. Используйте метод для `Settings.saveAsync` сохранения новых или обновленных параметров в документе.

### <a name="getting-the-value-of-a-setting"></a>Получение значения параметра

В следующем примере показано, как использовать метод [Settings.get](/javascript/api/office/office.settings#office-office-settings-get-member(1)) для получения значения параметра "themeColor". Единственным параметром `get` метода является _имя_ параметра с учетом регистра.

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 Метод `get` возвращает значение, которое было ранее сохранено для _переданного имени_ параметра. Если параметр не существует, метод возвращает **null**.

### <a name="removing-a-setting"></a>Удаление параметра

В следующем примере показано, как использовать метод [Settings.remove](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) для удаления параметра с именем "themeColor". Единственным параметром `remove` метода является _имя_ параметра с учетом регистра.

```js
Office.context.document.settings.remove('themeColor');
```

Если параметр не существует, ничего не произойдет. Используйте метод для `Settings.saveAsync` сохранения удаления параметра из документа.

### <a name="saving-your-settings"></a>Сохранение параметров

Чтобы сохранить любые добавления, изменения или удаления, внесенные надстройкой в копию контейнера свойств параметров, хранящуюся в памяти, во время текущего сеанса надстройки, необходимо вызвать метод [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) для их сохранения в документе. Единственным параметром `saveAsync` метода является _обратный вызов_, который является функцией обратного вызова с одним параметром.

```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

Анонимная функция, передаваемая в `saveAsync` метод в качестве параметра _обратного вызова_ , выполняется по завершении операции. Параметр _asyncResult_ обратного вызова предоставляет доступ к объекту `AsyncResult` , который содержит состояние операции. В этом примере функция проверяет `AsyncResult.status` свойство, чтобы узнать, была ли операция сохранения успешной или неудачной, а затем отображает результат на странице надстройки.

## <a name="how-to-save-custom-xml-to-the-document"></a>Сохранение пользовательского кода XML в документе

> [!NOTE]
> В этом разделе рассматриваются пользовательские части XML в контексте общего API JavaScript для Office, поддерживаемого в Word. API JavaScript для конкретного приложения Excel также предоставляет доступ к пользовательским XML-частям. Интерфейсы API Excel и шаблоны программирования слегка отличаются. Дополнительные сведения см. в статье [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).

Существует дополнительный вариант хранения, если необходимо хранить сведения, размер которых превышает ограничения параметров документа или имеет структурированный символ. Вы можете сохранять пользовательскую разметку XML в надстройке области задач для Word (а также для Excel, но следует учитывать примечание в начале этого раздела). В Word можно использовать объект [CustomXmlPart](/javascript/api/office/office.customxmlpart) и его методы (еще раз, см. примечание для Excel выше). В приведенном ниже коде создается пользовательская часть XML, после чего в разделителях на странице отображается сначала ее ИД, а затем ее содержимое. Обратите внимание, что в строке XML должен быть указан атрибут `xmlns`.

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.value.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

Чтобы получить пользовательскую часть XML, используйте метод [getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)). Однако ИД — это GUID, генерируемый при создании части XML, поэтому его невозможно узнать во время написания кода. По этой причине при создании части XML рекомендуется сразу сохранить ее ИД в виде параметра с запоминающимся идентификатором. Ниже показано, как это сделать. (Но дополнительные сведения и рекомендации по работе с пользовательскими параметрами см. в предыдущих разделах этой статьи.)

 ```js
function createCustomXmlPartAndStoreId() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            Office.context.document.settings.set('ReviewersID', asyncResult.id);
            Office.context.document.settings.saveAsync();
        }
    );
}
```

В приведенном ниже коде показано, как получить часть XML, сначала получив ее ИД из параметра.

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID');
    Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId,
        (asyncResult) => {
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

## <a name="how-to-save-settings-in-an-outlook-add-in"></a>Сохранение параметров в надстройке Outlook

Сведения о сохранении параметров в надстройке Outlook см. в разделе [Управление состоянием и параметрами для надстройки Outlook](../outlook/manage-state-and-settings-outlook.md).

## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [Надстройки Outlook](../outlook/outlook-add-ins-overview.md)
- [Управление состоянием и параметрами надстройки Outlook](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
