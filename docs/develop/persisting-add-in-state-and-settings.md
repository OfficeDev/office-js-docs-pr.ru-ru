---
title: Сохранение состояния и параметров надстройки
description: Узнайте, как сохранять данные Office веб-приложениях надстройки, работающих в среде без состояния управления браузером.
ms.date: 12/15/2021
ms.localizationpriority: medium
ms.openlocfilehash: dc99285aaaedfe5aa8385709fc5512f7cedfb2a6
ms.sourcegitcommit: a8bfb169b9e0b26d34a2839843e480a47ca098cc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2021
ms.locfileid: "61583800"
---
# <a name="persisting-add-in-state-and-settings"></a>Сохранение состояния и параметров надстройки

[!include[information about the common API](../includes/alert-common-api-info.md)]

Надстройки Office, по сути, представляют собой веб-приложения, которые выполняются в среде без сведений о состоянии элемента управления браузером. Вследствие этого надстройке может потребоваться сохранять данные для обеспечения непрерывности определенных операций или функций во время сеансов ее использования. Например, у надстройки могут быть настраиваемые параметры или другие значения, которые должны быть сохранены и повторно загружены при следующей инициализации, такие как выбранное пользователем представление или расположение по умолчанию. Это можно реализовать указанными ниже способами.

- Используйте членов API Office JavaScript, которые хранят данные как:
  - пар имя-значение в контейнере свойств, расположение которого зависит от типа надстройки;
  - пользовательского кода XML в документе.

- Использовать способы, предоставленные базовыми элементами управления браузером: cookie-файлы браузера или веб-хранилище HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) или [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).
    > [!NOTE]
    > Некоторые браузеры или параметры браузера пользователя могут блокировать методы хранения на основе браузера. Необходимо проверить доступность, как описано в [API служба хранилища веб служба хранилища.](https://developer.mozilla.org/docs/Web/API/Web_Storage_API/Using_the_Web_Storage_API)

В этой статье основное внимание уделяется использованию Office API JavaScript для сохраняемого состояния надстройки к текущему документу. Если необходимо сохранять состояние между документами, например отслеживание предпочтений пользователей в любых открытых документах, необходимо использовать другой подход. Например, можно использовать [SSO](sso-in-office-add-ins.md#using-the-sso-token-as-an-identity) для получения удостоверения пользователя, а затем сохранить идентификатор пользователя и его параметры в базе данных в Интернете.

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a>Сохраняемая надстройка и параметры с Office API JavaScript

API Office JavaScript предоставляет [объекты Параметры,](/javascript/api/office/office.settings) [RoamingSettings](/javascript/api/outlook/office.roamingsettings)и [CustomProperties](/javascript/api/outlook/office.customproperties) для сохранения состояния надстройки во всех сеансах, как описано в следующей таблице. Во всех случаях сохраненные значения параметров связаны с [Id](../reference/manifest/id.md) создавшей их надстройки.

|**Объект**|**Поддерживаемый тип надстроек**|**Расположение хранилища**|**Office поддержки приложений**|
|:-----|:-----|:-----|:-----|
|[Параметры](/javascript/api/office/office.settings)|Надстройки области задач и контентные надстройки|Документ, электронная таблица или презентация, с которыми работает надстройка. Параметры надстроек области задач и контентных надстроек доступны создавшей их надстройке в документе, в котором они сохранены.<br/><br/>**Внимание!** Не храните в объекте **Settings** пароли и другие конфиденциальные персональные данные. Сохраненные данные не видны пользователям, но содержатся документе, доступ к которому можно получить при прямом считывании. Необходимо ограничить использование надстройкой персональных данных и использовать для их хранения сервер, на котором эта надстройка размещена, как защищенный от пользователей ресурс.|Word, Excel или PowerPoint<br/><br/> **Примечание:** Надстройки области задач Project 2013 г. не  поддерживают API Параметры для хранения состояния или параметров надстройки. Однако для надстройок, работающих в Project (как и Office клиентских приложениях), можно использовать такие методы, как cookie-файлы браузера или веб-хранилище. Дополнительные сведения об этих методах см. в [Excel-add-in-JavaScript-PersistCustomSettings.](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings) |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Outlook|Почтовый ящик пользователя на сервере Exchange, на котором установлена надстройка. Поскольку эти параметры хранятся в почтовом ящике сервера пользователя, они могут "перемещаться" с пользователем и доступны надстройке, когда она запущена в контексте любого поддерживаемого Office клиентского приложения или браузера, доступ к почтовому ящику этого пользователя.<br/><br/> Параметры перемещения надстройки Outlook доступны только создавшей их надстройке и только в том почтовом ящике, в котором она установлена.|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Outlook|Элемент сообщения, встречи, запроса на собрание для которого была запущена надстройка. Пользовательские свойства элемента надстройки Outlook доступны только для создавшей их надстройки и только в элементе, в котором они сохранены.|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|Надстройки области задач|Документ, электронная таблица или презентация, с которыми работает надстройка. Параметры надстроек области задач доступны создавшей их надстройке в том документе, где они сохранены.<br/><br/>**Внимание!** Не храните пароли и другие конфиденциальные личные сведения в пользовательской части XML. Сохраненные данные не видны пользователям, но содержатся в документе, доступ к которому можно получить при прямом считывании формата файла. Необходимо ограничить использование надстройкой личных сведений и хранить их только на том сервере, где размещена эта надстройка, так как этот ресурс защищен от пользователей.|Word (с Office общего API JavaScript) Excel (с использованием Excel API JavaScript)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Данные параметров обрабатываются в памяти во время выполнения.

> [!NOTE]
> В следующих двух разделах рассматриваются параметры в контексте общего API JavaScript для Office. API JavaScript, Excel приложения, также предоставляет доступ к настраиваемой настройке. Интерфейсы API Excel и шаблоны программирования слегка отличаются. Дополнительные сведения см. в статье [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).

Внутренне данные в пакете свойств, доступ к нему с объектом , или объекты хранятся в качестве последовательного объекта Нотации объектов `Settings` `CustomProperties` JavaScript (JSON), который содержит пары имен `RoamingSettings` и значений. Имя (ключ) для каждого значения должно быть значением , и сохраненное значение может быть JavaScript , или , или , но `string` `string` не `number` `date` `object` **функции**.

Пример структуры контейнера свойств, содержащего три определенных **строковых** значения с именами `firstName`, `location` и `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

После сохранения контейнера свойств параметров во время предыдущего сеанса надстройки он может быть загружен при инициализации надстройки или в любое время после этого в течение текущего сеанса надстройки. Во время сеанса параметры управляются полностью в памяти с помощью , и методы объекта, который соответствует типу параметров, которые вы создаете `get` `set` `remove` **(Параметры,** **CustomProperties**, или **RoamingSettings**).

> [!IMPORTANT]
> Чтобы сохранить все дополнения, обновления или удаления, сделанные во время текущего сеанса надстройки в хранилище, необходимо вызвать метод соответствующего объекта, используемого для работы с такого рода `saveAsync` настройками. И методы работают только на копии свойств параметров в `get` `set` `remove` памяти. Если надстройка закрыта без вызова, все изменения, внесенные в параметры во время `saveAsync` сеанса, будут потеряны.

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Сохранение состояния надстройки и параметров документа для контентных надстроек и надстроек области задач

Чтобы сохранить состояние или пользовательские параметры в контентной надстройке или надстройке области задач в Word, Excel или PowerPoint, следует использовать объект [Settings](/javascript/api/office/office.settings) и его методы. Пакет свойств, созданный с методами объекта, доступен только экземпляру созданной надстройки содержимого или области задач и только из документа, в котором он `Settings` сохранен.

Объект автоматически загружается как часть объекта Document и доступен при активации области задач или `Settings` надстройки [](/javascript/api/office/office.document) контента. После мгновенного создания объекта можно получить доступ к объекту `Document` `Settings` с свойством [параметров](/javascript/api/office/office.document#settings) `Document` объекта. В течение всего срока службы сеанса можно просто использовать , и методы чтения, записи или удаления сохраняемых параметров и состояния надстройки из копии пакета свойств в `Settings.get` `Settings.set` `Settings.remove` памяти.

Поскольку методы "set" и "remove" работают только в копии контейнера свойств параметров, содержащейся в памяти, для сохранения новых или измененных параметров документа, с которым сопоставлена надстройка, необходимо вызвать метод [Settings.saveAsync](/javascript/api/office/office.settings#saveAsync_options__callback_).

### <a name="creating-or-updating-a-setting-value"></a>Создание или обновление значения параметра

Следующий пример кода демонстрирует использование метода [Settings.set](/javascript/api/office/office.settings#set_name__value_) для создания параметра с именем `'themeColor'`, имеющий значение  `'green'`. Первый параметр этого метода — это зависящий от регистра идентификатор  _name_ параметра, который следует определить или создать. Второй параметр — это _value_ параметра.

```js
Office.context.document.settings.set('themeColor', 'green');
```

 Создается параметр с указанным именем, если таковой еще не существует или обновляется значение, если параметр существует. Используйте `Settings.saveAsync` метод, чтобы сохранить новые или обновленные параметры документа.

### <a name="getting-the-value-of-a-setting"></a>Получение значения параметра

В следующем примере показано, как использовать метод [Settings.get](/javascript/api/office/office.settings#get_name_) для получения значения параметра "themeColor". Единственным параметром метода `get` является конфиденциальное _имя_ параметра.

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 Метод возвращает значение, которое было сохранено ранее для имени `get` _параметра,_ которое было передано. Если параметр не существует, метод возвращает **null**.

### <a name="removing-a-setting"></a>Удаление параметра

В следующем примере показано, как использовать метод [Settings.remove](/javascript/api/office/office.settings#remove_name_) для удаления параметра с именем "themeColor". Единственным параметром метода `remove` является конфиденциальное _имя_ параметра.

```js
Office.context.document.settings.remove('themeColor');
```

Если параметр не существует, ничего не произойдет. Используйте метод `Settings.saveAsync` для сохраняемого удаления параметра из документа.

### <a name="saving-your-settings"></a>Сохранение параметров

Чтобы сохранить любые добавления, изменения или удаления, внесенные надстройкой в копию контейнера свойств параметров, хранящуюся в памяти, во время текущего сеанса надстройки, необходимо вызвать метод [Settings.saveAsync](/javascript/api/office/office.settings#saveAsync_options__callback_) для их сохранения в документе. Единственным параметром метода является вызов, который является функцией вызова с `saveAsync` одним параметром. 

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

Анонимная функция передается в метод по мере выполнения параметра обратного вызова по `saveAsync` завершению  операции. Параметр _asyncResult_ от вызываемого вызова предоставляет доступ к объекту, который содержит `AsyncResult` состояние операции. В примере функция проверяет свойство, удалось ли операция сохранения или не удалось, а затем отображает результат на странице `AsyncResult.status` надстройки.

## <a name="how-to-save-custom-xml-to-the-document"></a>Сохранение пользовательского кода XML в документе

> [!NOTE]
> В этом разделе рассматриваются пользовательские части XML в контексте общего API JavaScript для Office, поддерживаемого в Word. API JavaScript, Excel приложения, также предоставляет доступ к настраиваемой части XML. Интерфейсы API Excel и шаблоны программирования слегка отличаются. Дополнительные сведения см. в статье [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).

Существует дополнительный параметр хранения, когда необходимо хранить сведения, которые превышают ограничения размера документа Параметры имеет структурированный характер. Вы можете сохранять пользовательскую разметку XML в надстройке области задач для Word (а также для Excel, но следует учитывать примечание в начале этого раздела). В Word можно использовать объект [CustomXmlPart](/javascript/api/office/office.customxmlpart) и его методы (еще раз, см. примечание для Excel выше). В приведенном ниже коде создается пользовательская часть XML, после чего в разделителях на странице отображается сначала ее ИД, а затем ее содержимое. Обратите внимание, что в строке XML должен быть указан атрибут `xmlns`.

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

Чтобы получить пользовательскую часть XML, используйте метод [getByIdAsync](/javascript/api/office/office.customxmlparts#getByIdAsync_id__options__callback_). Однако ИД — это GUID, генерируемый при создании части XML, поэтому его невозможно узнать во время написания кода. По этой причине при создании части XML рекомендуется сразу сохранить ее ИД в виде параметра с запоминающимся идентификатором. Ниже показано, как это сделать. (Но см. в более ранних разделах этой статьи сведения и передовую практику при работе с настраиваемые параметры.)

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

## <a name="how-to-save-settings-in-an-outlook-add-in"></a>Сохранение параметров в Outlook надстройки

Сведения о том, как сохранить параметры в Outlook надстройки, см. в статью Управление состоянием и [Outlook надстройки.](../outlook/manage-state-and-settings-outlook.md)

## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [Надстройки Outlook](../outlook/outlook-add-ins-overview.md)
- [Управление состоянием и настройками Outlook надстройки](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
