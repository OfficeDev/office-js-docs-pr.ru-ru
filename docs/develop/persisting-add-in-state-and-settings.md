---
title: Сохранение состояния и параметров надстройки
description: Сведения о том, как хранить данные в веб-приложениях надстройки Office, работающих в среде без сохранения состояния элемента управления браузера.
ms.date: 05/08/2020
localization_priority: Normal
ms.openlocfilehash: 81f149bdff540b236252a02a0c368799a11fed10
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609403"
---
# <a name="persisting-add-in-state-and-settings"></a>Сохранение состояния и параметров надстройки

[!include[information about the common API](../includes/alert-common-api-info.md)]

Надстройки Office, по сути, представляют собой веб-приложения, которые выполняются в среде без сведений о состоянии элемента управления браузером. Вследствие этого надстройке может потребоваться сохранять данные для обеспечения непрерывности определенных операций или функций во время сеансов ее использования. Например, у надстройки могут быть настраиваемые параметры или другие значения, которые должны быть сохранены и повторно загружены при следующей инициализации, такие как выбранное пользователем представление или расположение по умолчанию. Это можно реализовать указанными ниже способами.

- Используйте элементы API JavaScript для Office, которые хранят данные, как один из следующих:
    -  пар имя-значение в контейнере свойств, расположение которого зависит от типа надстройки;
    -  пользовательского кода XML в документе.

- Использовать способы, предоставленные базовыми элементами управления браузером: cookie-файлы браузера или веб-хранилище HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) или [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).

В этой статье рассказывается, как использовать API JavaScript для Office для сохранения состояния надстройки. Примеры использования файлов cookie браузера и веб-хранилища приведены в статье [Excel-Add-in-JavaScript-персисткустомсеттингс](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a>Сохранение состояния и параметров надстройки с помощью API JavaScript для Office

API JavaScript для Office предоставляет объекты [параметров](/javascript/api/office/office.settings), [roamingSettings](/javascript/api/outlook/office.roamingsettings)и [CustomProperties](/javascript/api/outlook/office.customproperties) для сохранения состояния надстройки во всех сеансах, как описано в следующей таблице. Во всех случаях сохраненные значения параметров связаны с [Id](../reference/manifest/id.md) создавшей их надстройки.

|**Объект**|**Поддерживаемый тип надстроек**|**Расположение хранилища**|**Поддержка ведущих приложений Office**|
|:-----|:-----|:-----|:-----|
|[Параметры](/javascript/api/office/office.settings)|Надстройки области задач и контентные надстройки|Документ, электронная таблица или презентация, с которыми работает надстройка. Параметры надстроек области задач и контентных надстроек доступны создавшей их надстройке в документе, в котором они сохранены.<br/><br/>**Внимание!** Не храните в объекте **Settings** пароли и другие конфиденциальные персональные данные. Сохраненные данные не видны пользователям, но содержатся документе, доступ к которому можно получить при прямом считывании. Необходимо ограничить использование надстройкой персональных данных и использовать для их хранения сервер, на котором эта надстройка размещена, как защищенный от пользователей ресурс.|Word, Excel или PowerPoint<br/><br/> **Примечание.** Надстройки области задач для Project 2013 не поддерживают API **Settings** для хранения данных о состоянии или параметров. Однако для надстроек, работающих в Project (а также в других ведущих приложениях Office), можно использовать cookie-файлы браузера или веб-хранилище. Дополнительные сведения об этих технологиях см. в статье [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Outlook|Почтовый ящик пользователя на сервере Exchange, на котором установлена надстройка. Поскольку параметры сохраняются на сервере почтового ящика пользователя, они могут "перемещаться" с пользователем и доступны надстройке при запуске в контексте любого поддерживаемого клиентского ведущего приложения или браузера с получением доступа к почтовому ящику нужного пользователя.<br/><br/> Параметры перемещения надстройки Outlook доступны только создавшей их надстройке и только в том почтовом ящике, в котором она установлена.|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Outlook|Элемент сообщения, встречи, запроса на собрание для которого была запущена надстройка. Пользовательские свойства элемента надстройки Outlook доступны только для создавшей их надстройки и только в элементе, в котором они сохранены.|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|Надстройки области задач|Документ, электронная таблица или презентация, с которыми работает надстройка. Параметры надстроек области задач доступны создавшей их надстройке в том документе, где они сохранены.<br/><br/>**Внимание!** Не храните пароли и другие конфиденциальные личные сведения в пользовательской части XML. Сохраненные данные не видны пользователям, но содержатся в документе, доступ к которому можно получить при прямом считывании формата файла. Необходимо ограничить использование надстройкой личных сведений и хранить их только на том сервере, где размещена эта надстройка, так как этот ресурс защищен от пользователей.|Word (с использованием общего API JavaScript для Office), Excel (с использованием специального API JavaScript для Excel)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Данные параметров обрабатываются в памяти во время выполнения.

> [!NOTE]
> В следующих двух разделах рассматриваются параметры в контексте общего API JavaScript для Office. Специальный API JavaScript для Excel также предоставляет доступ к настраиваемым параметрам. Интерфейсы API Excel и шаблоны программирования слегка отличаются. Дополнительные сведения см. в статье [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).

Внутренние данные в контейнере свойств, доступ к которым осуществляется с `Settings` помощью `CustomProperties` объектов, или `RoamingSettings` объектов, хранятся как сериализованный объект нотации объектов JavaScript (JSON), содержащий пары "имя-значение". Имя (ключ) для каждого значения должно иметь значение `string` , а хранимое значение может быть JavaScript,, `string` `number` `date` или `object` , но не **функцией**.

Пример структуры контейнера свойств, содержащего три определенных **строковых** значения с именами `firstName`, `location` и `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

После сохранения контейнера свойств параметров во время предыдущего сеанса надстройки он может быть загружен при инициализации надстройки или в любое время после этого в течение текущего сеанса надстройки. Во время сеанса параметры полностью управляются в памяти с помощью `get` `set` методов, и `remove` объекта, соответствующего типу создаваемых параметров (**Settings**, **CustomProperties**или **roamingSettings**).


> [!IMPORTANT]
> Для сохранения добавлений, обновлений или удалений, внесенных в текущем сеансе надстройки, в место хранения необходимо вызвать `saveAsync` метод соответствующего объекта, который используется для работы с этими параметрами. `get`Методы, `set` и, Кроме того, `remove` работают только в копии контейнера свойств параметров, нашедшегося в памяти. Если ваша надстройка закрывается без вызова `saveAsync` , любые изменения, внесенные в параметры во время этого сеанса, будут потеряны.


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Сохранение состояния надстройки и параметров документа для контентных надстроек и надстроек области задач


Чтобы сохранить состояние или пользовательские параметры в контентной надстройке или надстройке области задач в Word, Excel или PowerPoint, следует использовать объект [Settings](/javascript/api/office/office.settings) и его методы. Контейнер свойств, созданный с помощью методов объекта, `Settings` доступен только для экземпляра созданной и созданной надстройкой области задач и только из документа, в котором она сохранена.

`Settings`Объект автоматически загружается как часть объекта [Document](/javascript/api/office/office.document) и становится доступным при активации надстройки области задач или контентной надстройки. После `Document` создания экземпляра объекта можно получить доступ к `Settings` объекту с помощью свойства [Settings](/javascript/api/office/office.document#settings) `Document` объекта. Во время существования сеанса можно просто использовать `Settings.get` `Settings.set` методы, и и `Settings.remove` для чтения, записи или удаления сохраненных параметров и состояния надстройки из копии контейнера свойств в памяти.

Поскольку методы "set" и "remove" работают только в копии контейнера свойств параметров, содержащейся в памяти, для сохранения новых или измененных параметров документа, с которым сопоставлена надстройка, необходимо вызвать метод [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-).


### <a name="creating-or-updating-a-setting-value"></a>Создание или обновление значения параметра

Следующий пример кода демонстрирует использование метода [Settings.set](/javascript/api/office/office.settings#set-name--value-) для создания параметра с именем `'themeColor'`, имеющий значение  `'green'`. Первый параметр этого метода — это зависящий от регистра идентификатор  _name_ параметра, который следует определить или создать. Второй параметр — это _value_ параметра.


```js
Office.context.document.settings.set('themeColor', 'green');
```

 Создается параметр с указанным именем, если таковой еще не существует или обновляется значение, если параметр существует. Используйте `Settings.saveAsync` метод для сохранения новых или обновленных параметров в документе.


### <a name="getting-the-value-of-a-setting"></a>Получение значения параметра

В следующем примере показано, как использовать метод [Settings.get](/javascript/api/office/office.settings#get-name-) для получения значения параметра "themeColor". Единственный параметр `get` метода — это _имя_ параметра с учетом регистра.


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 `get`Метод возвращает значение, сохраненное ранее для _имени_ параметра, которое было передано. Если параметр не существует, метод возвращает **null**.


### <a name="removing-a-setting"></a>Удаление параметра

В следующем примере показано, как использовать метод [Settings.remove](/javascript/api/office/office.settings#remove-name-) для удаления параметра с именем "themeColor". Единственный параметр `remove` метода — это _имя_ параметра с учетом регистра.


```js
Office.context.document.settings.remove('themeColor');
```

Если параметр не существует, ничего не произойдет. Используйте `Settings.saveAsync` метод для сохранения удаления параметра из документа.


### <a name="saving-your-settings"></a>Сохранение параметров

Чтобы сохранить любые добавления, изменения или удаления, внесенные надстройкой в копию контейнера свойств параметров, хранящуюся в памяти, во время текущего сеанса надстройки, необходимо вызвать метод [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) для их сохранения в документе. Единственный параметр `saveAsync` метода — _обратный вызов_, который является функцией обратного вызова с одним параметром. 


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

Анонимная функция, передаваемая в `saveAsync` метод в качестве параметра _callback_ , выполняется по завершении операции. Параметр _asyncResult_ обратного вызова предоставляет доступ к `AsyncResult` объекту, который содержит состояние операции. В этом примере функция проверяет `AsyncResult.status` свойство, чтобы убедиться, что операция сохранения выполнена успешно или не выполнена, а затем отображает результат на странице надстройки.

## <a name="how-to-save-custom-xml-to-the-document"></a>Сохранение пользовательского кода XML в документе

> [!NOTE]
> В этом разделе рассматриваются пользовательские части XML в контексте общего API JavaScript для Office, поддерживаемого в Word. Специальный API JavaScript для Excel также предоставляет доступ к пользовательским частям XML. Интерфейсы API Excel и шаблоны программирования слегка отличаются. Дополнительные сведения см. в статье [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).

Если требуется сохранить данные, размер которых превышает ограничения для параметров документа, или структурированные данные, то используется дополнительный параметр хранения. Вы можете сохранять пользовательскую разметку XML в надстройке области задач для Word (а также для Excel, но следует учитывать примечание в начале этого раздела). В Word можно использовать объект [CustomXmlPart](/javascript/api/office/office.customxmlpart) и его методы (еще раз, см. примечание для Excel выше). В приведенном ниже коде создается пользовательская часть XML, после чего в разделителях на странице отображается сначала ее ИД, а затем ее содержимое. Обратите внимание, что в строке XML должен быть указан атрибут `xmlns`.

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

Чтобы получить пользовательскую часть XML, используйте метод [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-). Однако ИД — это GUID, генерируемый при создании части XML, поэтому его невозможно узнать во время написания кода. По этой причине при создании части XML рекомендуется сразу сохранить ее ИД в виде параметра с запоминающимся идентификатором. Ниже показано, как это сделать. В предыдущих разделах этой статьи вы найдете подробные сведения и рекомендации по работе с настраиваемыми параметрами.

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

Сведения о том, как сохранить параметры в надстройке Outlook, можно узнать в статье [Управление состоянием и настройками надстройки Outlook](../outlook/manage-state-and-settings-outlook.md).


## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [Надстройки Outlook](../outlook/outlook-add-ins-overview.md)
- [Управление состоянием и параметрами для надстройки Outlook](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
