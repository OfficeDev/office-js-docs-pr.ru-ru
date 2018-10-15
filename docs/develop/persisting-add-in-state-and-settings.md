---
title: Сохранение состояния и параметров надстройки
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7ea35f00809fbe960155137c7cdae3f6dfd60b90
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945384"
---
# <a name="persisting-add-in-state-and-settings"></a>Сохранение состояния и параметров надстройки

Надстройки Office, по сути, представляют собой веб-приложения, которые выполняются в среде без сведений о состоянии элемента управления браузером. Вследствие этого надстройке может потребоваться сохранять данные для обеспечения непрерывности определенных операций или функций во время сеансов ее использования. Например, у надстройки могут быть настраиваемые параметры или другие значения, которые должны быть сохранены и повторно загружены при следующей инициализации, такие как выбранное пользователем представление или расположение по умолчанию. Это можно реализовать указанными ниже способами.

- Использовать элементы API JavaScript для Office, чтобы хранить данные в виде:
    -  пар имя-значение в контейнере свойств, расположение которого зависит от типа надстройки;
    -  пользовательского кода XML в документе.
    
- Использовать способы, предоставленные базовыми элементами управления браузером: cookie-файлы браузера или веб-хранилище HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) или [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).
    
Эта статья содержит сведения об использовании API JavaScript для сохранения состояния надстройки. Примеры использования cookie-файлов браузера и веб-хранилища см. в примере кода [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a>Сохранение состояния и параметров надстройки с помощью JavaScript API для Office

API JavaScript для Office предоставляет объекты [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js), [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) и [CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js) для сохранения состояния надстройки во время сеансов, как показано в следующей таблице. Во всех случаях сохраненные значения параметров связаны с [Id](https://docs.microsoft.com/javascript/office/manifest/id?view=office-js) создавшей их надстройки.

|**Объект**|**Поддерживаемый тип надстроек**|**Расположение хранилища**|**Поддержка ведущих приложений Office**|
|:-----|:-----|:-----|:-----|
|[Параметры](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js)|Надстройки области задач и контентные надстройки|Документ, электронная таблица или презентация, с которой работает надстройка. Параметры надстроек области задач и контентных надстроек доступны создавшей их надстройке в том документе, где они сохранены.<br/><br/>**Внимание!** Не храните в объекте **Settings** пароли и другие конфиденциальные персональные данные. Сохраненные данные не видны пользователям, но содержатся документе, доступ к которому можно получить при прямом считывании. Необходимо ограничить использование надстройкой персональных данных и использовать для их хранения сервер, на котором эта надстройка размещена, как защищенный от пользователей ресурс.|Word, Excel или PowerPoint<br/><br/> **Примечание.** Надстройки области задач для Project 2013 не поддерживают API **Settings** для хранения данных о состоянии или параметров. Однако для надстроек, работающих в Project (а также в других ведущих приложениях Office), можно использовать cookie-файлы браузера или веб-хранилище. Дополнительные сведения об этих технологиях см. в статье [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js)|Outlook|Почтовый ящик пользователя на сервере Exchange Server, где установлена надстройка. Так как параметры сохраняются на сервере почтового ящика пользователя, они могут "перемещаться" с пользователем и доступны надстройке при запуске в контексте любого поддерживаемого клиентского ведущего приложения или браузера, получающего доступ к почтовому ящику этого пользователя.<br/><br/> Параметры перемещения надстройки Outlook доступны только создавшей их надстройке и только в том почтовом ящике, в котором она установлена.|Outlook|
|[CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js)|Outlook|Элемент сообщения, встречи, запроса на собрание для которого была запущена надстройка. Пользовательские свойства элемента надстройки Outlook доступны только для создавшей их надстройки и только в элементе, в котором они сохранены.|Outlook|
|[CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js)|Надстройки области задач|Документ, электронная таблица или презентация, с которыми работает надстройка. Параметры надстроек области задач доступны создавшей их надстройке в том документе, где они сохранены.<br/><br/>**Внимание!** Не храните пароли и другие конфиденциальные личные сведения в пользовательской части XML. Сохраненные данные не видны пользователям, но содержатся в документе, доступ к которому можно получить при прямом считывании формата файла. Необходимо ограничить использование надстройкой личных сведений и хранить их только на том сервере, где размещена эта надстройка, так как этот ресурс защищен от пользователей.|Word (с использованием общего API JavaScript для Office), Excel (с использованием специального API JavaScript для Excel)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Данные параметров обрабатываются в памяти во время выполнения.

> [!NOTE]
> В следующих двух разделах рассматриваются параметры в контексте общего API JavaScript для Office. Специальный API JavaScript для Excel также предоставляет доступ к настраиваемым параметрам. Интерфейсы API Excel и шаблоны программирования слегка отличаются. Дополнительные сведения см. в статье [Excel SettingCollection](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection?view=office-js).

Для внутренних целей данные в контейнере свойств, доступных с помощью объектов**Settings**,**CustomProperties** или **RoamingSettings**, сохраняются в качестве сериализованного объекта JSON, содержащего пары "имя-значение". Имя (ключ) для каждого значения должно быть**string** и значение, сохраненное в свойстве, может быть JavaScript **string**,**number**, **date** или **object**, но не должно быть**function**.

Пример структуры контейнера свойств, содержащего три определенных значения  **string** с именами `firstName`,  `location` и `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

После сохранения контейнера свойств параметров во время предыдущего сеанса надстройки, он может быть загружен при инициализации надстройки или в любое время после этого в течение текущего сеанса приложения. Во время сеанса параметры изменяются только в памяти с помощью методов объекта **get**,**set** и **remove**, соответствующего типу создаваемых параметров ( **Settings**,**CustomProperties** или **RoamingSettings**). 


> [!IMPORTANT]
> Чтобы операции добавления, обновления и удаления, выполненные в текущем сеансе надстройки, не были отменены, необходимо вызвать метод **saveAsync** соответствующего объекта, используемого для работы с заданным типом параметров. Методы **get**, **set** и **remove** работают только в копии контейнера свойств параметров, содержащейся в памяти. Если закрыть надстройку, не вызывая метод **saveAsync**, то все изменения, внесенные в параметры во время сеанса, будут потеряны. 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Сохранение состояния надстройки и параметров документа для надстроек области задач и контентных надстроек


Чтобы сохранить состояние или пользовательские параметры в контентной надстройке или надстройке области задач в Word, Excel или PowerPoint, следует использовать объект [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) и его методы. Контейнер свойств, созданный с помощью методов объекта **Settings**, доступен только тому экземпляру контентной надстройки или надстройки области задач, который создал этот контейнер, и только в том документе, где он сохранен.

Объект**Settings** автоматически загружается как часть объекта [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) и доступен при активации надстройки области задач или контентной надстройки. После создания экземпляра объекта **Document**, вы можете получить доступ к объекту **Settings** с помощью свойства [settings](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#settings) объекта **Document**. Во время действия сеанса можно использовать методы**Settings.get**,**Settings.set** и **Settings.remove** для чтения, записи или удаления сохраненных параметров и состояния надстройки из копии контейнера свойств, содержащейся в памяти.

Поскольку методы "set" и "remove" работают только в копии контейнера свойств параметров, содержащейся в памяти, для сохранения новых или измененных параметров документа, с которым сопоставлена надстройка, необходимо вызвать метод [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-).


### <a name="creating-or-updating-a-setting-value"></a>Создание или обновление значения параметра

Следующий пример кода демонстрирует использование метода [Settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#set-name--value-) для создания параметра с именем `'themeColor'`, имеющий значение  `'green'`. Первый параметр этого метода — это зависящий от регистра идентификатор  _name_ параметра, который следует определить или создать. Второй параметр — это _value_ параметра.


```js
Office.context.document.settings.set('themeColor', 'green');
```

 Создается параметр с указанным именем, если таковой еще не существует или обновляется значение, если параметр существует. Используйте метод **Settings.saveAsync** для сохранения новых или обновления существующих параметров документа.


### <a name="getting-the-value-of-a-setting"></a>Получение значения параметра

В следующем примере показано, как использовать метод [Settings.get](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#get-name-) для получения значения параметра "themeColor". Единственным параметром метода **get** является зависящий от регистра параметр _name_.


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 Метод **get** возвращает значение, которое было ранее сохранено для переданного параметра _name_. Если параметр не существует, метод возвращает **null**.


### <a name="removing-a-setting"></a>Удаление параметра

В следующем примере показано, как использовать метод [Settings.remove](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#remove-name-) для удаления параметра с именем "themeColor". Единственным параметром метода **remove** является зависящий от регистра параметр _name_.


```js
Office.context.document.settings.remove('themeColor');
```

Если параметр не существует, ничего не произойдет. Используйте метод**Settings.saveAsync** чтобы предотвратить удаление указанного параметра в документе.


### <a name="saving-your-settings"></a>Сохранение параметров

Чтобы сохранить любые добавления, изменения или удаления, внесенные надстройкой в копию контейнера свойств параметров, хранящуюся в памяти, во время текущего сеанса надстройки, необходимо вызвать метод [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-) для их сохранения в документе. Единственный параметр метода **saveAsync** — это _callback_, представляющий собой функцию обратного вызова с одним параметром. 


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

Анонимная функция, переданная в метод**saveAsync** в качестве параметра _callback_, выполняется после завершения операции. Параметр обратного вызова _asyncResult_ предоставляет доступ к объекту **AsyncResult**, содержащему сведения о состоянии операции. В этом примере функция проверяет свойство  **AsyncResult.status** для проверки успешного или неудачного выполнения операции с последующим отображением результата на странице надстройки.

## <a name="how-to-save-custom-xml-to-the-document"></a>Сохранение пользовательского кода XML в документе

> [!NOTE]
> В этом разделе рассматриваются пользовательские части XML в контексте общего API JavaScript для Office, поддерживаемого в Word. Специальный API JavaScript для Excel также предоставляет доступ к пользовательским частям XML. Интерфейсы API Excel и шаблоны программирования слегка отличаются. Дополнительные сведения см. в статье [Excel CustomXmlPart](https://docs.microsoft.com/javascript/api/excel/excel.customxmlpart?view=office-js).

Существует параметр дополнительного хранилища, когда вам необходимо сохранить сведения, превышающие размер ограничения параметров документа, или содержащие структурированный символ. Настраиваемую XML-разметку задач области надстройки можно сохранить для Word (а для Excel, см. примечание в верхней части этого раздела). В Word, используйте объект [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) и его методы (еще раз, см. примечание выше для Excel). Следующий код создает настраиваемую пользовательскую часть XML и отображает её идентификатор, а затем содержимое элементов DIV на странице. Обратите внимание, что должен быть атрибут`xmlns` в строке XML.

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);                    
                }
            );
        }
    );
}
```

Чтобы получить пользовательскую часть XML, используйте метод [getByIdAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#getbyidasync-id--options--callback-). Однако ИД — это GUID, генерируемый при создании части XML, поэтому его невозможно узнать во время написания кода. По этой причине при создании части XML рекомендуется сразу сохранить ее ИД в виде параметра с запоминающимся идентификатором. Ниже показано, как это сделать. В предыдущих разделах этой статьи вы найдете подробные сведения и рекомендации по работе с настраиваемыми параметрами.

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
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID'));
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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Сохранение параметров в почтовом ящике пользователя для надстроек Outlook в качестве параметров перемещения


Надстройка Outlook может использовать [объект RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) для сохранения данных состояния надстройки и данных настроек, характерных для почтового ящика пользователя. Эти данные доступны только этой надстройке Outlook от имени пользователя, выполняющего надстройку. Данные хранятся в почтовом ящике сервера Exchange пользователя и доступны, когда этот пользователь входит в свою учетную запись и запускает надстройку Outlook.


### <a name="loading-roaming-settings"></a>Загрузка параметров перемещения


Надстройка Outlook обычно загружает параметры перемещения в обработчик событий [Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js). В следующем примере кода JavaScript показано, как выполняется загрузка существующих параметров перемещения.


```js
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}

```


### <a name="creating-or-assigning-a-roaming-setting"></a>Создание или назначение параметра перемещения


Развивая предыдущий пример, следующая функция  `setAppSetting`, показывает, как использовать метод [RoamingSettings.set](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#set-name--value-) для определения или обновления заданного параметра `cookie` с указанием сегодняшнего числа. Затем он позволяет заново сохранить все параметры перемещения на сервере Exchange при помощи метода [RoamingSettings.saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#saveasync-callback-).


```js
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

Метод **saveAsync** сохраняет асинхронно параметры роуминга и использует  необязательную функцию обратного вызова. Этот пример кода передает функцию обратного вызова с именем `saveMyAppSettingsCallback` метода **saveAsync**. Когда параметр асинхронного вызова_asyncResult_ `saveMyAppSettingsCallback` возвращается, функция предоставляет доступ к объекту [AsyncResult](https://docs.microsoft.com/javascript/api/outlook?view=office-js), который можно использовать для определения успешного или неудачного выполнения операции свойства **AsyncResult.status**


### <a name="removing-a-roaming-setting"></a>Удаление параметра перемещения


Предыдущие примеры дополняет следующая функция  `removeAppSetting`, демонстрирующая применение метода [RoamingSettings.remove](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#remove-name-) для удаления параметра `cookie` и повторного сохранения всех параметров перемещения на сервере Exchange.


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Сохранение параметров для каждого элемента надстройки Outlook в качестве пользовательских свойств


Пользовательские свойства позволяют надстройке Outlook сохранять сведения об элементе, который она использует. Например, если в надстройке Outlook создается встреча на основе приглашения на собрание в сообщении, с помощью пользовательских свойств можно сохранить сведения о факте создания собрания. Это гарантирует, что надстройка не предложит создать встречу еще раз при повторном открытии сообщения.

Перед использованием пользовательских свойств для определенного сообщения, встречи или элемента приглашения на собрание, необходимо загрузить свойства в память путем вызова метода [loadCustomPropertiesAsync](https://docs.microsoft.com/javascript/api/outlook/office.mailbox?view=office-js) объекта **Item**. Если какие-либо пользовательские свойства уже заданы для текущего элемента, на этом этапе они загружаются с сервера Exchange. После загрузки свойств можно использовать методы [set](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#set-name--value-) и [get](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) объекта **CustomProperties** для добавления, обновления и получения свойств в памяти. Чтобы сохранить любые изменения, внесенные в пользовательские свойства элемента, необходимо использовать метод [saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#saveasync-callback--asynccontext-) для сохранения изменений в элементе на сервере Exchange.


### <a name="custom-properties-example"></a>Пример пользовательских свойств

В следующем примере демонстрируется упрощенный набор функций для надстройки Outlook, применяющей пользовательские свойства. Этот пример можно использовать в качестве отправной точки для работы с такой надстройкой Outlook. 

Надстройка Outlook, использующая эти функции, получает любые пользовательские свойства, вызывая метод**get** для переменной `_customProps`, как показано в приведенном ниже примере.




```js
var property = _customProps.get("propertyName");
```

Этот пример включает следующие функции:



|**Имя функции**|**Описание**|
|:-----|:-----|
| `Office.initialize`|Инициализирует надстройку и загружает пользовательские свойства текущего элемента с сервера Exchange.|
| `customPropsCallback`|Получает пользовательские свойства, возвращенные сервером Exchange, и сохраняет их для дальнейшего использования.|
| `updateProperty`|Задает или обновляет определенное свойство, а затем сохраняет изменение на сервер Exchange.|
| `removeProperty`|Удаляет определенное свойство и сохраняет факт удаления на сервере Exchange.|
| `saveCallback`|Обратный вызов метода **saveAsync** в функциях `updateProperty` и `removeProperty`.|



```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change 
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal 
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method. 
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```


## <a name="see-also"></a>См. также

- [Общие сведения об интерфейсе API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
