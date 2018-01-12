
# <a name="persisting-add-in-state-and-settings"></a>Сохранение состояния и параметров надстройки

Надстройки Office — это, по сути, веб-приложения, которые выполняются в среде без сведений о состоянии элемента управления браузером. Вследствие этого надстройке может потребоваться сохранять данные для обеспечения непрерывности определенных операций или функций во время сеансов использования надстройки. Например, у надстройки есть настраиваемые параметры или иные значения, которые должны быть сохранены и перезагружены при следующей инициализации, такие как выбранное пользователем представление или расположение по умолчанию.

Для этого воспользуйтесь


- Используйте члены JavaScript API для Office, хранящих данные, такие как пары "имя-значение" в контейнере свойств, расположение которого зависит от типа надстройки.
    
- Используйте способы, предоставленные базовыми элементами управления браузером: cookie-файлы браузера или веб-хранилище HTML5 ([localStorage](http://msdn.microsoft.com/en-us/library/cc848902%28v=vs.85%29.aspx) или [sessionStorage](http://msdn.microsoft.com/en-us/library/cc197020%28v=vs.85%29.aspx)).
    
Эта статья содержит сведения об использовании API JavaScript для сохранения состояния надстройки. Примеры использования cookie-файлов браузера и веб-хранилища см. в примере кода [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a>Сохранение состояния и параметров надстройки с помощью JavaScript API для Office


API JavaScript для Office предоставляет объекты [Settings](../../reference/shared/settings.md), [RoamingSettings](../../reference/outlook/RoamingSettings.md) и [CustomProperties](../../reference/outlook/CustomProperties.md) для сохранения состояния надстройки во время сеансов, как показано в следующей таблице. Во всех случаях сохраненные значения параметров связаны с [Id](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx) создавшей их надстройки.



|**Объект**|**Поддерживаемый тип надстроек**|**Расположение хранилища**|**Поддержка ведущих приложений Office**|
|:-----|:-----|:-----|:-----|
|[Параметры](../../reference/shared/settings.md)|Контентные приложения и приложения области задач|Документ, электронная таблица или презентация, с которыми работает надстройка. Параметры надстроек области задач и контентных надстроек доступны создавшей их надстройке в документе, в котором они сохранены. **Внимание!** Не храните в объекте **Settings** пароли и другие конфиденциальные личные сведения. Сохраненные данные не видны пользователям, но хранятся в документе и доступны с помощью прямого считывания. Необходимо ограничить использование надстройкой личных сведений и хранить их на сервере с надстройкой как защищенный от пользователей ресурс.|Word, Excel или PowerPoint **Примечание.** Надстройки области задач для Project 2013 не поддерживают API **Settings** для хранения данных о состоянии или параметров. Однако для надстроек, работающих в Project (а также в других ведущих приложениях Office), можно использовать файлы cookie браузера или веб-хранилище. Дополнительные сведения об этих технологиях см. в статье [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](../../reference/outlook/RoamingSettings.md)|Outlook|Почтовый ящик пользователя на сервере Exchange, на котором установлена надстройка.Поскольку параметры сохраняются на сервере почтового ящика пользователя, они могут "перемещаться" с пользователем и доступны надстройке при запуске в контексте любого поддерживаемого клиентского ведущего приложения или браузера с получением доступа к почтовому ящику нужного пользователя. Параметры перемещения надстройки Outlook доступны только для создавшей их надстройки и только в почтовом ящике, в котором она установлена.|Outlook|
|[CustomProperties](../../reference/outlook/CustomProperties.md)|Outlook|Элемент сообщения, встречи, запроса на собрание для которого была запущена надстройка. Пользовательские свойства элемента надстройки Outlook доступны только для создавшей их надстройки и только в элементе, в котором они сохранены.|Outlook|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Данные параметров обрабатываются в памяти во время выполнения.


Для внутренних целей данные в контейнере свойств, открываемые с помощью объектов  **Settings**,  **CustomProperties** или **RoamingSettings**, сохраняются в качестве сериализованного объекта JSON, содержащего пары "имя-значение". Имя (ключ) для каждого значения должно быть  **string** и значение, сохраненное в свойстве, может быть JavaScript **string**,  **number**,  **date** или **object**, но не должно быть  **function**.

Пример структуры контейнера свойств, содержащего три определенных значения  **string** с именами `firstName`,  `location` и `defaultView`.




```
{
"firstName":"Erik",
"location":"98052",
"defaultView":"basic"
}
```

После сохранения контейнера свойств параметров во время предыдущего сеанса надстройки он может быть загружен при инициализации надстройки или в любое время после этого в течение текущего сеанса приложения. Во время сеанса параметры изменяются только в памяти с помощью методов объекта  **get**,  **set** и **remove**, соответствующего типу создаваемых параметров ( **Settings**,  **CustomProperties** или **RoamingSettings**). 


 >**Важно!**  Чтобы сохранить в месте хранения любые добавления, обновления или удаления, выполненные во время текущего сеанса надстройки, необходимо вызвать метод  **saveAsync** соответствующего объекта, используемого для работы с заданным типом параметров. Методы **get**,  **set** и **remove** работают только в копии контейнера свойств параметров, содержащейся в памяти. Если надстройка закрывается без вызова **saveAsync**, любые изменения, внесенные в параметры во время сеанса, будут утеряны. 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Сохранение состояния надстройки и параметров документа для контентных надстроек и надстроек области задач


Чтобы сохранить состояние или пользовательские параметры в контентной надстройке или надстройке области задач в Word, Excel или PowerPoint, следует использовать объект [Settings](../../reference/shared/settings.md) и его методы. Контейнер свойств, созданный с помощью методов объекта **Settings**, доступен только тому экземпляру контентной надстройки или надстройки области задач, который создал этот контейнер, и только в том документе, где он сохранен.

Объект  **Settings** автоматически загружается как часть объекта [Document](../../reference/shared/document.md) и доступен при активации надстройки области задач или контентной надстройки. После создания экземпляра объекта **Document** вы можете получить доступ к объекту **Settings** с помощью свойства [settings](../../reference/shared/document.settings.md) объекта **Document**. Во время действия сеанса можно использовать методы  **Settings.get**,  **Settings.set** и **Settings.remove** для чтения, записи или удаления сохраненных параметров и состояния надстройки из копии контейнера свойств, содержащейся в памяти.

Поскольку методы "set" и "remove" работают только в копии контейнера свойств параметров, содержащейся в памяти, для сохранения новых или измененных параметров документа, с которым сопоставлена надстройка, необходимо вызвать метод [Settings.saveAsync](../../reference/shared/settings.saveasync.md).


### <a name="creating-or-updating-a-setting-value"></a>Создание или обновление значения параметра

Следующий пример кода демонстрирует использование метода [Settings.set](../../reference/shared/settings.set.md) для создания параметра с именем `'themeColor'`, имеющий значение  `'green'`. Первый параметр этого метода — это зависящий от регистра идентификатор  _name_ параметра, который следует определить или создать. Второй параметр — это _value_ параметра.


```
Office.context.document.settings.set('themeColor', 'green');
```

 Создается параметр с указанным именем, если таковой еще не существует или обновляется значение, если параметр существует. Используйте метод **Settings.saveAsync** для сохранения новых или обновления существующих параметров документа.


### <a name="getting-the-value-of-a-setting"></a>Получение значения параметра

В следующем примере показано, как использовать метод [Settings.get](../../reference/shared/settings.get.md) для получения значения параметра "themeColor". Единственным параметром метода **get** является зависящий от регистра параметр _name_.


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 Метод **get** возвращает значение, которое было ранее сохранено для переданного параметра _name_. Если параметр не существует, метод возвращает **null**.


### <a name="removing-a-setting"></a>Удаление параметра

В следующем примере показано, как использовать метод [Settings.remove](../../reference/shared/settings.removehandlerasync.md) для удаления параметра с именем "themeColor". Единственным параметром метода **remove** является зависящий от регистра параметр _name_.


```
Office.context.document.settings.remove('themeColor');
```

Если параметр не существует, ничего не произойдет. Используйте метод  **Settings.saveAsync** чтобы предотвратить удаление указанного параметра в документе.


### <a name="saving-your-settings"></a>Сохранение параметров

Чтобы сохранить любые добавления, изменения или удаления, внесенные надстройкой в копию контейнера свойств параметров, хранящуюся в памяти, во время текущего сеанса надстройки, необходимо вызвать метод [Settings.saveAsync](../../reference/shared/settings.saveasync.md) для их сохранения в документе. Единственный параметр метода **saveAsync** — это _callback_, представляющий собой функцию обратного вызова с одним параметром. 


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

Анонимная функция, переданная в метод  **saveAsync** в качестве параметра _callback_, выполняется после завершения операции. Параметр обратного вызова  _asyncResult_ предоставляет доступ к объекту **AsyncResult**, содержащему сведения о состоянии операции. В этом примере функция проверяет свойство  **AsyncResult.status** для проверки успешного или неудачного выполнения операции с последующим отображением результата на странице надстройки.


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Сохранение параметров в почтовом ящике пользователя для надстроек Outlook в качестве параметров перемещения


Надстройка Outlook может использовать объект [RoamingSettings](../../reference/outlook/RoamingSettings.md) для сохранения сведений о состоянии и параметров надстройки, относящихся к почтовому ящику пользователя. Эти данные доступны только этой надстройке Outlook, запущенной от имени пользователя. Эти же данные хранятся в почтовом ящике пользователя на сервере Exchange Server и становятся доступны, когда пользователь войдет в свою учетную запись и запустит надстройку Outlook.


### <a name="loading-roaming-settings"></a>Загрузка параметров перемещения


Надстройка Outlook обычно загружает параметры перемещения в обработчик событий [Office.initialize](../../reference/shared/office.initialize.md). В следующем примере кода JavaScript показано, как выполняется загрузка существующих параметров перемещения.


```
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


Развивая предыдущий пример, следующая функция  `setAppSetting`, показывает, как использовать метод [RoamingSettings.set](../../reference/outlook/RoamingSettings.md) для определения или обновления заданного параметра `cookie` с указанием сегодняшнего числа. Затем он позволяет заново сохранить все параметры перемещения на сервере Exchange при помощи метода [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md).


```
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

Метод  **saveAsync** сохраняет параметры перемещения асинхронно и получает дополнительную функцию обратного вызова. Данный пример кода передает функцию вызова `saveMyAppSettingsCallback` в метод **saveAsync**. После возврата асинхронного вызова параметр  _asyncResult_ функции `saveMyAppSettingsCallback` предоставляет доступ к объекту [AsyncResult](../../reference/outlook/simple-types.md), который можно использовать для определения успешного или неудачного выполнения операции при помощи свойства  **AsyncResult.status**.


### <a name="removing-a-roaming-setting"></a>Удаление параметра перемещения


Предыдущие примеры дополняет следующая функция  `removeAppSetting`, демонстрирующая применение метода [RoamingSettings.remove](../../reference/outlook/RoamingSettings.md) для удаления параметра `cookie` и повторного сохранения всех параметров перемещения на сервере Exchange.


```
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Сохранение параметров для каждого элемента надстройки Outlook в качестве пользовательских свойств


Пользовательские свойства позволяют надстройке Outlook сохранять сведения об элементе, который она использует. Например, если в надстройке Outlook создается встреча на основе приглашения на собрание в сообщении, с помощью пользовательских свойств можно сохранить сведения о факте создания собрания. Это гарантирует, что надстройка не предложит создать встречу еще раз при повторном открытии сообщения.

Перед использованием пользовательских свойств для определенного сообщения, встречи или элемента приглашения на собрание, необходимо загрузить свойства в память путем вызова метода [loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md) объекта **Item**. Если какие-либо пользовательские свойства уже заданы для текущего элемента, на этом этапе они загружаются с сервера Exchange. После загрузки свойств можно использовать методы [set](../../reference/outlook/CustomProperties.md) и [get](../../reference/outlook/RoamingSettings.md) объекта **CustomProperties** для добавления, обновления и получения свойств в памяти. Чтобы сохранить любые изменения, внесенные в пользовательские свойства элемента, необходимо использовать метод [saveAsync](../../reference/outlook/CustomProperties.md) для сохранения изменений в элементе на сервере Exchange.


### <a name="custom-properties-example"></a>Пример пользовательских свойств

В следующем примере демонстрируется упрощенный набор функций для надстройки Outlook, применяющей пользовательские свойства. Этот пример можно использовать в качестве отправной точки для работы с такой надстройкой Outlook. 

Надстройка Outlook, использующая эти функции, получает любые пользовательские свойства, вызывая метод  **get** для переменной `_customProps`, как показано в приведенном ниже примере.




```
var property = _customProps.get("propertyName");
```

Этот пример включает следующие функции:



|**Имя функции**|**Описание**|
|:-----|:-----|
| `Office.initialize`|Инициализирует надстройку и загружает пользовательские свойства текущего элемента с сервера Exchange.|
| `customPropsCallback`|Получает пользовательские свойства, возвращенные сервером Exchange, и сохраняет их для дальнейшего использования.|
| `updateProperty`|Задает или обновляет определенное свойство, а затем сохраняет изменение на сервер Exchange.|
| `removeProperty`|Удаляет определенное свойство и сохраняет факт удаления на сервере Exchange.|
| `saveCallback`|Обратный вызов для вызова метода  **saveAsync** в функциях `updateProperty` и `removeProperty`.|



```
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


## <a name="additional-resources"></a>Дополнительные ресурсы



- [Общие сведения об интерфейсе API JavaScript для Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Надстройки Outlook](../outlook/outlook-add-ins.md)
    
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
