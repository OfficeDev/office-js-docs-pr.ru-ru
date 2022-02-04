---
title: Поддержка API JavaScript для Office для контентных надстроек и надстроек области задач в Office 2013
description: Используйте API Office JavaScript для создания области задач в Office 2013 г.
ms.date: 07/08/2021
ms.localizationpriority: medium
---

# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Поддержка API JavaScript для Office для контентных надстроек и надстроек области задач в Office 2013

[!include[information about the common API](../includes/alert-common-api-info.md)]

Вы можете использовать [API Office JavaScript](../reference/javascript-api-for-office.md) для создания области задач или надстройок контента для Office 2013 года. Объекты и методы, поддерживаемые контентными надстройками и надстройками области задач, можно сгруппировать указанным ниже образом.

1. **Общие объекты, Office надстройки.** Эти объекты [включают Office](/javascript/api/office), [Context](/javascript/api/office/office.context) и [AsyncResult](/javascript/api/office/office.asyncresult). Объект `Office` является корневым объектом API Office JavaScript. Объект `Context` представляет среду времени работы надстройки. Оба `Office` являются `Context` основными объектами для Office надстройки. Объект `AsyncResult` представляет результаты асинхронной операции, `getSelectedDataAsync` такие как данные, возвращенные методу, который считывая то, что пользователь выбрал в документе.

2. **Объект Document.** Большей частью API, доступной для контентных надстроек и надстроек области задач, можно воспользоваться с помощью методов, свойств и событий объекта [Document](/javascript/api/office/office.document). Контентная надстройка или надстройка области задач может использовать свойство [Office.context.document](/javascript/api/office/office.context#office-office-context-document-member) для доступа к объекту **Document** и с его помощью получать доступ к ключевым компонентам API для работы с данными в документах, например к объектам [Bindings](/javascript/api/office/office.bindings) и [CustomXmlParts](/javascript/api/office/office.customxmlparts) и методам [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)), [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1))и [getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)). Объект `Document` также предоставляет свойство mode [](/javascript/api/office/office.document#office-office-document-mode-member) для определения, является ли документ только для чтения или в режиме редактирования, url-адрес для получения URL-адреса текущего документа и доступа к объекту Параметры[](/javascript/api/office/office.settings).[](/javascript/api/office/office.document#office-office-document-url-member) Объект `Document` также поддерживает добавление обработчиков событий для события [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) , чтобы можно было определить, когда пользователь меняет свой выбор в документе.

   Надстройка `Document` содержимого или области задач может получить доступ к объекту только после загрузки среды DOM и среды запуска, как правило, в обработчике событий для [события Office.initialize](/javascript/api/office). Сведения о потоке событий при инициализации надстройки и проверке успешности загрузки модели DOM и среды выполнения см. в разделе [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).

3. **Объекты для работы с конкретными функциями.** Для работы с определенными функциями API используйте следующие объекты и методы.

    - Используйте методы объекта [Bindings](/javascript/api/office/office.bindings) для создания или получения привязок и методы и свойства объекта [Binding](/javascript/api/office/office.binding) для работы с данными.

    - Используйте объекты [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) и сопоставленные с ними объекты для создания пользовательских XML-частей в документах Word и управления ими.

    - Используйте объекты [File](/javascript/api/office/office.file) и [Slice](/javascript/api/office/office.slice) для создания копии всего документа, его разбивки на блоки или "фрагменты" и последующего считывания или передачи данных, содержащихся в этих фрагментах.

    - Используйте объект [Settings](/javascript/api/office/office.settings) для сохранения пользовательских данных, например настроек пользователей и состояния надстройки.

> [!IMPORTANT]
> Некоторые элементы API поддерживаются не всеми приложениями Office, в которых могут размещаться контентные надстройки и надстройки области задач. Чтобы определить, какие элементы поддерживаются, см. один из указанных ниже ресурсов.

Сводка о поддержке API Office JavaScript в Office клиентских приложений см. в Office [API JavaScript](understanding-the-javascript-api-for-office.md).

## <a name="read-and-write-to-an-active-selection-in-a-document-spreadsheet-or-presentation"></a>Чтение и написание активного выбора в документе, таблице или презентации

Вы можете считывать данные из текущего выделения пользователя в документе, электронной таблице или презентации, а также записывать их в это выделение. В зависимости Office приложения для надстройки можно указать тип структуры данных для чтения или записи в качестве параметра в [методах getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) и [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) объекта [Document](/javascript/api/office/office.document). Например, вы можете указать любой тип данных (текст, HTML, табличные данные или Office Open XML) для Word, текст и табличные данные для Excel, а также текст для PowerPoint и Project. Вы также можете создать обработчики событий для обнаружения изменений в выделении пользователя. В следующем примере получаются данные из выбора в виде текста с помощью `getSelectedDataAsync` метода.


```js
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        }
        else {
            write('Selected data: ' + asyncResult.value);
        }
    });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}

```

Дополнительные сведения и примеры см. в статье [Чтение и запись данных в активное выделение в документе или в электронной таблице](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).

## <a name="bind-to-a-region-in-a-document-or-spreadsheet"></a>Привязка к региону в документе или таблице

Вы можете использовать эти `getSelectedDataAsync` `setSelectedDataAsync` методы для чтения или записи текущего выбора пользователя в  документе, таблице или презентации. Тем не менее если вам необходимо обращаться к одной области документа в различных сеансах запущенной надстройки, не принуждая пользователя делать выбор, сначала потребуется создать привязку к этой области. Вы также сможете подписаться на события изменения данных и выделения только для этой привязанной области.

Привязку можно добавить с помощью методов [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)), [addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1)) или [addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) объекта [Bindings](/javascript/api/office/office.bindings). Они возвращают идентификатор, который можно использовать для доступа к данным в привязке или для подписки на события изменения данных или выделения в привязанной области.

Ниже приводится пример, который добавляет привязку к выбранному в настоящее время тексту в документе с помощью `Bindings.addFromSelectionAsync` метода.

```js
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Дополнительные сведения и примеры см. в статье [Привязка к областям в документе или электронной таблице](bind-to-regions-in-a-document-or-spreadsheet.md).

## <a name="get-entire-documents"></a>Получить все документы

Если надстройка области задач работает в PowerPoint или Word, то для получения презентации или документа целиком вы можете использовать методы [Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)), [File.getSliceAsync](/javascript/api/office/office.file#office-office-file-getsliceasync-member(1))и [File.closeAsync](/javascript/api/office/office.file#office-office-file-closeasync-member(1)).

При вызове `Document.getFileAsync` вы получите копию документа в [объекте File](/javascript/api/office/office.file) . Объект `File` предоставляет доступ к документу в "кусках", представленных как [объекты Slice](/javascript/api/office/office.slice) . При вызове `getFileAsync`можно указать тип файла (текстовый или сжатый формат Open Office XML) и размер срезов (до 4МБ). Чтобы получить доступ к содержимому `File` объекта, `File.getSliceAsync` необходимо позвонить, который возвращает необработанные данные в [свойстве Slice.data](/javascript/api/office/office.slice#office-office-slice-data-member) . Если вы выбрали сжатый формат, то получите данные файлов в виде массива байтов. Если вы передаете файл в веб-службу, перед отправкой можно преобразовать сжатые необработанные данные в строку с кодировкой Base64. Наконец, когда вы закончите получать срезы файла, используйте метод `File.closeAsync` , чтобы закрыть документ.

Дополнительные сведения см. в инструкции по [получению документа целиком из надстройки для PowerPoint или Word](../word/get-the-whole-document-from-an-add-in-for-word.md).

## <a name="read-and-write-custom-xml-parts-of-a-word-document"></a>Чтение и написание пользовательских XML-частей документа Word

С помощью формата файлов Open Office XML и элементов управления контентом вы можете добавлять пользовательские XML-части в документ Word и привязывать элементы в XML-частях к элементам управления контентом в этом документе. При открытии документа Word считывает и автоматически заполняет привязанные элементы управления контентом данными из пользовательских XML-частей. Кроме того, пользователи могут записывать данные в элементы управления контентом. Когда пользователь сохраняет документ, данные в элементах управления также сохраняются в привязанных XML-частях. Надстройки области задач для Word могут использовать свойство [Document.customXmlParts](/javascript/api/office/office.document#office-office-document-customxmlparts-member), а также объекты [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart)и [CustomXmlNode](/javascript/api/office/office.customxmlnode) для динамического считывания и записи данных в документ.

Пользовательские XML-части можно сопоставить с пространствами имен. Чтобы получить данные из пользовательских XML-частей в пространстве имен, используйте метод [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbynamespaceasync-member(1)).

Кроме того, вы можете использовать метод [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)), чтобы получить доступ к пользовательским XML-частям по их GUID. После этого можно получить XML-данные с помощью метода [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#office-office-customxmlpart-getxmlasync-member(1)).

Чтобы добавить новую настраиваемую часть XML в документ, `Document.customXmlParts` используйте свойство для получения пользовательских частей XML, которые находятся в документе, и позвоните [методу CustomXmlParts.addAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-addasync-member(1)) .

Подробные сведения о работе с пользовательскими XML-частями с помощью надстройки области задач см. в статье [Создание улучшенных надстроек для Word с помощью Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).

## <a name="persisting-add-in-settings"></a>Сохранение настроек надстроек

Надстройкам часто требуется сохранять пользовательские данные, например настройки пользователя или состояние надстройки, и запрашивать их при следующем запуске. Для этого можно использовать стандартные методы веб-программирования, например файлы cookie в браузере или веб-хранилище HTML 5. Кроме того, если надстройка работает в Excel, PowerPoint или Word, вы можете использовать методы объекта [Settings](/javascript/api/office/office.settings). Данные, созданные с `Settings` объектом, хранятся в таблице, презентации или документе, в который была вставлена и сохранена надстройка. Данные доступны только для создавшей их надстройки.

Чтобы избежать округлий на сервер, на котором хранится документ, данные, созданные с объектом, `Settings` управляются в памяти во время запуска. Ранее сохраненные данные настроек загружаются в память при инициализации надстройки, а внесенные в эти данные изменения сохраняются в документе только при вызове метода [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)). Для внутренних целей данные хранятся в сериализованном объекте JSON в виде пар "имя-значение". Вы можете использовать методы [get](/javascript/api/office/office.settings#office-office-settings-get-member(1)), [set](/javascript/api/office/office.settings#office-office-settings-set-member(1)) и [remove](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) объекта **Settings** для чтения, записи и удаления элементов из копии данных, хранящейся в памяти. Ниже приведена строка кода, позволяющая создать настройку с именем `themeColor` и задать для нее значение green.

```js
Office.context.document.settings.set('themeColor', 'green');
```

`set` `remove` Так как данные параметров, созданные или удаленные с помощью методов, действуют на копию данных в памяти, `saveAsync` необходимо вызвать для сохраняемого изменения параметров данных в документе, с который работает надстройка.

Дополнительные сведения о работе с пользовательскими данными `Settings` с использованием методов объекта см. в материале [Persisting add-in state and settings](persisting-add-in-state-and-settings.md).

## <a name="read-properties-of-a-project-document"></a>Чтение свойств документа проекта

Если надстройка области задач выполняется в Project, то она может считывать данные из некоторых полей, ресурсов и полей задач в активном проекте. Для этого используются методы и события объекта [ProjectDocument](/javascript/api/office/office.document), `Document` который расширяет объект, чтобы предоставить дополнительные Project-конкретные функции.

Примеры считывания данных Project см. в статье [Создание первой надстройки области задач для Project 2013 с использованием текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).

## <a name="permissions-model-and-governance"></a>Модель разрешений и управление

Ваша надстройка `Permissions` использует элемент в манифесте для запроса разрешения на доступ к требуемого уровня функциональности из API JavaScript Office JavaScript. Например, если надстройка требует доступа к документу для чтения и записи, `ReadWriteDocument` его манифест должен указываться как текстовое значение в элементе `Permissions` . Так как разрешения обеспечивают конфиденциальность и безопасность пользователей, рекомендуется запросить минимальный уровень разрешений, необходимый для работы надстройки. В примере ниже показано, как запросить разрешение **ReadDocument** в манифесте надстройки области задач.

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

Дополнительные сведения см. в дополнительных сведениях о запросе [разрешений на использование API в надстройки](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).

## <a name="see-also"></a>См. также

- [API JavaScript для Office](../reference/javascript-api-for-office.md)
- [Справочная схема по манифестам надстроек для Office](../develop/add-in-manifests.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](../testing/testing-and-troubleshooting.md)
