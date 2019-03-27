---
title: Поддержка API JavaScript для Office для контентных надстроек и надстроек области задач в Office 2013
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: aeb75d2dbdcae5636bbe81820b9387bd840120c7
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870928"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Поддержка API JavaScript для Office для контентных надстроек и надстроек области задач в Office 2013


Вы можете использовать [API JavaScript для Office](/office/dev/add-ins/reference/javascript-api-for-office) для создания надстроек области задач или контентных надстроек для ведущих приложений Office 2013. Объекты и методы, поддерживаемые контентными надстройками и надстройками области задач, можно сгруппировать указанным ниже образом.


1. **Стандартные объекты, используемые совместно с другими надстройками Office.** Среди них объекты [Office](/javascript/api/office), [Context](/javascript/api/office/office.context) и [AsyncResult](/javascript/api/office/office.asyncresult). Объект **Office** — это корневой объект API JavaScript для Office. Объект **Context** представляет среду выполнения надстройки. **Office** и **Context** — базовые объекты для любой надстройки Office. Объект **AsyncResult** представляет результаты асинхронной операции, например данные, возвращенные в метод **getSelectedDataAsync**, который считывает сведения о том, какие элементы пользователь выделил в документе.

2. **The Document object.** The majority of the API available to content and task pane add-ins is exposed through the methods, properties, and events of the [Document](/javascript/api/office/office.document) object. A content or task pane add-in can use the [Office.context.document](/javascript/api/office/office.context#document) property to access the **Document** object, and through it, can access the key members of the API for working with data in documents, such as the [Bindings](/javascript/api/office/office.bindings) and [CustomXmlParts](/javascript/api/office/office.customxmlparts) objects, and the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-), [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-), and [getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) methods. The **Document** object also provides the [mode](/javascript/api/office/office.document#mode) property for determining whether a document is read-only or in edit mode, the [url](/javascript/api/office/office.document#url) property to get the URL of the current document, and access to the [Settings](/javascript/api/office/office.settings) object. The **Document** object also supports adding event handlers for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event, so you can detect when a user changes their selection in the document.

   Контентная надстройка или надстройка области задач может получить доступ к объекту **Document** только после загрузки модели DOM и среды выполнения (как правило, это происходит в обработчике события [Office.initialize](/javascript/api/office)). Сведения о потоке событий при инициализации надстройки и о том, как проверить успешность загрузки модели DOM и среды выполнения, см. в статье [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md).

3. **Объекты для работы с конкретными функциями.** Для работы с конкретными функциями API используйте указанные ниже объекты и методы.

    - Используйте методы объекта [Bindings](/javascript/api/office/office.bindings) для создания или получения привязок и методы и свойства объекта [Binding](/javascript/api/office/office.binding) для работы с данными.

    - Используйте объекты [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) и сопоставленные с ними объекты для создания пользовательских XML-частей в документах Word и управления ими.

    - Используйте объекты [File](/javascript/api/office/office.file) и [Slice](/javascript/api/office/office.slice) для создания копии всего документа, его разбивки на блоки или "фрагменты" и последующего считывания или передачи данных, содержащихся в этих фрагментах.

    - Используйте объект [Settings](/javascript/api/office/office.settings) для сохранения пользовательских данных, например настроек пользователей и состояния надстройки.


> [!IMPORTANT]
> Некоторые элементы API поддерживаются не всеми приложениями Office, в которых могут размещаться контентные надстройки и надстройки области задач. Чтобы определить, какие элементы поддерживаются, см. один из указанных ниже ресурсов.

Краткие сведения о поддержке API JavaScript для Office ведущими приложениями Office см. в [этой статье](understanding-the-javascript-api-for-office.md).


## <a name="reading-and-writing-to-an-active-selection"></a>Чтение и запись данных в активное выделение

Вы можете считывать данные из текущего выделения пользователя в документе, электронной таблице или презентации, а также записывать их в это выделение. В зависимости от ведущего приложения можно указать тип структуры данных, которая будет считана или записана, в качестве параметра методов [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) и [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) объекта [Document](/javascript/api/office/office.document). Например, вы можете указать любой тип данных (текст, HTML, табличные данные или Office Open XML) для Word, текст и табличные данные для Excel, а также текст для PowerPoint и Project. Вы также можете создать обработчики событий для обнаружения изменений в выделении пользователя. Ниже приведен пример получения данных из выделения в качестве текста с помощью метода **getSelectedDataAsync**.


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


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>Привязка к областям в документе или электронной таблице

Вы можете использовать методы **getSelectedDataAsync** и **setSelectedDataAsync** для чтения и записи данных в *текущее* выделение в документе, электронной таблице или презентации. Тем не менее если вам необходимо обращаться к одной области документа в различных сеансах запущенной надстройки, не принуждая пользователя делать выбор, сначала потребуется создать привязку к этой области. Вы также сможете подписаться на события изменения данных и выделения только для этой привязанной области.

Привязку можно добавить с помощью методов [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-), [addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-) или [addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) объекта [Bindings](/javascript/api/office/office.bindings). Они возвращают идентификатор, который можно использовать для доступа к данным в привязке или для подписки на события изменения данных или выделения в привязанной области.

Ниже представлен пример добавления привязки к текущему выбранному тексту в документе с помощью метода **Bindings.addFromSelectionAsync**.



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


## <a name="getting-entire-documents"></a>Получение документов целиком

Если надстройка области задач работает в PowerPoint или Word, то для получения презентации или документа целиком вы можете использовать методы [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-), [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)и [File.closeAsync](/javascript/api/office/office.file#closeasync-callback-).

При вызове метода **Document.getFileAsync** вы получаете копию документа в объекте [File](/javascript/api/office/office.file). Объект **File** обеспечивает доступ к документу в "блоках", представленных в качестве объектов [Slice](/javascript/api/office/office.slice). При вызове метода **getFileAsync** можно указать тип файла (текст или сжатый формат Open Office XML) и размер фрагментов (до 4 МБ). Для доступа к содержимому объекта **File** нужно вызвать метод **File.getSliceAsync**, который возвращает необработанные данные в свойстве [Slice.data](/javascript/api/office/office.slice#data). Если вы выбрали сжатый формат, то получите данные файлов в виде массива байтов. Если вы передаете файл в веб-службу, перед отправкой можно преобразовать сжатые необработанные данные в строку с кодировкой Base64. Получив фрагменты файла, закройте документ с помощью метода **File.closeAsync**.

Дополнительные сведения см. в инструкции по [получению документа целиком из надстройки для PowerPoint или Word](../word/get-the-whole-document-from-an-add-in-for-word.md).


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>Чтение и запись настраиваемых XML-частей документа Word

С помощью формата файлов Open Office XML и элементов управления контентом вы можете добавлять пользовательские XML-части в документ Word и привязывать элементы в XML-частях к элементам управления контентом в этом документе. При открытии документа Word считывает и автоматически заполняет привязанные элементы управления контентом данными из пользовательских XML-частей. Кроме того, пользователи могут записывать данные в элементы управления контентом. Когда пользователь сохраняет документ, данные в элементах управления также сохраняются в привязанных XML-частях. Надстройки области задач для Word могут использовать свойство [Document.customXmlParts](/javascript/api/office/office.document#customxmlparts), а также объекты [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart)и [CustomXmlNode](/javascript/api/office/office.customxmlnode) для динамического считывания и записи данных в документ.

Пользовательские XML-части можно сопоставить с пространствами имен. Чтобы получить данные из пользовательских XML-частей в пространстве имен, используйте метод [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getbynamespaceasync-ns--options--callback-).

Кроме того, вы можете использовать метод [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-), чтобы получить доступ к пользовательским XML-частям по их GUID. После этого можно получить XML-данные с помощью метода [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getxmlasync-options--callback-).

Чтобы добавить новую пользовательскую XML-часть в документ, с помощью свойства **Document.customXmlParts** получите пользовательские XML-части документа, а затем вызовите метод [CustomXmlParts.addAsync](/javascript/api/office/office.customxmlparts#addasync-xml--options--callback-).

Подробные сведения о работе с пользовательскими XML-частями с помощью надстройки области задач см. в статье [Создание улучшенных надстроек для Word с помощью Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).


## <a name="persisting-add-in-settings"></a>Сохранение настроек надстроек


Часто пользовательские настройки, состояние надстройки или другие данные требуется сохранять для работы с ними после перезапуска надстройки. Для этого можно использовать стандартные методы веб-программирования, например файлы cookie в браузере или веб-хранилище HTML 5. Кроме того, если надстройка работает в Excel, PowerPoint или Word, вы можете использовать методы объекта [Settings](/javascript/api/office/office.settings). Данные, созданные с помощью объекта **Settings**, хранятся в электронной таблице, презентации или документе, в который добавлена надстройка и с которым она сохранена. Эти данные доступны только для надстройки, которая их создала.

To avoid roundtrips to the server where the document is stored, data created with the **Settings** object is managed in memory at run time. Previously saved settings data is loaded into memory when the add-in is initialized, and changes to that data are only saved back to the document when you call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method. Для внутренних целей данные хранятся в сериализованном объекте JSON в виде пар "имя-значение". Вы можете использовать методы [get](/javascript/api/office/office.settings#get-name-), [set](/javascript/api/office/office.settings#set-name--value-) и [remove](/javascript/api/office/office.settings#remove-name-) объекта **Settings** для чтения, записи и удаления элементов из копии данных, хранящейся в памяти. Ниже приведена строка кода, позволяющая создать настройку с именем `themeColor` и задать для нее значение green.




```js
Office.context.document.settings.set('themeColor', 'green');
```

Так как созданные или удаленные с помощью методов **set** и **remove** данные настроек размещаются в хранящейся в памяти копии данных, для сохранения изменений, внесенных в данные настроек документа, с которым работает надстройка, необходимо вызвать метод **saveAsync**.

Дополнительные сведения о работе с пользовательскими данными с помощью методов объекта **Settings** см. в статье [Сохранение состояния и параметров надстройки](persisting-add-in-state-and-settings.md).


## <a name="reading-properties-of-a-project-document"></a>Чтение свойств документа проекта

Если надстройка области задач выполняется в Project, то она может считывать данные из некоторых полей, ресурсов и полей задач в активном проекте. Для этого используются методы и события объекта [ProjectDocument](/javascript/api/office/office.document), которые расширяют объект **Document** путем добавления функций работы с Project.

Примеры считывания данных Project см. в статье [Создание первой надстройки области задач для Project 2013 с использованием текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## <a name="permissions-model-and-governance"></a>Модель разрешений и управление

Ваша надстройка использует элемент **Permissions** в своем манифесте для запроса разрешения на доступ к уровню функциональных возможностей, который требуется получить от API JavaScript для Office. Например, если надстройке требуется доступ на чтение и запись для документа, в манифесте надстройки необходимо указать разрешение `ReadWriteDocument` в качестве текстового значения элемента **Permissions**. Так как разрешения обеспечивают конфиденциальность и безопасность пользователей, рекомендуется запросить минимальный уровень разрешений, необходимый для работы надстройки. В примере ниже показано, как запросить разрешение **ReadDocument** в манифесте надстройки области задач.


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

Дополнительные сведения см. в статье [Запрос разрешений на использование API в контентных надстройках и надстройках области задач](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).


## <a name="see-also"></a>См. также

- [API JavaScript для Office](/office/dev/add-ins/reference/javascript-api-for-office)
- [Справочная схема по манифестам надстроек для Office](../develop/add-in-manifests.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](../testing/testing-and-troubleshooting.md)
