---
title: "Поддержка API\_JavaScript для Office для контентных надстроек и надстроек области задач в Office\_2013"
description: ''
ms.date: 12/04/2017
---


# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Поддержка API JavaScript для Office для контентных надстроек и надстроек области задач в Office 2013


Вы можете использовать [API JavaScript для Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) для создания надстроек области задач или контентных надстроек для ведущих приложений Office 2013. Объекты и методы, поддерживаемые контентными надстройками и надстройками области задач, можно сгруппировать указанным ниже образом.


1. **Стандартные объекты, используемые совместно с другими надстройками Office.** Среди них объекты [Office](https://dev.office.com/reference/add-ins/shared/office), [Context](https://dev.office.com/reference/add-ins/shared/office.context) и [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult). Объект **Office** — это корневой объект API JavaScript для Office. Объект **Context** представляет среду выполнения надстройки. **Office** и **Context** — базовые объекты для любой надстройки Office. Объект **AsyncResult** представляет результаты асинхронной операции, например данные, возвращенные в метод **getSelectedDataAsync**, который считывает сведения о том, какие элементы пользователь выделил в документе.
    
2.  **Объект Document.** Большей частью API, доступной для контентных надстроек и надстроек области задач, можно воспользоваться с помощью методов, свойств и событий объекта [Document](https://dev.office.com/reference/add-ins/shared/document). Контентная надстройка или надстройка области задач может использовать свойство [Office.context.document](https://dev.office.com/reference/add-ins/shared/office.context.document) для доступа к объекту **Document** и с его помощью получать доступ к ключевым компонентам API для работы с данными в документах, например к объектам [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) и [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) и методам [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync)и [getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync). Кроме того, в объекте **Document** имеется свойство [mode](https://dev.office.com/reference/add-ins/shared/document.mode), с помощью которого можно определить, в каком режиме находится документ, в режиме "только для чтения" или в режиме редактирования, свойство [url](https://dev.office.com/reference/add-ins/shared/document.url) для получения URL-адреса текущего документа и доступа к объекту [Settings](https://dev.office.com/reference/add-ins/shared/settings). Объект **Document** также позволяет добавлять обработчики события [SelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event), позволяющие обнаруживать действия пользователя по изменению выделения в документе.
    
   Контентная надстройка или надстройка области задач может получить доступ к объекту **Document** только после загрузки модели DOM и среды выполнения (как правило, это происходит в обработчике события [Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize)). Сведения о потоке событий при инициализации надстройки и о том, как проверить успешность загрузки модели DOM и среды выполнения, см. в статье [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md).
    
3.  **Объекты для работы с конкретными функциями.** Для работы с конкретными функциями API используйте указанные ниже объекты и методы.
    
    - Используйте методы объекта [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) для создания или получения привязок и методы и свойства объекта [Binding](https://dev.office.com/reference/add-ins/shared/binding) для работы с данными.
    
    - Используйте [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts), [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) и связанные с ними объекты для создания пользовательских XML-частей в документах Word и управления ими.
    
    - Используйте объекты [File](https://dev.office.com/reference/add-ins/shared/file) и [Slice](https://dev.office.com/reference/add-ins/shared/slice) для создания копии всего документа, его разбивки на блоки или "фрагменты" и последующего считывания или передачи данных, содержащихся в этих фрагментах.
    
    - Используйте объект [Settings](https://dev.office.com/reference/add-ins/shared/settings) для сохранения пользовательских данных, например настроек пользователей и состояния надстройки.
    

> [!IMPORTANT]
> Некоторые элементы API поддерживаются не всеми приложениями Office, в которых могут размещаться контентные надстройки и надстройки области задач. Чтобы определить, какие элементы поддерживаются, см. один из указанных ниже ресурсов.

Краткие сведения о поддержке API JavaScript для Office ведущими приложениями Office см. в [этой статье](understanding-the-javascript-api-for-office.md).


## <a name="reading-and-writing-to-an-active-selection"></a>Чтение и запись данных в активное выделение

Вы можете считывать данные из текущего выделения пользователя в документе, электронной таблице или презентации, а также записывать их в это выделение. В зависимости от ведущего приложения можно указать тип структуры данных, которая будет считана или записана, в качестве параметра методов [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) и [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) объекта [Document](https://dev.office.com/reference/add-ins/shared/document). Например, вы можете указать любой тип данных (текст, HTML, табличные данные или Office Open XML) для Word, текст и табличные данные для Excel, а также текст для PowerPoint и Project. Вы также можете создать обработчики событий для обнаружения изменений в выделении пользователя. Ниже приведен пример получения данных из выделения в качестве текста с помощью метода **getSelectedDataAsync**.


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

Дополнительные сведения и примеры см. в статье [Чтение и запись данных в текущую выделенную область документа или электронной таблицы](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>Привязка к областям в документе или электронной таблице

Вы можете использовать методы **getSelectedDataAsync** и **setSelectedDataAsync** для чтения и записи данных в *текущее* выделение в документе, электронной таблице или презентации. Тем не менее если вам необходимо обращаться к одной области документа в различных сеансах запущенной надстройки, не принуждая пользователя делать выбор, сначала потребуется создать привязку к этой области. Вы также сможете подписаться на события изменения данных и выделения только для этой привязанной области.

Привязку можно добавить с помощью методов [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync), [addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync) или [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) объекта [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings). Они возвращают идентификатор, который можно использовать для доступа к данным в привязке или для подписки на события изменения данных или выделения в привязанной области.

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

Дополнительные сведения и примеры см. в статье [Привязка к областям документа или электронной таблицы](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="getting-entire-documents"></a>Получение документов целиком

Если надстройка области задач работает в PowerPoint или Word, то для получения презентации или документа целиком вы можете использовать методы [Document.getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync), [File.getSliceAsync](https://dev.office.com/reference/add-ins/shared/file.getsliceasync)и [File.closeAsync](https://dev.office.com/reference/add-ins/shared/file.closeasync).

При вызове метода **Document.getFileAsync** вы получаете копию документа в объекте [File](https://dev.office.com/reference/add-ins/shared/file). Объект **File** обеспечивает доступ к документу в "блоках", представленных в качестве объектов [Slice](https://dev.office.com/reference/add-ins/shared/document). При вызове метода **getFileAsync** можно указать тип файла (текст или сжатый формат Open Office XML) и размер фрагментов (до 4 МБ). Для доступа к содержимому объекта **File** нужно вызвать метод **File.getSliceAsync**, который возвращает необработанные данные в свойстве [Slice.data](https://dev.office.com/reference/add-ins/shared/slice.data). Если вы выбрали сжатый формат, то получите данные файлов в виде массива байтов. Если вы передаете файл в веб-службу, перед отправкой можно преобразовать сжатые необработанные данные в строку с кодировкой Base64. Получив фрагменты файла, закройте документ с помощью метода **File.closeAsync**.

Дополнительные сведения см. в инструкции по [получению всего документа из надстройки для PowerPoint или Word](../word/get-the-whole-document-from-an-add-in-for-word.md). 


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>Чтение и запись настраиваемых XML-частей документа Word

С помощью формата файлов Open Office XML и элементов управления контентом вы можете добавлять пользовательские XML-части в документ Word и привязывать элементы в XML-частях к элементам управления контентом в этом документе. При открытии документа Word считывает и автоматически заполняет привязанные элементы управления контентом данными из пользовательских XML-частей. Кроме того, пользователи могут записывать данные в элементы управления контентом. Когда пользователь сохраняет документ, данные в элементах управления также сохраняются в привязанных XML-частях. Надстройки области задач для Word могут использовать свойство [Document.customXmlParts](https://dev.office.com/reference/add-ins/shared/document.customxmlparts), а также объекты [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts), [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart)и [CustomXmlNode](https://dev.office.com/reference/add-ins/shared/customxmlnode.customxmlnode) для динамического считывания и записи данных в документ.

Пользовательские XML-части можно сопоставлять с пространствами имен. Чтобы получать данные из пользовательских XML-частей в пространстве имен, используйте метод [CustomXmlParts.getByNamespaceAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbynamespaceasync).

Кроме того, вы можете использовать метод [CustomXmlParts.getByIdAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbyidasync), чтобы получить доступ к пользовательским XML-частям по их GUID. После этого можно получить XML-данные с помощью метода [CustomXmlPart.getXmlAsync](https://dev.office.com/reference/add-ins/shared/customxmlpart.getxmlasync).

Чтобы добавить новую пользовательскую XML-часть в документ, с помощью свойства **Document.customXmlParts** получите пользовательские XML-части документа, а затем вызовите метод [CustomXmlParts.addAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.addasync).

Подробные сведения о работе с пользовательскими XML-частями с помощью надстройки области задач см. в статье [Создание улучшенных надстроек для Word с помощью Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).


## <a name="persisting-add-in-settings"></a>Сохранение настроек надстроек


Часто пользовательские настройки, состояние надстройки или другие данные требуется сохранять для работы с ними после перезапуска надстройки. Для этого можно использовать стандартные методы веб-программирования, например файлы cookie в браузере или веб-хранилище HTML 5. Кроме того, если надстройка работает в Excel, PowerPoint или Word, вы можете использовать методы объекта [Settings](https://dev.office.com/reference/add-ins/shared/settings). Данные, созданные с помощью объекта **Settings**, хранятся в электронной таблице, презентации или документе, в который добавлена надстройка и с которым она сохранена. Эти данные доступны только для надстройки, которая их создала.

Во избежание циклических прохождений пакетов на сервере, где хранится документ, управление созданными с помощью объекта **Settings** данными выполняется в памяти в среде выполнения. Ранее сохраненные данные настроек загружаются в память при инициализации надстройки, а внесенные в эти данные изменения сохраняются в документе только при вызове метода [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync). Для внутренних целей данные хранятся в сериализованном объекте JSON в виде пар "имя-значение". Вы можете использовать методы [get](https://dev.office.com/reference/add-ins/shared/settings.get), [set](https://dev.office.com/reference/add-ins/shared/settings.set) и [remove](https://dev.office.com/reference/add-ins/shared/settings.removehandlerasync) объекта **Settings** для чтения, записи и удаления элементов из копии данных, хранящейся в памяти. Ниже приведена строка кода, позволяющая создать настройку с именем `themeColor` и задать для нее значение green.




```js
Office.context.document.settings.set('themeColor', 'green');
```

Так как созданные или удаленные с помощью методов **set** и **remove** данные настроек размещаются в хранящейся в памяти копии данных, для сохранения изменений, внесенных в данные настроек документа, с которым работает надстройка, необходимо вызвать метод **saveAsync**.

Дополнительные сведения о работе с пользовательскими данными с помощью методов объекта **Settings** см. в статье [Сохранение состояния и параметров надстройки](persisting-add-in-state-and-settings.md).


## <a name="reading-properties-of-a-project-document"></a>Чтение свойств документа проекта

Если надстройка области задач выполняется в Project, то она может считывать данные из некоторых полей, ресурсов и полей задач в активном проекте. Для этого используются методы и события объекта [ProjectDocument](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument), которые расширяют объект **Document** путем добавления функций работы с Project.

Примеры считывания данных Project см. в статье [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## <a name="permissions-model-and-governance"></a>Модель разрешений и управление

Ваша надстройка использует элемент **Permissions** в своем манифесте для запроса разрешения на доступ к уровню функциональных возможностей, который требуется получить от API JavaScript для Office. Например, если надстройке требуется доступ на чтение и запись для документа, в манифесте надстройки необходимо указать разрешение `ReadWriteDocument` в качестве текстового значения элемента **Permissions**. Так как разрешения обеспечивают конфиденциальность и безопасность пользователей, рекомендуется запросить минимальный уровень разрешений, необходимый для работы надстройки. В примере ниже показано, как запросить разрешение **ReadDocument** в манифесте надстройки области задач.


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

Дополнительные сведения см. в статье [Запрос разрешений на использование API в контентных надстройках и надстройках области задач](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).


## <a name="see-also"></a>См. также

- [API JavaScript для Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)
- [Типовые схемы для манифестов надстроек Office](../develop/add-in-manifests.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](../testing/testing-and-troubleshooting.md)
    
