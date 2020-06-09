---
title: Поддержка API JavaScript для Office для контентных надстроек и надстроек области задач в Office 2013
description: Используйте API JavaScript для Office, чтобы создать область задач в Office 2013.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 334db88bbec07755678e3ba35e0d4998951ff5ab
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609708"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Поддержка API JavaScript для Office для контентных надстроек и надстроек области задач в Office 2013

[!include[information about the common API](../includes/alert-common-api-info.md)]

Вы можете использовать [API JavaScript для Office](../reference/javascript-api-for-office.md) для создания надстроек области задач или контентных надстроек для ведущих приложений Office 2013. Объекты и методы, поддерживаемые контентными надстройками и надстройками области задач, можно сгруппировать указанным ниже образом.

1. **Общие объекты, которые используются совместно с другими надстройками Office.** К этим объектам относятся [Office](/javascript/api/office), [context](/javascript/api/office/office.context)и [asyncResult](/javascript/api/office/office.asyncresult). `Office`Объект является корневым объектом API JavaScript для Office. `Context`Объект представляет среду выполнения надстройки. Оба `Office` и `Context` являются основными объектами для любой надстройки Office. `AsyncResult`Объект представляет результаты асинхронной операции, например данные, возвращенные в `getSelectedDataAsync` метод, который считывает сведения о том, что пользователь выбрал в документе.

2. **Объект Document.** Большей частью API, доступной для контентных надстроек и надстроек области задач, можно воспользоваться с помощью методов, свойств и событий объекта [Document](/javascript/api/office/office.document). Контентная надстройка или надстройка области задач может использовать свойство [Office.context.document](/javascript/api/office/office.context#document) для доступа к объекту **Document** и с его помощью получать доступ к ключевым компонентам API для работы с данными в документах, например к объектам [Bindings](/javascript/api/office/office.bindings) и [CustomXmlParts](/javascript/api/office/office.customxmlparts) и методам [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-), [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)и [getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-). `Document`Объект также предоставляет свойство [mode](/javascript/api/office/office.document#mode) для определения того, является ли документ предназначен только для чтения или находится в режиме редактирования, свойство [URL](/javascript/api/office/office.document#url) для получения URL-адреса текущего документа и доступа к объекту [Settings](/javascript/api/office/office.settings) . `Document`Объект также поддерживает добавление обработчиков событий для события [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) , чтобы можно было обнаружить, когда пользователь изменяет свой выбор в документе.

   Надстройка области задач или контентная надстройка может получать доступ к `Document` объекту только после загрузки модели DOM и среды выполнения, обычно в обработчике события для события [Office. Initialize](/javascript/api/office) . Сведения о потоке событий при инициализации надстройки и проверке успешности загрузки модели DOM и среды выполнения см. в разделе [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).

3. **Объекты для работы с конкретными функциями.** Для работы с конкретными функциями API используйте указанные ниже объекты и методы.

    - Используйте методы объекта [Bindings](/javascript/api/office/office.bindings) для создания или получения привязок и методы и свойства объекта [Binding](/javascript/api/office/office.binding) для работы с данными.

    - Используйте объекты [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) и сопоставленные с ними объекты для создания пользовательских XML-частей в документах Word и управления ими.

    - Используйте объекты [File](/javascript/api/office/office.file) и [Slice](/javascript/api/office/office.slice) для создания копии всего документа, его разбивки на блоки или "фрагменты" и последующего считывания или передачи данных, содержащихся в этих фрагментах.

    - Используйте объект [Settings](/javascript/api/office/office.settings) для сохранения пользовательских данных, например настроек пользователей и состояния надстройки.


> [!IMPORTANT]
> Некоторые элементы API поддерживаются не всеми приложениями Office, в которых могут размещаться контентные надстройки и надстройки области задач. Чтобы определить, какие элементы поддерживаются, см. один из указанных ниже ресурсов.

Сводка по поддержке API JavaScript для Office в ведущих приложениях Office приведена [в статье Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md).


## <a name="reading-and-writing-to-an-active-selection"></a>Чтение и запись данных в активное выделение

Вы можете считывать данные из текущего выделения пользователя в документе, электронной таблице или презентации, а также записывать их в это выделение. В зависимости от ведущего приложения можно указать тип структуры данных, которая будет считана или записана, в качестве параметра методов [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) и [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) объекта [Document](/javascript/api/office/office.document). Например, вы можете указать любой тип данных (текст, HTML, табличные данные или Office Open XML) для Word, текст и табличные данные для Excel, а также текст для PowerPoint и Project. Вы также можете создать обработчики событий для обнаружения изменений в выделении пользователя. В примере ниже показано, как получить данные из выделенного фрагмента в виде текста с помощью `getSelectedDataAsync` метода.


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

Можно использовать `getSelectedDataAsync` `setSelectedDataAsync` методы и для чтения или записи в выделенный фрагмент пользователя в *current* документе, электронной таблице или презентации. Тем не менее если вам необходимо обращаться к одной области документа в различных сеансах запущенной надстройки, не принуждая пользователя делать выбор, сначала потребуется создать привязку к этой области. Вы также сможете подписаться на события изменения данных и выделения только для этой привязанной области.

Привязку можно добавить с помощью методов [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-), [addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-) или [addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) объекта [Bindings](/javascript/api/office/office.bindings). Они возвращают идентификатор, который можно использовать для доступа к данным в привязке или для подписки на события изменения данных или выделения в привязанной области.

Ниже приведен пример, в котором показано, как добавить привязку к текущему выбранному тексту в документе с помощью `Bindings.addFromSelectionAsync` метода.



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

При вызове `Document.getFileAsync` вы получите копию документа в объекте [File](/javascript/api/office/office.file) . `File`Объект предоставляет доступ к документу в виде фрагментов, представленных в виде объектов [slice](/javascript/api/office/office.slice) . При вызове `getFileAsync` можно указать тип файла (текст или сжатый формат Open Office XML) и размер срезов (до 4 МБ). Чтобы получить доступ к содержимому `File` объекта, затем вызывается метод, `File.getSliceAsync` который возвращает необработанные данные в свойстве [slice. Data](/javascript/api/office/office.slice#data) . Если вы выбрали сжатый формат, то получите данные файлов в виде массива байтов. Если вы передаете файл в веб-службу, перед отправкой можно преобразовать сжатые необработанные данные в строку с кодировкой Base64. Наконец, когда вы закончите извлечение фрагментов файла, используйте метод, `File.closeAsync` чтобы закрыть документ.

Дополнительные сведения см. в инструкции по [получению документа целиком из надстройки для PowerPoint или Word](../word/get-the-whole-document-from-an-add-in-for-word.md).


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>Чтение и запись настраиваемых XML-частей документа Word

С помощью формата файлов Open Office XML и элементов управления контентом вы можете добавлять пользовательские XML-части в документ Word и привязывать элементы в XML-частях к элементам управления контентом в этом документе. При открытии документа Word считывает и автоматически заполняет привязанные элементы управления контентом данными из пользовательских XML-частей. Кроме того, пользователи могут записывать данные в элементы управления контентом. Когда пользователь сохраняет документ, данные в элементах управления также сохраняются в привязанных XML-частях. Надстройки области задач для Word могут использовать свойство [Document.customXmlParts](/javascript/api/office/office.document#customxmlparts), а также объекты [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart)и [CustomXmlNode](/javascript/api/office/office.customxmlnode) для динамического считывания и записи данных в документ.

Пользовательские XML-части можно сопоставить с пространствами имен. Чтобы получить данные из пользовательских XML-частей в пространстве имен, используйте метод [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getbynamespaceasync-ns--options--callback-).

Кроме того, вы можете использовать метод [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-), чтобы получить доступ к пользовательским XML-частям по их GUID. После этого можно получить XML-данные с помощью метода [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getxmlasync-options--callback-).

Чтобы добавить в документ новую пользовательскую XML-часть, используйте `Document.customXmlParts` свойство для получения настраиваемых XML-частей в документе и вызовите метод [CustomXmlParts. addAsync](/javascript/api/office/office.customxmlparts#addasync-xml--options--callback-) .

Подробные сведения о работе с пользовательскими XML-частями с помощью надстройки области задач см. в статье [Создание улучшенных надстроек для Word с помощью Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).


## <a name="persisting-add-in-settings"></a>Сохранение настроек надстроек


Надстройкам часто требуется сохранять пользовательские данные, например настройки пользователя или состояние надстройки, и запрашивать их при следующем запуске. Для этого можно использовать стандартные методы веб-программирования, например файлы cookie в браузере или веб-хранилище HTML 5. Кроме того, если надстройка работает в Excel, PowerPoint или Word, вы можете использовать методы объекта [Settings](/javascript/api/office/office.settings). Данные, созданные с помощью `Settings` объекта, хранятся в электронной таблице, презентации или документе, в которую надстройка была вставлена и сохранена. Данные доступны только для создавшей их надстройки.

Чтобы избежать обмена данными с сервером, на котором хранится документ, данные, созданные с помощью `Settings` объекта, управляются в памяти во время выполнения. Ранее сохраненные данные настроек загружаются в память при инициализации надстройки, а внесенные в эти данные изменения сохраняются в документе только при вызове метода [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-). Для внутренних целей данные хранятся в сериализованном объекте JSON в виде пар "имя-значение". Вы можете использовать методы [get](/javascript/api/office/office.settings#get-name-), [set](/javascript/api/office/office.settings#set-name--value-) и [remove](/javascript/api/office/office.settings#remove-name-) объекта **Settings** для чтения, записи и удаления элементов из копии данных, хранящейся в памяти. Ниже приведена строка кода, позволяющая создать настройку с именем `themeColor` и задать для нее значение green.




```js
Office.context.document.settings.set('themeColor', 'green');
```

Так как данные параметров, созданные или удаленные с помощью `set` методов и, `remove` работают с копией данных в памяти, необходимо вызвать метод `saveAsync` сохранения изменений данных параметров в документе, с которым работает ваша надстройка.

Дополнительные сведения о работе с настраиваемыми данными с помощью методов `Settings` объекта см. в статье [Сохранение состояния и параметров надстройки](persisting-add-in-state-and-settings.md).


## <a name="reading-properties-of-a-project-document"></a>Чтение свойств документа проекта

Если надстройка области задач выполняется в Project, то она может считывать данные из некоторых полей, ресурсов и полей задач в активном проекте. Для этого используются методы и события объекта [ProjectDocument](/javascript/api/office/office.document) , которые расширяют `Document` объект для предоставления дополнительных функциональных возможностей, связанных с проектом.

Примеры считывания данных Project см. в статье [Создание первой надстройки области задач для Project 2013 с использованием текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## <a name="permissions-model-and-governance"></a>Модель разрешений и управление

Надстройка использует `Permissions` элемент в своем манифесте, чтобы запросить разрешение на доступ к требуемому уровню функциональности API JavaScript для Office. Например, если надстройке требуется доступ для чтения и записи к документу, его манифест должен указывать в `ReadWriteDocument` качестве текстового значения в `Permissions` элементе. Так как разрешения обеспечивают конфиденциальность и безопасность пользователей, рекомендуется запросить минимальный уровень разрешений, необходимый для работы надстройки. В примере ниже показано, как запросить разрешение **ReadDocument** в манифесте надстройки области задач.


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

Дополнительные сведения см в статье [запрашивание разрешений для использования API в](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)надстройках.


## <a name="see-also"></a>См. также

- [API JavaScript для Office](../reference/javascript-api-for-office.md)
- [Справочная схема по манифестам надстроек для Office](../develop/add-in-manifests.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](../testing/testing-and-troubleshooting.md)
