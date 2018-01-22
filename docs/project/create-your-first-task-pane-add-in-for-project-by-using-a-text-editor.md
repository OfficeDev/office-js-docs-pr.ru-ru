
# <a name="create-your-first-task-pane-add-in-for-project-2013-by-using-a-text-editor"></a>Создание первой надстройки области задач для Project 2013 с помощью текстового редактора

Надстройку области задач для Project стандартный 2013 или Project профессиональный 2013 можно создать с помощью Visual Studio 2015 (подходит для создания сложного веб-приложения) или с помощью текстового редактора для создания файлов локальной надстройки. В этой статье описывается, как создать простую надстройку, в которой используется манифест XML, указывающий на HTML-файл в общей папке. Пример надстройки Project OM Test проверяет некоторые функции JavaScript, которые используют объектную модель для надстроек. После использования **центра управления безопасностью** в Project 2013 для регистрации общей папки, содержащей файл манифеста, можно открыть надстройку области задач на вкладке ленты **ПРОЕКТ**. (Код примера в этой статье основан на тестовом приложении, написанном Арвиндом Айером (Arvind Iyer), специалистом корпорации Майкрософт.)

В Project 2013 используется та же схема манифеста надстройки, что и в других клиентах Microsoft Office 2013, и, в основном, тот же самый интерфейс API JavaScript. Полный код надстройки, описанной в этой статье, доступен в подкаталоге `Samples\Apps` загружаемого пакета SDK для Project 2013.

Пример надстройки Project OM Test может получить GUID задач и свойств приложения и активного проекта. Если Project профессиональный 2013 открывает проект, находящийся в библиотеке SharePoint, надстройка может показать URL-адрес проекта. [Загружаемый пакет SDK для Project 2013](https://www.microsoft.com/en-us/download/details.aspx?id=30435%20) содержит полный исходный код. При извлечении и установке пакета SDK и примеров, находящихся в файле Project2013SDK.msi, просмотрите подкаталог `\Samples\Apps\Copy_to_AppManifests_FileShare`, где содержится файл манифеста, а также подкаталог `\Samples\Apps\Copy_to_AppSource_FileShare`, где хранится исходный код. В примере JSOMCall.html используются функции JavaScript, которые находятся в файлах office.js и project-15.js, также входящих в пакет. Можно использовать соответствующие файлы отладки (office.debug.js и project-15.debug.js) для изучения работы функций. Общие сведения об использовании JavaScript в надстройках Office см. в статье [Общие сведения об интерфейсе API JavaScript для Office](../../docs/develop/understanding-the-javascript-api-for-office.md).

## <a name="procedure-1-to-create-the-add-in-manifest-file"></a>Процедура 1. Создание файла манифеста надстройки



- Создайте XML-файл в локальном каталоге. XML-файл включает элемент **OfficeApp** и дочерние элементы, которые описаны в статье, посвященной [XML-манифесту надстроек для Office](../../docs/overview/add-in-manifests.md). Например, создайте файл с именем JSOM_SimpleOMCalls.xml, содержащий следующий XML-код (измените значение GUID элемента **Id**).
    
```XML
     <?xml version="1.0" encoding="utf-8"?>
   <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
              xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
              xsi:type="TaskPaneApp">
     <Id>93A26520-9414-492F-994B-4983A1C7A607</Id>
     <Version>15.0</Version>
     <ProviderName>Microsoft</ProviderName>
     <DefaultLocale>en-us</DefaultLocale>
     <DisplayName DefaultValue="Project OM Test">
       <Override Locale="fr-fr" Value="Le Project OM Test"/>
     </DisplayName>
     <Description DefaultValue="Test the task pane add-in object model for Project - English (US)">
       <Override Locale="fr-fr" Value="Test the task pane add-in object model for Project - French (France)"/>
     </Description>
     <Hosts>
       <Host Name="Project"/>
       <Host Name="Workbook"/>
       <Host Name="Document"/>
     </Hosts>
    <DefaultSettings>
       <SourceLocation DefaultValue="\\ServerName\AppSource\JSOMCall.html">
         <Override Locale="fr-fr" Value="\\ServerName\AppSource\JSOMCall.html"/>
       </SourceLocation>
     </DefaultSettings>
     <Permissions>ReadWriteDocument</Permissions>
     <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
       <Override Locale="fr-fr" Value="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>
     </IconUrl>
     <AllowSnapshot>true</AllowSnapshot>
   </OfficeApp>
```


    For Project, the  **OfficeApp** element must include the `xsi:type="TaskPaneApp"` attribute value. The **Id** element is a GUID. The **SourceLocation** value must be a file share path or a SharePoint URL for the add-in HTML source file or the web application that runs in the task pane. For an explanation of the other elements in manifest file, see [Task pane add-ins for Project](../project/project-add-ins.md).
    
В процедуре 2 показано, как создавать файл HTML, который манифест JSOM_SimpleOMCalls.xml определяет как тестовую надстройку для Project. Кнопки, определенные в HTML-файле, вызывают связанные функции JavaScript. Можно добавить функции JavaScript в HTML-файл или поместить их в отдельный JS-файл.

## <a name="procedure-2-to-create-the-source-files-for-the-project-om-test-add-in"></a>Процедура 2. Создание исходных файлов для надстройки Project OM Test



1. Создайте HTML-файл с именем, указанным в элементе **SourceLocation** манифеста JSOM_SimpleOMCalls.xml. Например, создайте файл JSOMCall.html в каталоге `C:\Project\AppSource`. Хотя вы можете создавать исходные файлы с помощью простого текстового редактора, проще использовать такой инструмент, как Visual Studio 2015, который работает с определенными типами документов (например, HTML и JavaScript) и содержит различные вспомогательные компоненты, упрощающие редактирование. Если вы еще не делали пример с поиском Bing, описанный в статье [Надстройки области задач для Project](../project/project-add-ins.md), просмотрите процедуру 3, чтобы узнать, как создавать общую папку `\\ServerName\AppSource`, на которую указывает манифест.
    
    Файл JSOMCall.html использует общий файл MicrosoftAjax.js для функций AJAX, а файл Office.js — для функций надстройки в приложениях Microsoft Office 2013.
    


```HTML
  <!DOCTYPE html>
<html>
<head>
    <title>Project OM Sample Code</title>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <script type="text/javascript" src="MicrosoftAjax.js"></script>

    <!-- Use the CDN reference to office.js when deploying your add-in. -->
    <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"></script> -->
    <script type="text/javascript" src="Office.js"></script>
    <script type="text/javascript" src="JSOM_Sample.js"></script>
</head>
<body>
    <div id="Common_JSOM_API">
        OBJECT MODEL TESTS
    </div>

    <textarea id="text" rows="6" cols="25">This is the text result.</textarea>
</body>
</html>
```


    The  **textarea** element specifies a text box that shows results of the JavaScript functions.
    
     >**Note**  For the Project OM Test sample to work, copy the following files from the Project 2013 SDK download to the same directory as the JSOMCall.html file: Office.js, Project-15.js, and MicrosoftAjax.js.

    Step 2 adds the JSOM_Sample.js file for specific functions that the Project OM Test sample add-in uses. In later steps, you will add other HTML elements for buttons that call JavaScript functions.
    
2. Создайте файл JavaScript с именем JSOM_Sample.js в том же каталоге, где находится файл JSOMCall.html. Следующий код позволяет получить контекст приложения и сведения о документе с помощью функций в файле Office.js. Объект **text** является идентификатором элемента управления **textarea** в HTML-файле.
    
    Переменная **_projDoc** инициализируется с объектом **ProjectDocument**. Код включает в себя функции обработки простых ошибок, а также функцию **getContextValues**, позволяет получить контекст приложения и свойства контекста для документа проекта. Дополнительные сведения об объектной модели JavaScript для Project см. в статье [API JavaScript для Office](http://dev.office.com/reference/add-ins/javascript-api-for-office).
    


```js
  /*
* JavaScript functions for the Project OM Test example app
* in the Project 2013 SDK.
*/

var _projDoc;
var _app;
var taskGuid;
var resourceGuid;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        _projDoc = Office.context.document;
        _app = Office.context;
    });
}

function logError(errorText) {
    text.value = "Error in " + errorText;
}

function logEventError(erroneousEvent) {
    logError("event " + erroneousEvent);
}

function logMethodError(methodName, errorName, errorMessage) {
    logError(methodName + " method.\nError name: " + errorName + "\nMessage: " + errorMessage);
}

// . . . Add other JavaScript functions here.

function getContextValues() {
    getDocumentUrl();
    getDocumentMode();
    getApplicationContentLanguage();
    getApplicationDisplayLanguage();
}

function getDocumentUrl() {
    text.value ="Document URL:\n" +_projDoc.url;
}

function getDocumentMode() {
    var docMode = _projDoc.mode;
    text.value = text.value + "\n\nDocument mode: " + docMode;
}

function getApplicationContentLanguage() {
    text.value = text.value + "\nApp language: " + _app.contentLanguage;
}

function getApplicationDisplayLanguage() {
    text.value = text.value + "\nDisplay language: " + _app.displayLanguage;
}
```


    For information about the functions in the Office.debug.js file, see [JavaScript API for Office](http://dev.office.com/reference/add-ins/javascript-api-for-office). For example, the  **getDocumentUrl** function gets the URL or file path of the open project.
    
3. Добавьте функции JavaScript, которые вызывают асинхронные функции в Office.js и Project-15.js для получения выбранных данных.
    
      - Например, **getSelectedDataAsync** — это общая функция в Office.js, которая принимает неформатированный текст из выбранных данных. Дополнительные сведения см. в статье [Объект AsyncResult](http://dev.office.com/reference/add-ins/shared/asyncresult).
    
  - Функция **getSelectedTaskAsync** в файле Project-15.js принимает идентификатор GUID выбранной задачи. Аналогичным образом функция **getSelectedResourceAsync** получает GUID выбранного ресурса. Если вызвать эти функции, не выбрав задачи или ресурса, функции отобразят неопределенную ошибку.
    
  - Функция **getTaskAsync** получает имя задачи и имена назначенных ресурсов. Если задача находится в синхронизированном списке задач SharePoint, **getTaskAsync** получает идентификатор задач из списка SharePoint; в противном случае идентификатор задачи SharePoint равен 0.
    
     >**Note**  For demonstration purposes, the example code includes a bug. If  **taskGuid** is undefined, the **getTaskAsync** function errors off. If you get a valid task GUID and then select a different task, the **getTaskAsync** function gets data for the most recent task that was operated on by the **getSelectedTaskAsync** function.
  -  **getTaskFields**, **getResourceFields** и **getProjectFields** являются локальными функциями, которые вызывают **getTaskFieldAsync**, **getResourceFieldAsync** или **getProjectFieldAsync** несколько раз для получения указанных полей задачи или ресурса, в файле project-15.debug.js перечисления **ProjectTaskFields** и **ProjectResourceFields** показывают, какие поля поддерживаются.
    
  - Функция **getSelectedViewAsync** получает тип представления (определяется в перечислении **ProjectViewTypes** в файле project-15.debug.js) и имя представления.
    
  - Если проект синхронизируется со списком задач SharePoint, функция **getWSSUrlAsync** получает URL-адрес и имя списка задач. Если проект не синхронизируется со списком задач SharePoint, функция **getWSSUrlAsync** выдает ошибку.
    
     >**Примечание.** Чтобы получить URL-адрес SharePoint и имя списка задач, рекомендуем использовать функцию **getProjectFieldAsync** с константами **WSSUrl** и **WSSList** в перечислении [ProjectProjectFields](http://dev.office.com/reference/add-ins/shared/projectprojectfields-enumeration)

    Все функции в коде ниже содержат анонимную функцию, которую определяет `function (asyncResult)`обратная функция, получающая асинхронный результат. Вместо анонимных можно использовать именованные функции. Благодаря этому будет удобней обеспечивать поддержку сложных надстроек.
    


```js
  // Get the data in the selected cells of the grid in the active view.
function getSelectedDataAsync() {
    _projDoc.getSelectedDataAsync(
        Office.CoercionType.Text,
        { ValueFormat: "Formatted" },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded)
                text.value = asyncResult.value;
            else
                logMethodError("getSelectedDataAsync", asyncResult.error.name,
                               asyncResult.error.message);
        }
    );
}

// Get the GUID of the selected task.
function getSelectedTaskAsync() {
    _projDoc.getSelectedTaskAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            text.value = asyncResult.value;
            taskGuid = asyncResult.value;
        }
        else {
            logMethodError("getSelectedTaskAsync", asyncResult.error.name,
                               asyncResult.error.message);
        }
    });
}

// Get the GUID of the selected resource.
function getSelectedResourceAsync() {
    _projDoc.getSelectedResourceAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            text.value = asyncResult.value;
            resourceGuid = asyncResult.value;
        }
        else {
            logMethodError("getSelectedResourceAsync", asyncResult.error.name,
                               asyncResult.error.message);
        }
    });
}

// Get data for the specified task.
function getTaskAsync() {
    if (taskGuid != undefined) {
        _projDoc.getTaskAsync(
            taskGuid,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    logMethodError("getTaskAsync", asyncResult.error.name,
                               asyncResult.error.message);
                } else {
                    var taskInfo = asyncResult.value;
                    var taskOutput = "Task name: " + taskInfo.taskName +
                                     "\nGUID: " + taskGuid +
                                     "\nWSS Id: " + taskInfo.wssTaskId +
                                     "\nResourceNames: " + taskInfo.resourceNames;
                    text.value = taskOutput;
                }
            }
        );
    } else {
        text.value = 'Task GUID not valid:\n' + taskGuid;
    } 
}

// Get additional data for task fields.
function getTaskFields() {
    text.value = "";

    _projDoc. getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Name,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Name: "
                    + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getTaskFieldAsync", asyncResult.error.name,
                               asyncResult.error.message);
            }
        }
    );

    _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.ID,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "ID: "
                    + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getTaskFieldAsync", asyncResult.error.name,
                               asyncResult.error.message);
            }
        }
    );

    _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Start,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Start: "
                    + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getTaskFieldAsync", asyncResult.error.name,
                               asyncResult.error.message);
            }
        }
    );

    _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Duration,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Duration: "
                    + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getTaskFieldAsync", asyncResult.error.name,
                               asyncResult.error.message);
            }
        }
    );

    _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Priority,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Priority: "
                    + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getTaskFieldAsync", asyncResult.error.name,
                               asyncResult.error.message);
            }
        }
    );

    _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Notes,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Notes: "
                    + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getTaskFieldAsync", asyncResult.error.name,
                               asyncResult.error.message);
            }
        }
    ); 
}

// Get data for the specified resource fields.
function getResourceFields() {
    text.value = "";

    _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Name,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Resource name: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getResourceFieldAsync", asyncResult.error.name,
                               asyncResult.error.message);
            }
        }
    );

    _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Cost,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Cost: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getResourceFieldAsync", asyncResult.error.name,
                               asyncResult.error.message);
            }
        }
    );

    _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.StandardRate,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Standard Rate: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getResourceFieldAsync", asyncResult.error.name, asyncResult.error.message);
            }
        }
    );

    _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.ActualCost,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Actual Cost: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getResourceFieldAsync", asyncResult.error.name, asyncResult.error.message);
            }
        }
    );

    _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.ActualWork,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Actual Work: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getResourceFieldAsync", asyncResult.error.name,
                               asyncResult.error.message);
            }
        }
    );

    _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Units,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Units: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getResourceFieldAsync", asyncResult.error.name,
                               asyncResult.error.message);
            }
        }
    );
}

// Get the URL and list name of the synchronized SharePoint task list.
// Recommended: use getProjectField instead.
function getWSSUrlAsync() {
    _projDoc.getWSSUrlAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            text.value = "SharePoint URL:\n" + asyncResult.value.serverUrl
                + "\nList name: " + asyncResult.value.listName;
        }
        else {
            logMethodError("getWSSUrlAsync", asyncResult.error.name, asyncResult.error.message);
        }
    });
}

// Get the type and name of the selected view.
function getSelectedViewAsync() {
    _projDoc.getSelectedViewAsync(function (asyncResult) {
        text.value = "View type: " + asyncResult.value.viewType
            + "\nName: " + asyncResult.value.viewName;
    });
}

// Get information about the active project.
function getProjectFields() {
    text.value = "";

    _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.GUID,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Project GUID: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
            }
        }
    );

    _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.Start,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "\nStart: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
            }
        }
    );

    _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.Finish,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "\nFinish: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getProject " + errorText);
            }
        }
    );

    _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencyDigits,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "\nCurrency digits: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
            }
        }
    );


    _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencySymbol,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "Currency symbol: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
            }
        }
    );

    _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencySymbolPosition,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "\nSymbol position: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
            }
        }
    );

    _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.ProjectServerUrl,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "\nProject web app URL:\n  " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
            }
        }
    );

    _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.WSSUrl,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "\nSharePoint URL:\n  " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
            }
        }
    );

    _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.WSSList,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = text.value + "\nSharePoint list: " + asyncResult.value.fieldValue + "\n";
            }
            else {
                logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}
```

4. Добавьте обратные вызовы и функции обработчика событий JavaScript для регистрации обработчиков событий для случаев изменения выделения задачи, выделения ресурса и выделения представления, а также для отмены регистрации обработчиков событий. Функция **manageEventHandlerAsync** добавляет или удаляет указанный обработчик события в зависимости от параметра _operation_. Возможны два типа операций: **addHandlerAsync** или **removeHandlerAsync**.
    
    С помощью функций **manageTaskEventHandler**, **manageResourceEventHandler** и **manageViewEventHandler** можно добавлять и удалять обработчик событий в соответствии с параметром _docMethod_.
    


```js
  // Task selection changed event handler.
function onTaskSelectionChanged(eventArgs) {
    text.value = "In task selection change event handler";
}

// Resource selection changed event handler.
function onResourceSelectionChanged(eventArgs) {
    text.value = "In Resource selection changed event handler";
}

// View selection changed event handler.
function onViewSelectionChanged(eventArgs) {
    text.value = "In View selection changed event handler";
}

// Add or remove the specified event handler.
function manageEventHandlerAsync(eventType, handler, operation, onComplete) {
    _projDoc[operation]   //The operation is addHandlerAsync or removeHandlerAsync.
    (
        eventType,
        handler,
        function (asyncResult) {
            if (onComplete) {
                onComplete(asyncResult, operation);
            } else {
                var message = "Operation: " + operation;
                message = message + "\nStatus: " + asyncResult.status + "\n";
                text.value = message;
            }
        }
    );
}

// Write the asyncResult status from the manageEventHandlerAsync function (optional). 
function onComplete(asyncResult, operation) {
    var message = "In onComplete function for " + operation;
    message = message + "\nStatus: " + asyncResult.status;
    text.value = message;
}

// Add or remove a task selection changed event handler.
function manageTaskEventHandler(docMethod) {
    manageEventHandlerAsync(
        Office.EventType.TaskSelectionChanged,      // The task selection changed event.
        onTaskSelectionChanged,                     // The event handler.
        docMethod,                // The Office.Document method to add or remove an event handler.
        onComplete                // Manages the successful asyncResult data (optional).
    );
}

// Add or remove a resource selection changed event handler.
function manageResourceEventHandler(docMethod) {
    manageEventHandlerAsync(
        Office.EventType.ResourceSelectionChanged,  // The resource selection changed event.
        onResourceSelectionChanged,                 // The event handler.
        docMethod,                // The Office.Document method to add or remove an event handler.
        onComplete                // Manages the successful asyncResult data (optional).
    );
}

// Add or remove a view selection changed event handler.
function manageViewEventHandler(docMethod) {
    manageEventHandlerAsync(
        Office.EventType.ViewSelectionChanged,      // The view selection changed event.
        onViewSelectionChanged,                     // The event handler.
        docMethod,                // The Office.Document method to add or remove an event handler.
        onComplete                // Manages the successful asyncResult data (optional).
    );
}
```

5. Что касается основной части документа HTML, добавьте кнопки, которые вызывают функции JavaScript для тестирования. Например, добавьте в элемент **div** интерфейса JSOM API кнопку ввода, которая вызывает общую функцию **getSelectedDataAsync**.
    
```HTML
  <body>
    <div id="Common_JSOM_API">
    OBJECT MODEL TESTS
    <br /><br />       
    <strong>General function:</strong>
    <br />
    <input id="Button5" class="button-wide" type="button" onclick="getSelectedDataAsync()" 
        value="getSelectedDataAsync" />
    </div>
   <!--  more code . . .  -->
```

6. Добавьте раздел **div** с кнопками для функций задач, относящихся к проекту, а также для события **TaskSelectionChanged**.
    
```HTML
  <div id="ProjectSpecificTask">
  <br />
  <strong>Project-specific task methods:</strong><br />
  <button class="button-wide" onclick="getSelectedTaskAsync()">getSelectedTaskAsync</button><br />
  <button class="button-wide" onclick="getTaskAsync()">getTaskAsync</button><br />
  <button class="button-wide" onclick="getTaskFields()">Get Task Fields</button><br />
  <button class="button-wide" onclick="getWSSUrlAsync()">getWSSUrlAsync</button>
  <strong>Task selection changed:</strong>
  <button class="button-narrow" onclick="manageTaskEventHandler('addHandlerAsync')">Add</button>
  <button class="button-narrow" onclick="manageTaskEventHandler('removeHandlerAsync')">Remove</button>         
</div>
```

7. Добавьте разделы **div** с кнопками для методов и событий ресурсов, методов и событий представления, свойств проекта и свойств контекста.
    
```HTML
  <div id="ResourceMethods">
  <br />
  <strong>Resource methods:</strong>
  <button class="button-wide" onclick="getSelectedResourceAsync()">getSelectedResourceAsync</button><br />
  <button class="button-wide" onclick="getResourceFields()">Get Resource Fields</button><br />
  <strong>Resource selection changed:</strong>
  <button class="button-narrow" onclick="manageResourceEventHandler('addHandlerAsync')">Add</button>
  <button class="button-narrow" onclick="manageResourceEventHandler('removeHandlerAsync')">Remove</button>
</div>
<div id="ViewMethods">
  <br />
  <strong>View method:</strong>
  <button class="button-wide" onclick="getSelectedViewAsync()">getSelectedViewAsync</button><br />
  <strong>View selection changed:</strong>
  <button class="button-narrow" onclick="manageViewEventHandler('addHandlerAsync')">Add</button>
  <button class="button-narrow" onclick="manageViewEventHandler('removeHandlerAsync')">Remove</button>         
</div>
<div id="ProjectMethods">
  <br />
  <strong>Project properties:</strong>
  <button class="button-wide" onclick="getProjectFields()">Get Project Fields</button><br />
</div>
<div id="ContextVariables">
  <br />
  <strong>Context properties:</strong>
  <button class="button-wide" onclick="getContextValues()">Get Context Values</button>
</div>
```

8. Чтобы отформатировать элементы кнопок, добавьте элемент таблицы стилей **style**. Например, добавьте следующий код, как дочерний объект элемента **head**.
    
```HTML
  <style type="text/css">
    .button-wide
    {
        width: 210px;
        margin-top: 2px;
    }
    .button-narrow
    {
        width: 80px;
        margin-top: 2px;
    }
</style>
```


     >**Note**  The  **Task Pane Add-in (Project)** template in Visual Studio 2015 includes default .css files for a common look and feel of add-ins.
В процедуре 3 показано, как устанавливать и использовать функциональные возможности надстройки Project OM Test.

## <a name="procedure-3-to-install-and-use-the-project-om-test-add-in"></a>Процедура 3. Установка и использование надстройки Project OM Test



1. Создайте общую папку для хранения манифеста JSOM_SimpleOMCalls.xml. Можно создать общую папку на локальном компьютере или на удаленном компьютере, если к нему есть доступ из сети. Например, если манифест расположен в каталоге `C:\Project\AppManifests` на локальном компьютере, выполните следующую команду:
    
```
  Net share AppManifests=C:\Project\AppManifests
```

    
2. Создайте сетевую папку для размещения файлов HTML и JavaScript, относящихся к надстройке Project OM Test. Убедитесь, что путь к сетевой папке совпадает с путем, указанным в манифесте JSOM_SimpleOMCalls.xml. Например, если файлы расположены в каталоге `C:\Project\AppSource` на локальном компьютере, выполните следующую команду:
    
```
  net share AppSource=C:\Project\AppSource
```

3. Откройте в Project диалоговое окно **Параметры Project**, выберите **Центр управления безопасностью**, затем **Параметры центра управления безопасностью**.
    
    С действиями, необходимыми для регистрации надстройки, и дополнительными сведениями также можно ознакомиться в статье [Надстройки области задач для Project](../project/project-add-ins.md).
    
4. В диалоговом окне **Центр управления безопасностью** выберите в левой области **Надежные каталоги надстроек**.
    
5. Если уже добавлен путь `\\ServerName\AppManifests` к надстройке "Поиск Bing", пропустите это действие. В противном случае в области **Доверенные каталоги надстроек** укажите путь `\\ServerName\AppManifests` в текстовом окне **URL-адрес каталога**, выберите пункт **Добавить каталог**, включите сетевую папку как источник по умолчанию (см. рис. 1), а затем нажмите кнопку **ОК**.
    
    **Рис. 1. Добавление сетевой папки для манифестов надстроек**

    ![Добавление сетевой папки для манифестов приложений](../images/pj15_CreateSimpleAgave_ManageCatalogs.png)

6. После добавления новых надстроек или изменения исходного кода перезапустите Project. На ленте **ПРОЕКТ** выберите в раскрывающемся меню **Надстройки Office** значение **Просмотреть все**. В диалоговом окне **Вставить надстройку** выберите **ОБЩАЯ ПАПКА** (см. рис. 2), выберите **Project OM Test**, затем **Вставить**. Надстройка Project OM Test запустится в области задач.
    
    **Рис. 2. Запуск надстройки Project OM Test, расположенной в общей папке**

    ![Вставка приложения](../images/pj15_CreateSimpleAgave_StartAgaveApp.png)

7. В Project создайте и сохраните простой проект, который содержит хотя бы две задачи. Например, создайте задачи с именами T1, T2 и веху с именем M1, затем задайте длительности задач и их предшественников примерно как показано на рис. 3. Выберите вкладку **ПРОЕКТ** на ленте выберите всю строку задачи T2, затем нажмите кнопку **getSelectedDataAsync** в области задач. На рис. 3 показаны данные, выбранные в текстовом окне надстройки **Project OM Test**.
    
    **Рис. 3. Использование надстройки Project OM Test**

    ![Использование приложения OM для тестирования проекта](../images/pj15_CreateSimpleAgave_ProjectOMTest.gif)

8. В столбце **Длительность** выберите ячейку, относящуюся к первой задаче, а затем нажмите кнопку **getSelectedDataAsync** в надстройке **Project OM Test**. Функция **getSelectedDataAsync** приведет к отображению в текстовом поле значения `2 days`. 
    
9. Выберите три ячейки **Длительность** для всех трех задач. Функция **getSelectedDataAsync** возвращает текстовые значения, разделенные точками с запятой, для ячеек, выбранных в различных строках, например: `2 days;4 days;0 days`.
    
    Функция **getSelectedDataAsync** возвращает разделенные запятыми текстовые значения для ячеек, выбранных в одной строке. Например, на рис. 3 для задачи T2 выбрана вся строка. Если выбрана функция **getSelectedDataAsync**, в текстовом поле отобразятся следующие данные: `,Auto Scheduled,T2,4 days,Thu 6/14/12,Tue 6/19/12,1,,<NA>`
    
    В текстовом массиве отображаются пустые значения для столбцов **Indicators** и **Resource Names**, так как оба они не заполнены. Для ячейки **Добавить новый столбец** отображается значение `<NA>`.
    
10. Выберите любую ячейку в строке задачи T2 или всю строку задачи T2, затем нажмите кнопку **getSelectedTaskAsync**. В текстовом поле появится значение GUID задачи, например: `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`. Project сохраняет значение в глобальной переменной **taskGuid** надстройки **Project OM Test**.
    
11. Нажмите кнопку **getTaskAsync**. Если переменная **taskGuid** содержит GUID задачи T2, в текстовом поле появятся сведения о задаче. Значение **ResourceNames** является пустым.
    
    Создайте два локальных ресурса R1 и R2, назначьте каждому из них по 50 % задачи T2 и снова выберите функцию **getTaskAsync**. Результаты в текстовом поле содержат сведения о ресурсе. Если задача включена в синхронизированный список задач SharePoint, в результатах также будет отображаться идентификатор задачи SharePoint.
    


```
  Task name: T2
GUID: {25D3E03B-9A7D-E111-92FC-00155D3BA208}
WSS Id: 0
ResourceNames: R1[50%],R2[50%]
```

12. Нажмите кнопку **Get Task Fields**. Функция **getTaskFields** несколько раз вызывает функцию **getTaskfieldAsync** для имени задачи, индекса, даты начала, длительности, приоритета и примечаний.
    
```
  Name: T2
ID: 2
Start: Thu 6/14/12
Duration: 4d
Priority: 500
Notes: This is a note for task T2. It is only a test note. If it had been a real note, there would be some real information.
```

13. Нажмите кнопку **getWSSUrlAsync**. Если проект принадлежит к одному из указанных ниже видов, то в результатах отображается список задач, URL-адрес и имя.
    
      - Список задач SharePoint, импортированный в Project Server.
    
  - Список задач SharePoint, импортированный в Project профессиональный, а затем снова сохраненный в SharePoint (без использования Project Server).
    
     >**Примечание.** Если Project профессиональный установлен на компьютере Windows Server, а вы хотите сохранить проект в SharePoint, добавьте функцию **возможностей рабочего стола** с помощью **диспетчера сервера**.

    Если проект локальный или вы используете Project профессиональный, чтобы открыть проект, управляемый Project Server, метод **getWSSUrlAsync** возвращает неопределенную ошибку.
    


```
  SharePoint URL: http://ServerName
List name: Test task list
```

14. Нажмите кнопку **Добавить** в разделе **событие TaskSelectionChanged**, что приведет к вызову функции **manageTaskEventHandler** для регистрации события изменения выделения задачи и к возврату `In onComplete function for addHandlerAsync Status: succeeded` для отображения в текстовом поле. Выберите другую задачу; в текстовом поле появится надпись `In task selection changed event handler`, которая является выводом функции обратного вызова для события изменения выделения задачи. Нажмите кнопку **Удалить**, чтобы отменить регистрацию обработчика события.
    
15. Чтобы использовать методы ресурсов, сначала выберите представление, например **Таблица ресурсов**, **Использование ресурсов** или **Форма ресурсов**, затем выберите ресурс в этом представлении. Выберите **getSelectedResourceAsync** для инициализации переменной **resourceGuid**, затем выберите **Получить поля ресурсов**, чтобы несколько раз вызвать **getResourceFieldAsync** для свойств ресурсов. Можно также добавить или удалить обработчик событий для случаев изменения выделения ресурса.
    
```
  Resource name: R1
Cost: $800.00
Standard Rate: $50.00/h
Actual Cost: $0.00
Actual Work: 0h
Units: 100%
```

16. Выберите **getSelectedViewAsync** для отображения типа и имени активного представления. Можно также добавить или удалить обработчик событий для случаев изменения выделения представления. Например, если активным представлением является **Форма ресурсов**, функция **getSelectedViewAsync** отображала бы в текстовом поле следующее значение:
    
```
  View type: 6
Name: Resource Form
```

17. Выберите **Получить поля проекта**, чтобы вызвать несколько раз функцию **getProjectFieldAsync** для различных свойств активного проекта. Если проект открыт в Project Web App, функция **getProjectFieldAsync** может получить URL-адрес экземпляра Project Web App.
    
```
  Project GUID: 9845922E-DAB4-E111-8AF3-00155D3BA208

Start: Tue 6/12/12
Finish: Tue 6/19/12

Currency digits: 2
Currency symbol: $
Symbol position: 0

Project web app URL:
  http://servername/pwa
```

18. Нажмите кнопку **Получить контекстные значения**, чтобы получить свойства документа и приложения, в котором запущена надстройка, считывая свойства объектов **Office.Context.document** и **Office.context.application**. Например, если файл Project1.mpp находится на рабочем столе локального компьютера, URL-адресом документа будет `C:\Users\UserAlias\Desktop\Project1.mpp`. Если MPP-файл находится в библиотеке SharePoint, значением будет URL-адрес документа. Если вы используете Project профессиональный 2013, чтобы открыть проект с именем Project1 в Project Web App, URL-адресом документа будет `<>\Project1`.
    
```
  Document URL:
<>\Project1
Document mode: readWrite
App language: en-US
Display language: en-US
```

19. Можно обновить надстройку после изменения исходного кода, закрыв и перезапустив Project. Недавно использовавшиеся надстройки содержатся на ленте **Проект** в раскрывающемся списке **Надстройки Office**.
    

## <a name="example"></a>Пример


В пакете загрузки SDK для Project 2013 в файле JSOMCall.html содержится полный код, файл JSOM_Sample.js, а также связанные файлы Office.js, Office.debug.js, Project-15.js и Project-15.debug.js. Ниже приведен код, содержащийся в файле JSOMCall.html.


```HTML
<!DOCTYPE html>
<html>
    <head>
        <title>Project OM Sample Code</title>
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

        <script type="text/javascript" src="MicrosoftAjax.js"></script>

        <!-- Use the CDN reference to office.js when deploying your add-in. -->
        <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"></script> -->
        <script type="text/javascript" src="Office.js"></script>
        <script type="text/javascript" src="JSOM_Sample.js"></script>

        <style type="text/css">           
            .button-wide {
                width: 210px;
                margin-top: 2px;
            }
            .button-narrow 
            {
                width: 80px;
                margin-top: 2px;
            }
        </style>
    </head>

    <body>
      <div id="Common_JSOM_API">
        OBJECT MODEL TESTS
        <br /><br />       
        <strong>General method:</strong>
        <br />
        <input id="Button5" class="button-wide" type="button" onclick="getSelectedDataAsync()" 
            value="getSelectedDataAsync" />
      </div>

      <div id="ProjectSpecificTask">
        <br />
        <strong>Project-specific task methods:</strong><br />
        <button class="button-wide" onclick="getSelectedTaskAsync()">getSelectedTaskAsync</button><br />
        <button class="button-wide" onclick="getTaskAsync()">getTaskAsync</button><br />
        <button class="button-wide" onclick="getTaskFields()">Get Task Fields</button><br />
        <button class="button-wide" onclick="getWSSUrlAsync()">getWSSUrlAsync</button>
        <strong>Task selection changed:</strong>
        <button class="button-narrow" onclick="manageTaskEventHandler('addHandlerAsync')">Add</button>
        <button class="button-narrow" onclick="manageTaskEventHandler('removeHandlerAsync')">Remove</button>         
      </div>
<div id="ResourceMethods">
  <br />
  <strong>Resource methods:</strong>
  <button class="button-wide" onclick="getSelectedResourceAsync()">getSelectedResourceAsync</button><br />
  <button class="button-wide" onclick="getResourceFields()">Get Resource Fields</button><br />
  <strong>Resource selection changed:</strong>
  <button class="button-narrow" onclick="manageResourceEventHandler('addHandlerAsync')">Add</button>
  <button class="button-narrow" onclick="manageResourceEventHandler('removeHandlerAsync')">Remove</button>
</div>
<div id="ViewMethods">
  <br />
  <strong>View method:</strong>
  <button class="button-wide" onclick="getSelectedViewAsync()">getSelectedViewAsync</button><br />
  <strong>View selection changed:</strong>
  <button class="button-narrow" onclick="manageViewEventHandler('addHandlerAsync')">Add</button>
  <button class="button-narrow" onclick="manageViewEventHandler('removeHandlerAsync')">Remove</button>         
</div>
<div id="ProjectMethods">
  <br />
  <strong>Project properties:</strong>
  <button class="button-wide" onclick="getProjectFields()">Get Project Fields</button><br />
</div>
<div id="ContextVariables">
  <br />
  <strong>Context properties:</strong>
  <button class="button-wide" onclick="getContextValues()">Get Context Values</button>
</div>

      <br />
      <textarea id="text" rows="10" cols="25">This is the text result.</textarea>
    </body>
</html
```


## <a name="robust-programming"></a>Надежное программирование


На примере надстройки **Project OM Test** показано использование некоторых функций JavaScript в Project 2013, которые включены в файлы Project-15.js и Office.js. Этот пример предназначен исключительно для тестирования, поэтому не содержит комплексной обработки ошибок. Например, если не выбрать ресурс и выполнить функцию **getSelectedResourceAsync**, переменная **resourceGuid** не инициализируется и вызовы **getResourceFieldAsync** возвращают ошибку. При разработке надстройки для производственной среды следует проверять поведение при возникновении определенных ошибок и игнорировать результаты, скрывать ненужные функциональные возможности или уведомлять пользователей о необходимости выбора представления и надлежащего параметра, прежде чем использовать функцию.

В качестве простого примера, вывод ошибки в следующем коде включает переменную **actionMessage**, указывающую действие, которое следует предпринять, чтобы избежать ошибки в функции **getSelectedResourceAsync**.




```js
function logError(errorText) {
    text.value = "Error in " + errorText;
}

function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);
}
// Get the GUID of the selected resource.
function getSelectedResourceAsync() {
    _projDoc.getSelectedResourceAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            text.value = asyncResult.value;
            resourceGuid = asyncResult.value;
        }
        else {
            var actionMessage = "Select a resource before running the getSelectedResourceAsync method.";
            logMethodError("getSelectedResourceAsync", asyncResult.error.name,
                               asyncResult.error.message, actionMessage);
        }
    });
}
```

Разрабатывать надстройки проще в Visual Studio 2015, где можно задавать точки останова для отладки кода JavaScript и быстро внедрить общие процедуры обработки ошибок. Например, образец **HelloProject_OData** в пакете SDK для Project 2013 включает файл SurfaceErrors.js, использующий библиотеку JQuery для отображения всплывающего сообщения об ошибке. На рисунке 4 показано сообщение об ошибке во всплывающем уведомлении. Этот образец также включает файл Office-vsdoc.js, позволяющий использовать IntelliSense для функций JavaScript в файлах Office.js и Project-15.js.

Приведенный ниже код, который содержится в файле SurfaceErrors.js, включает функцию **throwError**, создающую объект **Toast**.


```js
/*
 * Show error messages in a "toast" notification.
 */

// Throws a custom defined error.
function throwError(errTitle, errMessage) {
    try {
        // Define and throw a custom error.
        var customError = { name: errTitle, message: errMessage }
        throw customError;
    }
    catch (err) {
        // Catch the error and display it to the user.
        Toast.showToast(err.name, err.message);
    }
}

// Add a dynamically-created div "toast" for displaying errors to the user.
var Toast = {

    Toast: "divToast",
    Close: "btnClose",
    Notice: "lblNotice",
    Output: "lblOutput",

    // Show the toast with the specified information.
    showToast: function (title, message) {

        if (document.getElementById(this.Toast) == null) {
            this.createToast();
        }

        document.getElementById(this.Notice).innerText = title;
        document.getElementById(this.Output).innerText = message;

        $("#" + this.Toast).hide();
        $("#" + this.Toast).show("slow");
    },

    // Create the display for the toast.
    createToast: function () {
        var divToast;
        var lblClose;
        var btnClose;
        var divOutput;
        var lblOutput;
        var lblNotice;

        // Create the container div.
        divToast = document.createElement("div");
        var toastStyle = "background-color:rgba(220, 220, 128, 0.80);" +
            "position:absolute;" +
            "bottom:0px;" +
            "width:90%;" +
            "text-align:center;" +
            "font-size:11pt;";
        divToast.setAttribute("style", toastStyle);
        divToast.setAttribute("id", this.Toast);

        // Create the close button.
        lblClose = document.createElement("div");
        lblClose.setAttribute("id", this.Close);
        var btnStyle = "text-align:right;" +
            "padding-right:10px;" +
            "font-size:10pt;" +
            "cursor:default";
        lblClose.setAttribute("style", btnStyle);
        lblClose.appendChild(document.createTextNode("CLOSE "));

        btnClose = document.createElement("span");
        btnClose.setAttribute("style", "cursor:pointer;");
        btnClose.setAttribute("onclick", "Toast.close()");
        btnClose.innerText = "X";
        lblClose.appendChild(btnClose);

        // Create the div to contain the toast title and message.
        divOutput = document.createElement("div");
        divOutput.setAttribute("id", "divOutput");
        var outputStyle = "margin-top:0px;";
        divOutput.setAttribute("style", outputStyle);

        lblNotice = document.createElement("span");
        lblNotice.setAttribute("id", this.Notice);
        var labelStyle = "font-weight:bold;margin-top:0px;";
        lblNotice.setAttribute("style", labelStyle);

        lblOutput = document.createElement("span");
        lblOutput.setAttribute("id", this.Output);

        // Add the child nodes to the toast div.
        divOutput.appendChild(lblNotice);
        divOutput.appendChild(document.createElement("br"));
        divOutput.appendChild(lblOutput);
        divToast.appendChild(lblClose);
        divToast.appendChild(divOutput);

        // Add the toast div to the document body.
        document.body.appendChild(divToast);
    },

    // Close the toast.
    close: function () {
        $("#" + this.Toast).hide("slow");
    }
}
```

Чтобы использовать функцию **throwError**, включите библиотеку JQuery и сценарий SurfaceErrors.js в файл JSOMCall.html, а также добавьте вызов **throwError** в другие функции JavaScript, такие как **logMethodError**.


 >**Примечание.** Перед развертыванием надстройки измените ссылку office.js и ссылку jQuery на ссылку сети доставки содержимого (CDN). Ссылка CDN предоставляет самую последнюю версию и обеспечивает оптимальную производительность.




```HTML
<!DOCTYPE html>
<html>
<head>
    <title>Project OM Sample Code</title>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

    <script type="text/javascript" src="MicrosoftAjax.js"></script>

    <!-- Use the CDN reference to Office.js and jQuery when deploying your add-in. -->
    <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"></script> -->
    <script type="text/javascript" src="Office.js"></script>
    <script type="text/javascript" src="http://ajax.microsoft.com/ajax/jQuery/jquery-1.9.0.min.js"></script>

    <script type="text/javascript" src="JSOM_Sample.js"></script>
    <script type="text/javascript" src="SurfaceErrors.js"></script>

    <!-- . . . INVALID USE OF SYMBOLS
</head>

```




```js
function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);

    throwError(methodName + " error", actionMessage);
}
```


**Рис. 4. Функции в файле SurfaceErrors.js могут показывать всплывающее уведомление**

![Использование процедур SurfaceError для отображения ошибки](../images/pj15_CreateSimpleAgave_SurfaceError.gif)


## <a name="additional-resources"></a>Дополнительные ресурсы



- [Надстройки области задач для Project](../project/project-add-ins.md)
    
- [Общие сведения об API JavaScript для надстроек](../develop/understanding-the-javascript-api-for-office.md)
    
- [API JavaScript для надстроек Office](http://dev.office.com/reference/add-ins/javascript-api-for-office)

- [Справка по схеме для манифестов надстроек Office (версия 1.1)](../overview/add-in-manifests.md)     
    
- [Загрузка пакета SDK для Project 2013](https://www.microsoft.com/en-us/download/details.aspx?id=30435%20)
    
