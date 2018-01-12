# <a name="onenote-javascript-api-programming-overview"></a>Обзор создания кода с помощью API JavaScript для OneNote

В OneNote представлен API JavaScript для надстроек OneNote Online. Вы можете создавать надстройки области задач, контентные надстройки и команды надстроек, которые взаимодействуют с объектами OneNote и подключаются к веб-службам или другим веб-ресурсам.

>
  **Примечание.** Если вы планируете [публиковать](../publish/publish.md) надстройку в Магазине Office, она должна соответствовать [политикам проверки Магазина Office](https://msdn.microsoft.com/en-us/library/jj220035.aspx), чтобы пройти проверку. Например, работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) и на [странице с информацией о доступности и ведущих приложениях для надстроек Office](https://dev.office.com/add-in-availability).

## <a name="components-of-an-office-add-in"></a>Компоненты надстройки Office

Надстройки состоят из двух указанных ниже основных компонентов.

- **Веб-приложение**, состоящее из веб-страницы и необходимых JavaScript-, CSS- или других файлов. Эти файлы можно разместить на веб-сервере или в службе веб-хостинга, например в Microsoft Azure. В OneNote Online веб-приложение отображается в элементе управления браузера или в iFrame.
    
- **Манифест в формате XML**, в котором указан URL-адрес веб-страницы надстройки и все требования, необходимые для получения доступа, параметры и возможности для надстройки. Этот файл хранится на клиентском компьютере. Для надстроек OneNote используется такой же формат [манифеста](https://dev.office.com/docs/add-ins/overview/add-in-manifests), что и для других надстроек Office.

**Надстройка Office = манифест + веб-страница**

![Надстройка Office состоит из манифеста и веб-страницы](../../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>Использование API JavaScript

Для доступа к API JavaScript надстройки используют контекст среды выполнения ведущего приложения. API состоит из двух указанных ниже уровней. 

- **Многофункциональный API** для связанных с OneNote операций, доступ к которому осуществляется с помощью объекта **Application**.
- **Стандартный API**, используемый приложениями Office, доступ к которому осуществляется с помощью объекта **Document**.

### <a name="accessing-the-rich-api-through-the-application-object"></a>Доступ к многофункциональному API с помощью объекта *Application*

Для доступа к объектам OneNote, например к объектам **Notebook**, **Section** и **Page**, используйте объект **Application**. С помощью многофункциональных API вы можете запустить пакетные операции на прокси-объектах. Основной процесс выглядит примерно так, как указано ниже. 

1. Получение экземпляра приложения из контекста.

2. Создание прокси-объекта, представляющего объект OneNote, с которым вам необходимо работать. Для синхронного взаимодействия с прокси-объектами можно считывать и записывать их свойства и вызывать имеющиеся в них методы. 

3. Вызовите метод **load** прокси, чтобы указать для него значения свойств, указанные в параметре. Этот вызов будет добавлен в очередь команд.

    > **Примечание.** Выполняемые методами вызовы API (например, `context.application.getActiveSection().pages;`), также добавляются в очередь.

4. Чтобы запустить все поставленные в очередь команды в том порядке, в котором они были добавлены в очередь, вызовите метод **context.sync**. Этот метод синхронизирует состояния выполняющихся сценариев и реальных объектов, а также получает свойства загруженных объектов OneNote, которые необходимо использовать в сценарии. Вы можете использовать возвращенный объект обещания для связывания дополнительных действий в цепочку.

Например: 

```
    function getPagesInSection() {
        OneNote.run(function (context) {
            
            // Get the pages in the current section.
            var pages = context.application.getActiveSection().pages;
            
            // Queue a command to load the id and title for each page.            
            pages.load('id,title');
            
            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    
                    // Read the id and title of each page. 
                    $.each(pages.items, function(index, page) {
                        var pageId = page.id;
                        var pageTitle = page.title;
                        console.log(pageTitle + ': ' + pageId); 
                    });
                })
                .catch(function (error) {
                    app.showNotification("Error: " + error);
                    console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                });
        });
    }
```

Сведения о поддерживаемых объектах и операциях OneNote см. в [справочнике по API](../../reference/onenote/onenote-add-ins-javascript-reference.md).

### <a name="accessing-the-common-api-through-the-document-object"></a>Получение доступа к стандартному API с помощью объекта *Document*

Для доступа к стандартному API, например к методам **getSelectedDataAsync** и [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), используйте объект [Document](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync). 

Например:  

```
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```
Надстройки OneNote поддерживают только указанные ниже стандартные API.

| API | Примечания |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142294.aspx) | Только **Office.CoercionType.Text** и **Office.CoercionType.Matrix** |
| [Office.context.document.setSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142145.aspx) | Только **Office.CoercionType.Text**, **Office.CoercionType.Image** и **Office.CoercionType.Html** | 
| [var mySetting = Office.context.document.settings.get(имя);](https://msdn.microsoft.com/en-us/library/office/fp142180.aspx) | Параметры поддерживаются только контентными надстройками | 
| [Office.context.document.settings.set(имя, значение);](https://msdn.microsoft.com/en-us/library/office/fp161063.aspx) | Параметры поддерживаются только контентными надстройками | 
| [Office.EventType.DocumentSelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

Обычно стандартный API следует использовать только тогда, когда необходимые возможности не поддерживаются в многофункциональном API. Дополнительные сведения об использовании стандартного API см. в [документации](https://dev.office.com/docs/add-ins/overview/office-add-ins) и [справочнике](https://dev.office.com/reference/add-ins/javascript-api-for-office) по надстройкам Office.


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a>Схема объектной модели OneNote 
На схеме ниже показаны возможности, которые на данный момент доступны в API JavaScript для OneNote .

  ![Схема объектной модели OneNote](../../images/onenote-om.png)


## <a name="additional-resources"></a>Дополнительные ресурсы

- [Создание первой надстройки OneNote](onenote-add-ins-getting-started.md)
- [Справочник по API JavaScript для OneNote](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Пример надстройки Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Обзор платформы надстроек Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
