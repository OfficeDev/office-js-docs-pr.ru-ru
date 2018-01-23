# <a name="powerpoint-add-ins"></a>Надстройки PowerPoint

С помощью надстроек PowerPoint можно создавать удобные решения, подходящие для использования в презентациях на различных платформах, таких как Windows, iOS, Office Online и Mac. Можно создать два типа надстроек:

- **Контентные надстройки** позволяют добавлять динамический контент HTML5 в презентации. Например, ознакомьтесь с надстройкой [LucidChart Diagrams for PowerPoint](https://store.office.com/en-us/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false), с помощью которой можно добавить интерактивные схемы LucidChart в набор слайдов.
- **Надстройки области задач** позволяют добавлять справочные сведения или данные в слайд с помощью службы. Например, используя надстройку [Shutterstock Images](https://store.office.com/en-us/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false), вы можете вставить профессиональные фотографии в свою презентацию. 

>
  **Примечание.** Если вы планируете [публиковать](../publish/publish.md) надстройку в Магазине Office, она должна соответствовать [политикам проверки Магазина Office](https://msdn.microsoft.com/en-us/library/jj220035.aspx), чтобы пройти проверку. Например, работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) и на [странице с информацией о доступности и ведущих приложениях для надстроек Office](https://dev.office.com/add-in-availability).

## <a name="powerpoint-add-in-scenarios"></a>Сценарии надстроек PowerPoint

В приведенных в статье примерах кода показаны основные задачи по разработке контентных надстроек для PowerPoint. 

При отображении сведений эти примеры зависят от функции `app.showNotification`, включенной в шаблоны проектов надстроек Office в Visual Studio. Если для разработки надстройки вы не используете Visual Studio, замените функцию `showNotification` собственным кодом. Некоторые из этих примеров также зависят от объекта `globals`, объявленного за пределами указанных функций: `var globals = {activeViewHandler:0, firstSlideId:0};`

Эти примеры кода требуют, чтобы проект [ссылался на библиотеку Office.js 1.1 или более поздней версии](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>Определение активного представления презентации и обработка события ActiveViewChanged

При создании контентной надстройки вам понадобится получить активное представление презентации, а также обработать событие ActiveViewChanged в рамках обработчика событий Office.Initialize.


- Функция `getActiveFileView` вызывает метод [Document.getActiveViewAsync](http://dev.office.com/reference/add-ins/shared/document.getactiveviewasync), который возвращает текущее представление презентации: "edit" (представления, в которых можно редактировать слайды, например  **Обычный режим** или **Режим структуры**) или "read" (**Показ слайдов** или **Режим чтения**).


- Функция `registerActiveViewChanged` вызывает метод [addHandlerAsync](http://dev.office.com/reference/add-ins/shared/document.addhandlerasync) для регистрации обработчика для события [Document.ActiveViewChanged](http://dev.office.com/reference/add-ins/shared/document.activeviewchanged). 
> Примечание. В PowerPoint Online не удастся запустить событие [Document.ActiveViewChanged](http://dev.office.com/reference/add-ins/shared/document.activeviewchanged), поскольку режим показа слайдов обрабатывается как новый сеанс. В этом случае надстройке необходимо получить активное представление по загрузке, как указано ниже.



```js

//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    var currentView = getActiveFileView();

    //register for the active view changed handler
    registerActiveViewChanged();

    //render the content based off of the currentView
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });

}


function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
           app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```
    

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>Переход к определенному слайду презентации

Функция `getSelectedRange` вызывает метод [Document.getSelectedDataAsync](http://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), чтобы получить объект JSON, возвращаемый свойством  `asyncResult.value` и который включает в себя массив с именем slides, содержащий идентификаторы, заголовки и индексы выбранного диапазона слайдов (или текущего слайда). Кроме того, он сохраняет идентификатор первого слайда в выбранном диапазоне в глобальной переменной.


```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

Функция `goToFirstSlide` вызывает метод [Document.goToByIdAsync](http://dev.office.com/reference/add-ins/shared/document.gotobyidasync) для перехода к идентификатору первого слайда, сохраненному описанной выше функцией `getSelectedRange`.




```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## <a name="navigate-between-slides-in-the-presentation"></a>Переход между слайдами презентации

Функция `goToSlideByIndex` вызывает метод **Document.goToByIdAsync** для перехода к следующему слайду в презентации.


```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="get-the-url-of-the-presentation"></a>Получение URL-адреса презентации

Функция `getFileUrl` вызывает метод [Document.getFileProperties](http://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync), чтобы получить URL-адрес файла презентации.


```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```



## <a name="additional-resources"></a>Дополнительные ресурсы
- [Примеры кода PowerPoint](https://dev.office.com/code-samples#?filters=powerpoint)

- [Сохранение состояния надстройки и параметров документа для контентных надстроек и надстроек области задач](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [Считывание и запись данных в активное выделение документа или таблицы](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [Получение всего документа из надстройки для PowerPoint или Word](../develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [Использование тем документов в надстройках PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
