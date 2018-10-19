---
title: Надстройки PowerPoint
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 390497e74d4dc52b9d400f242850ab72bdb0eabc
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640080"
---
# <a name="powerpoint-add-ins"></a>Надстройки PowerPoint

С помощью надстроек PowerPoint можно создавать удобные решения, подходящие для использования в презентациях на различных платформах, таких как Windows, iOS, Office Online и Mac. Можно создать два типа надстроек PowerPoint:

- **Контентные надстройки** позволяют добавлять динамический контент HTML5 в презентации. Например, ознакомьтесь с надстройкой [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false), с помощью которой можно добавить интерактивные схемы LucidChart в набор слайдов.

- **Надстройки области задач** позволяют добавлять справочные сведения или данные в слайд с помощью службы. Например, используя надстройку [Shutterstock  Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false), вы можете вставить профессиональные фотографии в свою презентацию. 

## <a name="powerpoint-add-in-scenarios"></a>Сценарии надстроек PowerPoint

В примерах кода, приведенных в этой статье, показаны основные задачи по разработке надстроек для PowerPoint. Обратите внимание!

- При отображении сведений эти примеры зависят от функции `app.showNotification`, включенной в шаблоны проектов надстроек Office в Visual Studio. Если для разработки надстройки вы не используете Visual Studio, замените функцию `showNotification` собственным кодом. 

- Некоторые из этих примеров также зависят от объекта `Globals`, объявленного за пределами указанных функций:   `var Globals = {activeViewHandler:0, firstSlideId:0};`

- Для использования этих примеров необходимо, чтобы проект надстройки [ссылался на библиотеку Office.js 1.1 или более поздней версии](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>Определение активного представления презентации и обработка события ActiveViewChanged

При создании контентной надстройки вам понадобится получить активное представление презентации, а также обработать событие `ActiveViewChanged`  в рамках обработчика событий `Office.Initialize`. 

> [!NOTE]
> В PowerPoint Online не удастся запустить событие [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js), поскольку режим показа слайдов обрабатывается как новый сеанс. В этом случае надстройке необходимо получить активное представление по загрузке, как показано в следующем примере кода.

В следующем примере кода:

- Функция `getActiveFileView`  вызывает метод [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-) , который возвращает текущее представление презентации: "edit" (представления, в которых можно редактировать слайды, например **Обычный режим**  или **Режим структуры**) или "read" (**Показ слайдов** или **Режим чтения**).

- Функция `registerActiveViewChanged` вызывает метод [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) для регистрации обработчика для события [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js). 


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

В следующем примере кода функция `getSelectedRange` вызывает метод [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-), чтобы получить объект JSON, возвращаемый свойством `asyncResult.value`, и который включает в себя массив с именем **slides**. Массив **slides** содержит идентификаторы, заголовки и индексы выбранного диапазона слайдов (или текущего слайда, если не выбрано несколько слайдов). Кроме того, он сохраняет идентификатор первого слайда в выбранном диапазоне в глобальной переменной.

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

В следующем примере кода функция `goToFirstSlide` вызывает метод [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) для перехода к первому слайду, который был идентифицирован показанной ранее функцией `getSelectedRange`.

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

В следующем примере кода функция `goToSlideByIndex` вызывает метод **Document.goToByIdAsync** для перехода к следующему слайду в презентации.

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

В следующем примере кода `getFileUrl` функция вызывает метод [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) , чтобы получить URL-адрес файла презентации.

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



## <a name="see-also"></a>См. также
- [Примеры кода PowerPoint](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [Сохранение состояния надстройки и параметров документа для надстроек области задач и контентных надстроек](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [Чтение и запись данных при активном выделении фрагмента в документе или электронной таблице](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [Получение всего документа из надстройки для PowerPoint или Word](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [Использование тем документов в надстройках PowerPoint](use-document-themes-in-your-powerpoint-add-ins.md)
    
