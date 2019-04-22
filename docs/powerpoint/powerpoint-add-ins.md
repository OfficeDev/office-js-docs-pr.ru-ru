---
title: Надстройки PowerPoint
description: ''
ms.date: 04/15/2019
localization_priority: Priority
ms.openlocfilehash: 6e518d0bfd37291e39ee17e96ded8debb183c19f
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914230"
---
# <a name="powerpoint-add-ins"></a>Надстройки PowerPoint

С помощью надстроек PowerPoint можно создавать удобные решения, подходящие для использования в презентациях на различных платформах, таких как Windows, iOS, Office Online и Mac. Можно создать два типа надстроек PowerPoint:

- **Контентные надстройки** позволяют добавлять динамический контент HTML5 в презентации. Например, ознакомьтесь с надстройкой [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/WA104380117), с помощью которой можно добавить интерактивные схемы LucidChart в набор слайдов.

- **Надстройки области задач** позволяют добавлять справочные сведения или данные в презентацию с помощью службы. Например, используя надстройку [Pixton Comic Characters](https://appsource.microsoft.com/product/office/WA104380907), вы можете вставить профессиональные фотографии в свою презентацию. 

## <a name="powerpoint-add-in-scenarios"></a>Сценарии надстроек PowerPoint

В приведенных в этой статье примерах кода показаны основные задачи по разработке надстроек для PowerPoint. Обратите внимание на следующее:

- При отображении сведений эти примеры используют функцию `app.showNotification`, включенную в шаблоны проектов надстроек Office в Visual Studio. Если для разработки надстройки вы не используете Visual Studio, замените функцию `showNotification` собственным кодом. 

- Некоторые из этих примеров также используют объект `Globals`, объявленный за пределами указанных функций как `var Globals = {activeViewHandler:0, firstSlideId:0};`.

- Для использования этих примеров необходимо, чтобы проект надстройки [ссылался на библиотеку Office.js 1.1 или более поздней версии](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>Определение активного представления презентации и обработка события ActiveViewChanged

При создании контентной надстройки вам понадобится получить активное представление презентации, а также обработать событие `ActiveViewChanged` в рамках обработчика событий `Office.Initialize`.

> [!NOTE]
> В PowerPoint Online не удастся запустить событие [Document.ActiveViewChanged](/javascript/api/office/office.document), поскольку режим показа слайдов обрабатывается как новый сеанс. В этом случае надстройке необходимо получить активное представление по загрузке, как показано в примере кода ниже.

В представленном ниже примере кода:

- Функция `getActiveFileView` вызывает метод [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-), который возвращает текущее представление презентации: "edit" (представления, в которых можно редактировать слайды, например  **Обычный режим** или **Режим структуры**) или "read" (**Показ слайдов** или **Режим чтения**).

- Функция `registerActiveViewChanged` вызывает метод [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) для регистрации обработчика для события [Document.ActiveViewChanged](/javascript/api/office/office.document).


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

В приведенном ниже примере кода функция `getSelectedRange` вызывает метод [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-), чтобы получить возвращаемый свойством `asyncResult.value` объект JSON, содержащий массив с именем **slides**. Массив **slides** содержит идентификаторы, заголовки и индексы выбранного диапазона слайдов (или текущего слайда, если не выбрано несколько слайдов). Кроме того, он сохраняет идентификатор первого слайда в выбранном диапазоне в глобальной переменной.

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

В приведенном ниже примере кода функция `goToFirstSlide` вызывает метод [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) для перехода к первому слайду, который был определен показанной ранее функцией `getSelectedRange`.

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

В приведенном ниже примере кода функция `goToSlideByIndex` вызывает метод **Document.goToByIdAsync** для перехода к следующему слайду в презентации.

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

В приведенном ниже примере кода функция `getFileUrl` вызывает метод [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-), чтобы получить URL-адрес файла презентации.

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

## <a name="create-a-presentation"></a>Создание презентации

Ваша надстройка может создать новую презентацию, отдельную от экземпляра PowerPoint, в котором в настоящее время работает надстройка. Для этой цели в пространстве имен PowerPoint есть метод `createPresentation`. При вызове этого метода сразу открывается и отображается новая презентация в новом экземпляре программы PowerPoint. Ваша надстройка остается открытой и запущенной в предыдущей презентации.

```js
PowerPoint.createPresentation();
```

С помощью метода `createPresentation` также можно создать копию существующей презентации. Метод принимает в качестве необязательного параметра строковое представление PPTX-файла в кодировке base64. Полученная презентация будет копией этого файла, предполагая, что строковый аргумент является допустимым PPTX-файлом. Преобразование файла в нужную строку в кодировке base64 можно выполнить с помощью класса [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader), как показано в приведенном ниже примере.

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = function (event) {
    // strip off the metadata before the base64-encoded string
    var startIndex = event.target.result.indexOf("base64,");
    var copyBase64 = event.target.result.substr(startIndex + 7);

    PowerPoint.createPresentation(copyBase64);
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="see-also"></a>См. также

- [Примеры кода PowerPoint](https://developer.microsoft.com/office/gallery/?filterBy=Samples,PowerPoint)
- [Сохранение состояния надстройки и параметров документа для надстроек области задач и контентных надстроек](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [Чтение и запись данных при активном выделении фрагмента в документе или электронной таблице](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [Получение всего документа из надстройки для PowerPoint или Word](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [Использование тем документов в надстройках PowerPoint](use-document-themes-in-your-powerpoint-add-ins.md)