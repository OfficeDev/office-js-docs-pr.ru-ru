---
title: Надстройки PowerPoint
description: Узнайте, как использовать надстройки PowerPoint для создания удобных решений для презентаций на разных платформах, включая Windows, iPad, Mac и в браузере.
ms.date: 07/18/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 54e840c62a78af289a1c401ceb0961d610292f99
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889592"
---
# <a name="powerpoint-add-ins"></a>Надстройки PowerPoint

С помощью надстроек PowerPoint можно создавать удобные решения, подходящие для использования в презентациях на различных платформах, таких как Windows, iPad, Mac и в браузере. Можно создать два типа надстроек PowerPoint:

- **Контентные надстройки** позволяют добавлять динамический контент HTML5 в презентации. Например, ознакомьтесь с надстройкой [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/wa104380117), с помощью которой можно добавить интерактивные схемы LucidChart в набор слайдов.

- **Надстройки области задач** позволяют добавлять справочные сведения или данные в презентацию с помощью службы. Например, используя надстройку [Pexels - Бесплатные стоковые фото](https://appsource.microsoft.com/product/office/wa104379997), вы можете вставить профессиональные фотографии в свою презентацию.

## <a name="powerpoint-add-in-scenarios"></a>Сценарии надстроек PowerPoint

В приведенных в этой статье примерах кода показаны основные задачи по разработке надстроек для PowerPoint. Обратите внимание на перечисленные ниже аспекты.

- При отображении сведений эти примеры используют функцию `app.showNotification`, включенной в шаблоны проектов надстроек Office в Visual Studio. Если для разработки надстройки вы не используете Visual Studio, замените функцию `showNotification` собственным кодом.

- Некоторые из этих примеров также используют объект `Globals`, объявленный за пределами указанных функций как `var Globals = {activeViewHandler:0, firstSlideId:0};`.

- Для использования этих примеров необходимо, чтобы проект надстройки [ссылался на библиотеку Office.js 1.1 или более поздней версии](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>Определение активного представления презентации и обработка события ActiveViewChanged

При создании контентной надстройки вам понадобится получить активное представление презентации, а также обработать событие `ActiveViewChanged` в рамках обработчика событий `Office.Initialize`.

> [!NOTE]
> В PowerPoint в Интернете не удастся запустить событие [Document.ActiveViewChanged](/javascript/api/office/office.document), поскольку режим показа слайдов обрабатывается как новый сеанс. В этом случае надстройке необходимо получить активное представление по загрузке, как показано в представленном ниже примере кода.

В представленном ниже примере кода:

- Функция `getActiveFileView` вызывает метод [Document.getActiveViewAsync](/javascript/api/office/office.document#office-office-document-getactiveviewasync-member(1)), который возвращает текущее представление презентации: "edit" (представления, в которых можно редактировать слайды, например  **Обычный режим** или **Режим структуры**) или "read" (**Показ слайдов** или **Режим чтения**).

- Функция `registerActiveViewChanged` вызывает метод [addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) для регистрации обработчика для события [Document.ActiveViewChanged](/javascript/api/office/office.document).

```js
//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    const currentView = getActiveFileView();

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

В представленном ниже примере кода функция `getSelectedRange` вызывает метод [Document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)), чтобы получить объект JSON, возвращаемый свойством `asyncResult.value`, который включает в себя массив с именем `slides`. Массив `slides` содержит идентификаторы, заголовки и индексы выбранного диапазона слайдов (или текущего слайда, если не выбрано несколько слайдов). Кроме того, он сохраняет идентификатор первого слайда в выбранном диапазоне в глобальной переменной.

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

В приведенном ниже примере кода функция `goToFirstSlide` вызывает метод [Document.goToByIdAsync](/javascript/api/office/office.document#office-office-document-gotobyidasync-member(1)) для перехода к первому слайду, который был определен показанной ранее функцией `getSelectedRange`.

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

В следующем примере кода функция `goToSlideByIndex` вызывает метод `Document.goToByIdAsync` для перехода к следующему слайду презентации.

```js
function goToSlideByIndex() {
    const goToFirst = Office.Index.First;
    const goToLast = Office.Index.Last;
    const goToPrevious = Office.Index.Previous;
    const goToNext = Office.Index.Next;

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

В приведенном ниже примере кода функция `getFileUrl` вызывает метод [Document.getFileProperties](/javascript/api/office/office.document#office-office-document-getfilepropertiesasync-member(1)), чтобы получить URL-адрес файла презентации.

```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        const fileUrl = asyncResult.value.url;
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
const myFile = document.getElementById("file");
const reader = new FileReader();

reader.onload = function (event) {
    // strip off the metadata before the base64-encoded string
    const startIndex = reader.result.toString().indexOf("base64,");
    const copyBase64 = reader.result.toString().substr(startIndex + 7);

    PowerPoint.createPresentation(copyBase64);
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="see-also"></a>См. также

- [Разработка надстроек Office](../develop/develop-overview.md)
- [Сведения о программе для разработчиков Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Примеры кода PowerPoint](https://developer.microsoft.com/office/gallery/?filterBy=Samples,PowerPoint)
- [Сохранение состояния надстройки и параметров документа для надстроек области задач и контентных надстроек](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [Чтение и запись данных при активном выделении фрагмента в документе или электронной таблице](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [Получение всего документа из надстройки для PowerPoint или Word](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [Использование тем документов в надстройках PowerPoint](use-document-themes-in-your-powerpoint-add-ins.md)
