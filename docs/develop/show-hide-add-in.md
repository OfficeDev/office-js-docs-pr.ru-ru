---
title: Показать или скрыть области задач надстройки Office
description: Узнайте, как программным образом скрывать или показывать пользовательский интерфейс надстройки во время ее непрерывной работы.
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 20db609a3a6ded5624391f705dab1ad6b8f6e043
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789250"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Показать или скрыть области задач надстройки Office

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Вы можете отдемонстрировать области задач надстройки Office, вызывая `Office.addin.showAsTaskpane()` функцию.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

В предыдущем коде предполагается, что существует таблица Excel **с именем CurrentQuarterSales.** Надстройка делает области задач видимыми при активации этого таблицы. Этот метод `onCurrentQuarter` является обработом события [Office.Worksheet.onActivated,](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) зарегистрированного для этого таблицы.

Вы также можете скрыть области задач, вызывая `Office.addin.hide()` функцию.

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

Предыдущий код — это обработец, зарегистрированный для [события Office.Worksheet.onDeactivated.](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated)

## <a name="additional-details-on-showing-the-task-pane"></a>Дополнительные сведения о от показании области задач

При вызове Office отобразит в области задач файл, который назначен в качестве значения ресурса `Office.addin.showAsTaskpane()` ( `resid` ) области задач. Это `resid` значение можно навести или  изменить, открыв файлmanifest.xmlи выявив его `<SourceLocation>` внутри `<Action xsi:type="ShowTaskpane">` элемента.
[(Дополнительные сведения см.](configure-your-add-in-to-use-a-shared-runtime.md) в настройках надстройки Office для использования общей времени работы.)

Так `Office.addin.showAsTaskpane()` как это асинхронный метод, код будет работать до завершения работы функции. Дождись завершения с помощью ключевого слова или метода в зависимости от используемого синтаксиса `await` `then()` JavaScript.

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>Настройка надстройки для использования общей времени работы

Для использования этих `showAsTaskpane()` `hide()` методов надстройка должна использовать общую времени работы. Дополнительные сведения см. в настройках [надстройки Office для использования общей времени работы.](configure-your-add-in-to-use-a-shared-runtime.md)

## <a name="preservation-of-state-and-event-listeners"></a>Сохранение прослушивателей состояния и событий

Методы `hide()` `showAsTaskpane()` и методы изменяют только *видимость* области задач. Они не выгружают и не перезагружают его (или повторно ициализируют его состояние).

Рассмотрим следующий сценарий: в области задач есть вкладки. Вкладка **"Главная"** открывается при первом запуске надстройки. Предположим, что  пользователь открывает вкладку "Параметры", а затем код в области задач вызывается в ответ `hide()` на какое-либо событие. Тем не менее позже код `showAsTaskpane()` вызывается в ответ на другое событие. В области задач снова появится  вкладка "Параметры".

![Снимок экрана области задач с четырьмя вкладками "Главная", "Параметры", "Избранное" и "Учетные записи".](../images/TaskpaneWithTabs.png)

Кроме того, все прослушиватели событий, зарегистрированные в области задач, продолжают работать даже при скрытии области задач.

Рассмотрим следующий сценарий: в области задач есть зарегистрированный обработитель для Excel и события для `Worksheet.onActivated` `Worksheet.onDeactivated` листа **Sheet1.** Активированный обработок приводит к появления зеленой точки в области задач. Деактивированный обработок включает красный цвет точки (которое является состоянием по умолчанию). Предположим, что код `hide()` вызывается, **когда Лист1** не активирован, а точка является красной. Несмотря на то что области задач скрыты, **лист1** активируется. Последующие вызовы кода `showAsTaskpane()` в ответ на какое-либо событие. Когда откроется области задач, точка будет зеленой, так как прослушиватели событий и обработчики запустились, даже если она была скрыта.

## <a name="handle-the-visibility-changed-event"></a>Обработка события изменения видимости

Когда код изменяет видимость области задач с помощью или `showAsTaskpane()` `hide()` , Office активирует `VisibilityModeChanged` событие. Это событие может оказаться полезным. Например, предположим, что в области задач отображается список всех листов в книге. Если новый лист добавляется при скрытии области задач, то, чтобы сделать ее видимой, она сама по себе не добавит новое имя в список. Но ваш код может реагировать на событие, чтобы перезагрузить Worksheet.name всех таблиц в коллекции `VisibilityModeChanged` [Workbook.worksheets,](/javascript/api/excel/excel.workbook#worksheets) как показано в примере кода ниже. [](/javascript/api/excel/excel.worksheet#name)

Чтобы зарегистрировать обработатель для события, не используйте метод add handler, как в большинстве контекстов JavaScript для Office. Вместо этого существует специальная функция, которой вы передаете обработчивую функцию: [Office.addin.onVisibilityModeChanged.](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-) Ниже приведен пример. Обратите `args.visibilityMode` внимание, что свойство имеет тип [VisibilityMode.](/javascript/api/office/office.visibilitymode)

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

Функция возвращает другую функцию, которая *отрегистрировать* обработитель. Вот простой, но не надежный пример:

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

Метод является асинхронным и возвращает обещание, которое означает, что ваш код должен ожидать выполнения обещания, прежде чем он сможет вызвать обработитель `onVisibilityModeChanged` дерегистрации. 

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

Функция дерегистрации также является асинхронной и возвращает обещание. Таким образом, если у вас есть код, который не должен запускаться до завершения регистрации, следует дождаться обещания, возвращенного функцией дерегистрации.

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>См. также

- [Настройка надстройки Office для использования общей времени работы JavaScript](configure-your-add-in-to-use-a-shared-runtime.md)
- [Запуск кода в надстройки Office при запуске документа](run-code-on-document-open.md)
