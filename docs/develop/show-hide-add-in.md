---
title: Отображение и скрытие области задач надстройки Office
description: Узнайте, как программным образом скрыть или показать пользовательский интерфейс надстройки во время ее непрерывного работы.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: eb540b9b39870a02343e5a42fdbe3cc9cbd78f01
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150941"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Отображение и скрытие области задач надстройки Office

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Вы можете показать области задач Office надстройки, позвонив в `Office.addin.showAsTaskpane()` функцию.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

Предыдущий код предполагает сценарий, в котором имеется Excel с именем **CurrentQuarterSales.** Надстройка сделает области задач видимыми при активации этого таблицы. Метод `onCurrentQuarter` обработник для [Office. Событие Worksheet.onActivated,](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onActivated) зарегистрированное для таблицы.

Кроме того, можно скрыть области задач, позвонив в `Office.addin.hide()` эту функцию.

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

Предыдущий код — обработник, зарегистрированный для [Office. Событие Worksheet.onDeactivated.](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onDeactivated)

## <a name="additional-details-on-showing-the-task-pane"></a>Дополнительные сведения о показе области задач

При вызове Office будет отображаться в области задач файл, который назначен в качестве ID ресурса () значения области `Office.addin.showAsTaskpane()` `resid` задач. Это значение может быть назначено или изменено путем открытияmanifest.xmlфайла и размещения `resid`  `<SourceLocation>` внутри `<Action xsi:type="ShowTaskpane">` элемента.
(См. раздел Настройка Office надстройки [для использования](configure-your-add-in-to-use-a-shared-runtime.md) общего времени работы для получения дополнительных сведений.)

Так `Office.addin.showAsTaskpane()` как это асинхронный метод, код будет работать до завершения функции. Подождите завершения с помощью ключевого слова или метода, в зависимости от `await` `then()` того, какой синтаксис JavaScript вы используете.

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>Настройка надстройки для использования общего времени работы

Чтобы использовать эти `showAsTaskpane()` `hide()` методы, надстройка должна использовать общее время работы. Дополнительные сведения см. в Office надстройки для [использования общего времени работы.](configure-your-add-in-to-use-a-shared-runtime.md)

## <a name="preservation-of-state-and-event-listeners"></a>Сохранение слушателей состояния и событий

Только `hide()` методы и методы изменяют `showAsTaskpane()` *видимость* области задач. Они не разгружают или не перезагружают его (или повторно перезагружают его состояние).

Рассмотрим следующий сценарий: области задач разработаны со вкладками. Вкладка **Home** открывается при первом запуске надстройки. Предположим, пользователь открывает **вкладку Параметры,** а затем код в области задач вызывает в ответ `hide()` на какое-либо событие. Еще более поздние вызовы `showAsTaskpane()` кода в ответ на другое событие. Области задач будут появляться снова,  и вкладка Параметры по-прежнему выбрана.

![Снимок экрана области задач с четырьмя вкладками с метками Home, Параметры, Избранное и Учетные записи.](../images/TaskpaneWithTabs.png)

Кроме того, любые слушатели событий, зарегистрированные в области задач, продолжают работать даже при скрытии области задач.

Рассмотрим следующий сценарий: в области задач имеется зарегистрированный обработок для Excel и событий для листа `Worksheet.onActivated` `Worksheet.onDeactivated` с именем **Sheet1**. Активированный обработок вызывает в области задач зеленую точку. Деактивированный обработок превращает точку красной (это ее состояние по умолчанию). Предположим, что код `hide()` вызывается, когда **sheet1** не активируется, а точка красная. Несмотря на то, что области задач скрыты, **лист1** активируется. Позже код `showAsTaskpane()` вызывает в ответ на какое-либо событие. Когда открывается области задач, точка зеленая, так как слушатели и обработчики событий бежали, даже если области задач были скрыты.

## <a name="handle-the-visibility-changed-event"></a>Обработка измененного события видимости

Если код изменяет видимость области задач с помощью или `showAsTaskpane()` `hide()` Office запускает `VisibilityModeChanged` событие. Это событие может быть полезно для обработки. Например, предположим, что в области задач отображается список всех листов в книге. Если при скрытии области задач добавляется новая таблица, то само по себе не добавит новое имя таблицы в список. Но ваш код может ответить на событие, чтобы перезагрузить Worksheet.name всех таблиц в коллекции `VisibilityModeChanged` [Workbook.worksheets,](/javascript/api/excel/excel.workbook#worksheets) как показано в приведенной ниже примере кода. [](/javascript/api/excel/excel.worksheet#name)

Чтобы зарегистрировать обработатель для события, не используйте метод "добавить обработник", как в большинстве Office JavaScript. Вместо этого, существует специальная функция, к которой вы передаете обработок: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onVisibilityModeChanged_listener_). Ниже приведен пример. Обратите внимание, `args.visibilityMode` что свойством [является тип VisibilityMode.](/javascript/api/office/office.visibilitymode)

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

Функция возвращает другую функцию, которая *отстранив* обработитель. Вот простой, но не надежный пример.

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

Метод является асинхронным и возвращает обещание, что означает, что код должен ожидать выполнения обещания, прежде чем он может вызвать обработитель `onVisibilityModeChanged` дерегистрации. 

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

Функция дерегистрации также асинхронна и возвращает обещание. Таким образом, если у вас есть код, который не должен запускаться до завершения дерегулирования, то следует дождаться обещания, возвращенного функцией дерегистрации.

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>Дополнительные материалы

- [Настройка надстройки Office для использования общей среды выполнения JavaScript](configure-your-add-in-to-use-a-shared-runtime.md)
- [Запуск кода в надстройке Office при открытии документа](run-code-on-document-open.md)
