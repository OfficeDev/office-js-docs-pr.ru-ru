---
title: Отображение и скрытие области задач надстройки Office
description: 'Узнайте, как программным образом скрыть или показать пользовательский интерфейс надстройки во время ее непрерывного работы.'
ms.date: 07/08/2021
ms.localizationpriority: medium
---

# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Отображение и скрытие области задач надстройки Office

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Вы можете показать области задач Office надстройки, позвонив в функцию`Office.addin.showAsTaskpane()`.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

Предыдущий код предполагает сценарий, в котором имеется Excel с именем **CurrentQuarterSales**. Надстройка сделает области задач видимыми при активации этого таблицы. Метод обработник `onCurrentQuarter` для Office[. Событие Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-onactivated-member), зарегистрированное для таблицы.

Кроме того, можно скрыть области задач, позвонив в эту `Office.addin.hide()` функцию.

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

Предыдущий код — обработник, зарегистрированный для [Office. Событие Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-ondeactivated-member).

## <a name="additional-details-on-showing-the-task-pane"></a>Дополнительные сведения о показе области задач

При вызове `Office.addin.showAsTaskpane()`Office будет отображаться в области задач файл, который назначен в качестве ID ресурса (`resid`) значения области задач. Это `resid` значение может быть назначено или изменено путем **открытияmanifest.xml** файла `<SourceLocation>` и размещения внутри `<Action xsi:type="ShowTaskpane">` элемента.
(См[. раздел Настройка Office надстройки для использования](configure-your-add-in-to-use-a-shared-runtime.md) общего времени работы для получения дополнительных сведений.)

Так `Office.addin.showAsTaskpane()` как это асинхронный метод, код будет работать до завершения функции. Подождите завершения с помощью `await` `then()` ключевого слова или метода, в зависимости от того, какой синтаксис JavaScript вы используете.

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>Настройка надстройки для использования общего времени работы

Чтобы использовать эти методы `showAsTaskpane()` `hide()` , надстройка должна использовать общее время работы. Дополнительные сведения см. в Office [надстройки для использования общего времени работы](configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="preservation-of-state-and-event-listeners"></a>Сохранение слушателей состояния и событий

Только `hide()` методы `showAsTaskpane()` и методы *изменяют видимость* области задач. Они не разгружают или не перезагружают его (или повторно перезагружают его состояние).

Рассмотрим следующий сценарий: области задач разработаны со вкладками. **Вкладка Home** открывается при первом запуске надстройки. Предположим, пользователь **открывает вкладку Параметры**, `hide()` а затем код в области задач вызывает в ответ на какое-либо событие. Еще более поздние вызовы кода `showAsTaskpane()` в ответ на другое событие. Области задач будут появляться снова, и **вкладка** Параметры по-прежнему выбрана.

![Снимок экрана области задач с четырьмя вкладками с метками Home, Параметры, Избранное и Учетные записи.](../images/TaskpaneWithTabs.png)

Кроме того, любые слушатели событий, зарегистрированные в области задач, продолжают работать даже при скрытии области задач.

Рассмотрим следующий сценарий: в `Worksheet.onActivated` `Worksheet.onDeactivated` области задач имеется зарегистрированный обработок для Excel и событий для листа с именем **Sheet1**. Активированный обработок вызывает в области задач зеленую точку. Деактивированный обработок превращает точку красной (это ее состояние по умолчанию). Предположим, что код вызывается `hide()` , когда **sheet1** не активируется, а точка красная. Несмотря на то, что области задач скрыты, **лист1** активируется. Позже код вызывает `showAsTaskpane()` в ответ на какое-либо событие. Когда открывается области задач, точка зеленая, так как слушатели и обработчики событий бежали, даже если области задач были скрыты.

## <a name="handle-the-visibility-changed-event"></a>Обработка измененного события видимости

Когда код изменяет видимость `showAsTaskpane()` `hide()`области задач с помощью или Office запускает `VisibilityModeChanged` событие. Это событие может быть полезно для обработки. Например, предположим, что в области задач отображается список всех листов в книге. Если при скрытии области задач добавляется новая таблица, то само по себе не добавит новое имя таблицы в список. Но ваш код `VisibilityModeChanged` может откликнуться на событие, чтобы перезагрузить Worksheet.name всех таблиц в коллекции [Workbook.worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member), как показано в приведенной ниже примере кода.[](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member)

Чтобы зарегистрировать обработатель для события, не используйте метод "добавить обработник", как в большинстве Office JavaScript. Вместо этого существует специальная функция, к которой вы передаете обработчиве: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#office-office-addin-onvisibilitymodechanged-member(1)). Ниже приведен пример. Обратите внимание, что `args.visibilityMode` свойством [является тип VisibilityMode](/javascript/api/office/office.visibilitymode).

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

Функция возвращает другую функцию, которая *отстранив обработитель* . Вот простой, но не надежный пример.

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

Метод `onVisibilityModeChanged` является асинхронным и возвращает обещание, что означает, что код должен ожидать выполнения обещания, прежде чем он может вызвать обработитель дерегистрации.

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

## <a name="see-also"></a>См. также

- [Настройка надстройки Office для использования общей среды выполнения JavaScript](configure-your-add-in-to-use-a-shared-runtime.md)
- [Запуск кода в надстройке Office при открытии документа](run-code-on-document-open.md)
