---
title: Отображение и скрытие области задач надстройки Office
description: Узнайте, как программно скрыть или отобразить пользовательский интерфейс надстройки во время непрерывной работы.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 95f8c716bf1a0331fe47bc74e5aad49c17b65437
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660132"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Отображение и скрытие области задач надстройки Office

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Вы можете отобразить область задач надстройки Office, вызвав функцию `Office.addin.showAsTaskpane()` .

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

В предыдущем коде предполагается, что существует лист Excel с именем **CurrentQuarterSales**. Надстройка будет делать область задач видимой при активации этого листа. Этот метод `onCurrentQuarter` является обработчиком события [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-onactivated-member) , зарегистрированного для листа.

Вы также можете скрыть область задач, вызвав функцию `Office.addin.hide()` .

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

Предыдущий код — это обработчик, зарегистрированный для [события Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-ondeactivated-member) .

## <a name="additional-details-on-showing-the-task-pane"></a>Дополнительные сведения о отключении области задач

При вызове `Office.addin.showAsTaskpane()`Office отобразит в области задач файл, назначенный в качестве значения идентификатора ресурса (`resid`) области задач. Это `resid` значение можно назначить или изменить, открыв файлmanifest.xmlи **найдите** внутри **\<SourceLocation\>** `<Action xsi:type="ShowTaskpane">` элемента.
( [Дополнительные сведения см. в разделе "Настройка надстройки Office для использования](configure-your-add-in-to-use-a-shared-runtime.md) общей среды выполнения".)

Так `Office.addin.showAsTaskpane()` как это асинхронный метод, код будет выполняться до завершения работы функции. Дождитесь завершения с помощью ключевого `await` `then()` слова или метода в зависимости от используемого синтаксиса JavaScript.

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>Настройка надстройки для использования общей среды выполнения

Для использования этих `showAsTaskpane()` методов `hide()` надстройка должна использовать общую среду выполнения. Дополнительные сведения см. в разделе ["Настройка надстройки Office для использования общей среды выполнения"](configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="preservation-of-state-and-event-listeners"></a>Сохранение состояния и прослушивателей событий

Методы `hide()` и `showAsTaskpane()` методы изменяют *только видимость* области задач. Они не выгружают и не перезагружают его (или повторно инициализют его состояние).

Рассмотрим следующий сценарий: область задач разработана с помощью вкладок. **Вкладка "** Главная" открывается при первом запуске надстройки. Предположим, что пользователь открывает вкладку **"** Параметры", `hide()` а затем код в области задач вызывается в ответ на определенное событие. Но позже код вызывается `showAsTaskpane()` в ответ на другое событие. Область задач снова появится, а вкладка "Параметры" по-прежнему будет выбрана.

![Снимок экрана: область задач с четырьмя вкладками "Главная", "Параметры", "Избранное" и "Учетные записи".](../images/TaskpaneWithTabs.png)

Кроме того, все прослушиватели событий, зарегистрированные в области задач, продолжают работать, даже если область задач скрыта.

Рассмотрим следующий сценарий: в области задач есть зарегистрированный обработчик для Excel `Worksheet.onActivated` `Worksheet.onDeactivated` и события для листа **с именем Sheet1**. Активированный обработчик вызывает отображение зеленой точки в области задач. Отключенный обработчик преобразует точку в красный цвет (это состояние по умолчанию). Предположим, что код вызывается `hide()` , если **sheet1** не активирован, а точка красная. Пока область задач скрыта, **лист1** активируется. Последующие вызовы кода `showAsTaskpane()` в ответ на некоторое событие. Когда откроется область задач, точка будет зеленой, так как прослушиватели событий и обработчики выполнялись, даже если область задач была скрыта.

## <a name="handle-the-visibility-changed-event"></a>Обработка события изменения видимости

Когда код изменяет видимость области `showAsTaskpane()` задач с помощью `hide()`или, Office активирует `VisibilityModeChanged` событие. Это событие может быть полезно для обработки. Например, предположим, что в области задач отображается список всех листов в книге. Если новый лист добавляется, когда область задач скрыта, то, чтобы сделать область задач видимой, это само по себе не приведет к добавлению нового имени листа в список. Но ваш код `VisibilityModeChanged` может реагировать на событие, чтобы перезагрузить свойство [Worksheet.name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member) всех листов в коллекции [Workbook.worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member) , как показано в примере кода ниже.

Чтобы зарегистрировать обработчик события, не используйте метод add handler, как в большинстве контекстов JavaScript для Office. Вместо этого существует специальная функция, которой передается обработчик: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#office-office-addin-onvisibilitymodechanged-member(1)). Ниже приведен пример. Обратите внимание, что `args.visibilityMode` свойство имеет тип [VisibilityMode](/javascript/api/office/office.visibilitymode).

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

Функция возвращает другую функцию, которая *отменяет* регистрацию обработчика. Ниже приведен простой, но не надежный пример.

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

Метод `onVisibilityModeChanged` является асинхронным и возвращает обещание. Это означает, что код должен ожидать выполнения обещания, прежде чем он сможет вызвать обработчик **отмены** регистрации.

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

Функция отмены регистрации также является асинхронной и возвращает обещание. Таким образом, если у вас есть код, который не должен выполняться до завершения отмены регистрации, следует ожидать обещание, возвращенное функцией отмены регистрации.

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>См. также

- [Настройка надстройки Office для использования общей среды выполнения JavaScript](configure-your-add-in-to-use-a-shared-runtime.md)
- [Запуск кода в надстройке Office при открытии документа](run-code-on-document-open.md)
