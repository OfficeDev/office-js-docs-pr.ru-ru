---
title: Отображение и скрытие надстройки Office в общей среде выполнения
description: Сведения о том, как программно скрыть или отобразить пользовательский интерфейс надстройки, когда он работает постоянно
ms.date: 05/17/2020
localization_priority: Normal
ms.openlocfilehash: e09fa7d0a39c7157823911307558889e2ade89db
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430571"
---
# <a name="show-or-hide-an-office-add-in-in-a-shared-runtime"></a>Отображение и скрытие надстройки Office в общей среде выполнения

Надстройка Office может включать любые из следующих частей:

- Область задач
- Файл функции без пользовательского интерфейса (пользовательские функции, которые не используют область задач или другие элементы пользовательского интерфейса)
- Пользовательская функция Excel

По умолчанию каждая часть выполняется в отдельной среде выполнения JavaScript с собственным глобальным объектом и глобальными переменными.

Надстройки могут совместно использовать общую среду выполнения JavaScript с двумя или более частями. Эта общая функция среды выполнения включает новые API, которые скрывают и повторно открывают область задач во время выполнения надстройки.

## <a name="configure-an-add-in-to-use-a-shared-runtime"></a>Настройка надстройки для использования общей среды выполнения

Чтобы настроить надстройку для использования общей среды выполнения, ознакомьтесь со статьей [Настройка надстройки Office для использования общей среды выполнения](configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="show-and-hide-the-task-pane"></a>Отображение и скрытие области задач

Новые API находятся в `Office.addin` свойстве. Чтобы отобразить область задач, вызывается код `Office.addin.showAsTaskpane()` . В области задач Office будет отображаться страница, которая была назначена ИДЕНТИФИКАТОРу ресурса ( `resid` ) для области задач. Это то `resid` , которое было назначено в `<SourceLocation>` `<Action xsi:type="ShowTaskpane">` манифесте. (См. [Настройка надстройки Office для использования совместно используемой среды выполнения](configure-your-add-in-to-use-a-shared-runtime.md)).

Это асинхронный метод, поэтому код должен ожидать его, если следующий код не будет выполняться, пока он не будет завершен. Дождитесь этого завершения с помощью `await` ключевого слова или метода в зависимости от используемого `then()` синтаксиса JavaScript. Ниже предполагается, что имеется лист Excel с именем **курренткуартерсалес**. Надстройка должна сделать область задач видимой при активации этого листа. Метод `onCurrentQuarter` является обработчиком для события [Office. лист. OnActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) , зарегистрированного для листа.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

Чтобы скрыть область задач, вызывается код `Office.addin.hide()` . В следующем примере показан обработчик, зарегистрированный для события [Office. лист. OnDeactivate](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated) .

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

### <a name="preservation-of-state-and-event-listeners"></a>Сохранение прослушивателей состояний и событий

`hide()`Методы and `showAsTaskpane()` изменяют только *видимость* области задач. Они не выгружают и не загружают их (или повторно инициализируют состояние).

Рассмотрим следующий сценарий: область задач разработана с использованием вкладок. Вкладка **Главная** открывается при первом запуске надстройки. Предположим, что пользователь открывает вкладку **Параметры** , а затем код в области задач вызывается `hide()` в ответ на событие. Все еще позже вызывается код `showAsTaskpane()` в ответ на другое событие. Область задач будет снова отображаться, а вкладка **Параметры** все еще будет выбрана.

![Снимок экрана с областью задач с четырьмя вкладками "Главная", "Параметры", "Избранное" и "учетные записи".](../images/TaskpaneWithTabs.png)

Кроме того, все прослушиватели событий, зарегистрированные в области задач, продолжают выполняться, даже если область задач скрыта.

Рассмотрим следующий сценарий: область задач содержит зарегистрированный обработчик для Excel `Worksheet.onActivated` и `Worksheet.onDeactivated` событий для листа с именем **Лист1**. Активированный обработчик вызывает отображение зеленой точки в области задач. Отключенный обработчик включает красную точку (ее состояние по умолчанию). Предположим, что код вызывается, `hide()` когда **Лист1** не активирован, а точка имеет красный цвет. Когда область задач скрыта, **Лист1** активируется. Последующие вызовы кода `showAsTaskpane()` в ответ на событие. Когда откроется область задач, точка будет зеленым, так как прослушиватели и обработчики событий запускаются несмотря на то, что область задач скрыта.

### <a name="handle-visibility-changed-event"></a>Обработка события изменения видимости

Когда код изменяет видимость области задач с `showAsTaskpane()` или `hide()` , Office запускает `VisibilityModeChanged` событие. Это может пригодиться для обработки этого события. Например, предположим, что в области задач отображается список всех листов в книге. Если новый лист добавляется в то время, когда область задач скрыта, то при отображении видимой области задач в списке добавляется имя нового листа. Но ваш код может ответить на `VisibilityModeChanged` событие, чтобы перегрузить свойство [Worksheet.Name](/javascript/api/excel/excel.worksheet#name) для всех листов в коллекции [Workbook. листы](/javascript/api/excel/excel.workbook#worksheets) , как показано в приведенном ниже примере кода.

Чтобы зарегистрировать обработчик для события, не используйте метод "добавить обработчик", как в большинстве контекстов Office JavaScript. Вместо этого используется специальная функция, для которой вы передаете свой обработчик: [Office. AddIn. онвисибилитимодечанжед](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-). Ниже приведен пример. Обратите внимание, что `args.visibilityMode` свойство имеет тип [висибилитимоде](/javascript/api/office/office.visibilitymode).

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

Функция возвращает другую функцию, которая *отменяет регистрацию* обработчика. Вот простой, но не надежный, пример:

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

`onVisibilityModeChanged`Метод является асинхронным, что означает, что если ваш код вызывает обработчик *отмены регистрации* , который `onVisibilityModeChanged` возвращает значение, необходимо убедиться, что `onVisibilityModeChanged` оно завершено до вызова обработчика отмены регистрации. Один из способов сделать это — использовать `await` ключевое слово для вызова метода, как показано в следующем примере.

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

Если вы хотите использовать только пред ES2015 JavaScript, ваш код может использовать `then` метод, чтобы дождаться, пока возвращенный объект Promise не будет разрешен и присвоить возвращаемую функцию глобальной переменной, как показано в следующем примере.

```javascript
var removeVisibilityModeHandler;

Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
}).then(function(removeHandler) {
        removeVisibilityModeHandler = removeHandler;
    });

// In some later code path, deregister with:
removeVisibilityModeHandler();
```

Функция отмены регистрации сама по себе асинхронна. Таким образом, если код не должен выполняться до завершения отмены регистрации, функция дерегистрации должна также ожидаться с помощью `await` ключевого слова или `then` метода, как показано в следующих примерах.

Чтобы отменить регистрацию обработчика:

```javascript
await removeVisibilityModeHandler();
// subsequent code here

// or use pre-ES2015 syntax:
removeVisibilityModeHandler().then(function () {
        // subsequent code here
    })
```
