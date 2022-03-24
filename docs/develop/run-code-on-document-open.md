---
title: Запуск кода в надстройке Office при открытии документа
description: Узнайте, как запускать код Office надстройки при запуске документа.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 20cd7a90f34c0141ca166119ceae92960a904595
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744082"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>Запуск кода в надстройке Office при открытии документа

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Вы можете настроить Office надстройку для загрузки и запуска кода сразу после открытия документа. Это полезно, если необходимо зарегистрировать обработчики событий, предварительно загрузить данные для области задач, синхронизировать пользовательский интерфейс или выполнить другие задачи до того, как надстройка будет видна.

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>Настройка надстройки для загрузки при открываемом документе

Следующий код настраивает надстройку для загрузки и запуска при запуске документа.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> Метод `setStartupBehavior` асинхронный.

## <a name="place-startup-code-in-officeinitialize"></a>Поместите код запуска в Office.initialize

Когда надстройка настроена для загрузки открытого документа, она будет немедленно работать. Обработник `Office.initialize` событий будет вызван. Поместите код запуска в обработник `Office.initialize` `Office.onReady` событий или обработник событий.

В следующем Excel кода надстройки показано, как зарегистрировать обработник событий для событий изменения из активного таблицы. Если вы настроите надстройку для загрузки открытого документа, этот код зарегистрирует обработчик событий при открываемом документе. Перед открытием области задач можно обрабатывать события изменений.

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.initialize = () => {
  // Add the event handler.
  Excel.run(async context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(onChange);

    await context.sync();
    console.log("A handler has been registered for the onChanged event.");
  });
};

/**
 * Handle the changed event from the worksheet.
 *
 * @param event The event information from Excel
 */
async function onChange(event) {
    await Excel.run(async (context) => {    
        await context.sync();
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);
  });
}
```

В следующем PowerPoint надстройки показано, как зарегистрировать обработник событий для событий изменения выбора из PowerPoint документа. Если вы настроите надстройку для загрузки открытого документа, этот код зарегистрирует обработчик событий при открываемом документе. Перед открытием области задач можно обрабатывать события изменений.

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange);
    console.log("A handler has been registered for the onChanged event.");
  }
});

/**
 * Handle the changed event from the PowerPoint document.
 *
 * @param event The event information from PowerPoint
 */
async function onChange(event) {
  console.log("Change type of event: " + event.type);
}
```

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>Настройка надстройки для ненагрузки при открытом документе

Следующий код настраивает надстройки, чтобы не запускаться при открываемом документе. Вместо этого он начнется, когда пользователь вовлекет его каким-либо образом, например, выбирая кнопку ленты или открывая области задач.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>Получить текущее поведение нагрузки

Чтобы определить текущее поведение запуска, запустите следующую функцию, которая возвращает `Office.StartupBehavior` объект.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="see-also"></a>См. также

- [Настройка надстройки Office для использования общей среды выполнения JavaScript](configure-your-add-in-to-use-a-shared-runtime.md)
- [Обмениваться данными и событиями между Excel пользовательскими функциями и учебником по области задач](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Работа с событиями при помощи API JavaScript для Excel](../excel/excel-add-ins-events.md)
