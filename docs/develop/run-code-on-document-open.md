---
title: Запуск кода в надстройки Office при запуске документа
description: Узнайте, как запускать код в надстройки Office при запуске документа.
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 1655c053a4fa6f92aae95f2155991fa4f7f7a5a7
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789246"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>Запуск кода в надстройки Office при запуске документа

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Вы можете настроить надстройку Office для загрузки и запуска кода сразу после открытия документа. Это полезно, если вам нужно зарегистрировать обработчики событий, предварительно загрузить данные для области задач, синхронизировать пользовательский интерфейс или выполнить другие задачи, прежде чем надстройка будет видна.

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>Настройка надстройки для загрузки при ее открытие

Следующий код настраивает надстройку для загрузки и запуска при запуске документа.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> Метод `setStartupBehavior` является асинхронным.

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>Настройка надстройки для ненагрузки при открытом документе

Следующий код настраивает надстройки так, чтобы она не запускалась при запуске документа. Вместо этого он будет запускаться, когда пользователь каким-либо образом вовлекет его, например нажатие кнопки ленты или открытие области задач.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>Получить текущее поведение при загрузке

Чтобы определить текущее поведение при запуске, запустите следующую функцию, которая возвращает `Office.StartupBehavior` объект.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a>Запуск кода при запуске документа

Если надстройка настроена на загрузку при открытом документе, она будет запускаться немедленно. Будет `Office.initialize` вызван обработник событий. Поместите код запуска в обработок событий или `Office.initialize` `Office.onReady` обработок событий.

В следующем коде надстройки Excel показано, как зарегистрировать обработитель событий для событий изменения с активного листа. Если вы настроите надстройку для загрузки при открытом документе, этот код зарегистрирует обработчик событий при его открытом документе. События изменений можно обрабатывать до открытия области задач.

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
  return Excel.run(function(context) {
    return context.sync().then(function() {
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);
    });
  });
}
```

В следующем коде надстройки PowerPoint показано, как зарегистрировать обработок событий для событий изменения выделения из документа PowerPoint. Если вы настроите надстройку для загрузки при открытом документе, этот код зарегистрирует обработчик событий при его открытом документе. События изменений можно обрабатывать до открытия области задач.

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

## <a name="see-also"></a>См. также

- [Настройка надстройки Office для использования общей времени работы JavaScript](configure-your-add-in-to-use-a-shared-runtime.md)
- [Совместное работу с данными и событиями между пользовательскими функциями Excel и учебником по области задач](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Работа с событиями при помощи API JavaScript для Excel](../excel/excel-add-ins-events.md)
