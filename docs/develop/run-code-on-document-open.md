---
title: Запуск кода в надстройке Office при открытии документа
description: Узнайте, как выполнять код в надстройке Office, когда открывается документ.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1a1c3277a349dc4054da5f089c62331296590021
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958441"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>Запуск кода в надстройке Office при открытии документа

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Вы можете настроить надстройку Office для загрузки и выполнения кода сразу после открытия документа. Это полезно, если необходимо зарегистрировать обработчики событий, предварительно загрузить данные для области задач, синхронизировать пользовательский интерфейс или выполнить другие задачи, прежде чем надстройка будет видна.

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>Настройка надстройки для загрузки при открываемом документе

Следующий код настраивает загрузку и запуск надстройки при открытии документа.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> Метод `setStartupBehavior` является асинхронным.

## <a name="place-startup-code-in-officeinitialize"></a>Размещение кода запуска в Office.initialize

Если надстройка настроена для загрузки открытого документа, она будет выполняться немедленно. Будет `Office.initialize` вызван обработчик событий. Поместите код запуска в обработчик `Office.initialize` событий или обработчик `Office.onReady` событий.

В следующем коде надстройки Excel показано, как зарегистрировать обработчик событий для событий изменений с активного листа. Если вы настроите надстройку для загрузки открытого документа, этот код зарегистрирует обработчик событий при открытии документа. События изменений можно обрабатывать до открытия области задач.

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

В следующем коде надстройки PowerPoint показано, как зарегистрировать обработчик событий для выбора событий изменения из документа PowerPoint. Если вы настроите надстройку для загрузки открытого документа, этот код зарегистрирует обработчик событий при открытии документа. События изменений можно обрабатывать до открытия области задач.

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

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>Настройка надстройки для ненастройки при открытии документа

Следующий код настраивает надстройку не запускать при открытии документа. Вместо этого он запускается, когда пользователь вовсю его использует, например нажав кнопку ленты или открыв область задач.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>Получение текущего поведения загрузки

Чтобы определить текущее поведение при запуске, выполните следующий метод, который возвращает `Office.StartupBehavior` объект.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="see-also"></a>См. также

- [Настройка надстройки Office для использования общей среды выполнения JavaScript](configure-your-add-in-to-use-a-shared-runtime.md)
- [Руководство по совместному доступу к данным и событиям между пользовательскими функциями Excel и областью задач](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Работа с событиями при помощи API JavaScript для Excel](../excel/excel-add-ins-events.md)
