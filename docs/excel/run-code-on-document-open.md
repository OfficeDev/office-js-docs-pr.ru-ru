---
title: Запуск кода в надстройке Excel при открытии документа
description: Запуск кода в надстройке Excel при открытии документа.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: 0a9090315a4ddca80e25a94092c779a3f3271087
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217951"
---
# <a name="run-code-in-your-excel-add-in-when-the-document-opens"></a>Запуск кода в надстройке Excel при открытии документа

Вы можете настроить надстройку Excel для загрузки и запуска кода сразу после открытия документа. Это полезно, если необходимо зарегистрировать обработчики событий, предварительно загрузить данные для области задач, выполнить синхронизацию пользовательского интерфейса или выполнить другие задачи, чтобы надстройка стала видимой.

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>Настройка загрузки надстройки при открытии документа

Приведенный ниже код настраивает надстройку для загрузки и запуска при открытии документа.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> `setStartupBehavior`Метод является асинхронным.

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>Настройка надстройки на отсутствие режима загрузки при открытии документа

Приведенный ниже код настраивает надстройку, не запускаясь при открытии документа. Вместо этого он запускается, когда пользователь применяет его каким-либо способом (например, для выбора кнопки на ленте или открытия области задач).

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>Получение текущего поведения при загрузке

Чтобы определить, каково текущее поведение при запуске, выполните следующую функцию, которая возвращает объект Office. Стартупбехавиор.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a>Выполнение кода при открытии документа

Когда ваша надстройка настроена на загрузку документа, он будет запущен немедленно. `Office.initialize`Будет вызван обработчик событий. Поместите код запуска в `Office.initialize` обработчик событий.

В приведенном ниже коде показано, как зарегистрировать обработчик событий для событий Changes с активного листа. Если вы настраиваете надстройку для загрузки при открытии документа, этот код регистрирует обработчик событий при открытии документа. События изменения можно обработать до открытия области задач.


```JavaScript
//This is called as soon as the document opens.
//Put your startup code here.
Office.initialize = () => {
  // Add the event handler
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

## <a name="see-also"></a>См. также

- [Обмен данными и событиями между пользовательскими функциями и областью задач Excel](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)