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
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a><span data-ttu-id="afd13-103">Запуск кода в надстройки Office при запуске документа</span><span class="sxs-lookup"><span data-stu-id="afd13-103">Run code in your Office Add-in when the document opens</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="afd13-104">Вы можете настроить надстройку Office для загрузки и запуска кода сразу после открытия документа.</span><span class="sxs-lookup"><span data-stu-id="afd13-104">You can configure your Office Add-in to load and run code as soon as the document is opened.</span></span> <span data-ttu-id="afd13-105">Это полезно, если вам нужно зарегистрировать обработчики событий, предварительно загрузить данные для области задач, синхронизировать пользовательский интерфейс или выполнить другие задачи, прежде чем надстройка будет видна.</span><span class="sxs-lookup"><span data-stu-id="afd13-105">This is useful if you need to register event handlers, pre-load data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.</span></span>

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a><span data-ttu-id="afd13-106">Настройка надстройки для загрузки при ее открытие</span><span class="sxs-lookup"><span data-stu-id="afd13-106">Configure your add-in to load when the document opens</span></span>

<span data-ttu-id="afd13-107">Следующий код настраивает надстройку для загрузки и запуска при запуске документа.</span><span class="sxs-lookup"><span data-stu-id="afd13-107">The following code configures your add-in to load and start running when the document is opened.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> <span data-ttu-id="afd13-108">Метод `setStartupBehavior` является асинхронным.</span><span class="sxs-lookup"><span data-stu-id="afd13-108">The `setStartupBehavior` method is asynchronous.</span></span>

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a><span data-ttu-id="afd13-109">Настройка надстройки для ненагрузки при открытом документе</span><span class="sxs-lookup"><span data-stu-id="afd13-109">Configure your add-in for no load behavior on document open</span></span>

<span data-ttu-id="afd13-110">Следующий код настраивает надстройки так, чтобы она не запускалась при запуске документа.</span><span class="sxs-lookup"><span data-stu-id="afd13-110">The following code configures your add-in not to start when the document is opened.</span></span> <span data-ttu-id="afd13-111">Вместо этого он будет запускаться, когда пользователь каким-либо образом вовлекет его, например нажатие кнопки ленты или открытие области задач.</span><span class="sxs-lookup"><span data-stu-id="afd13-111">Instead, it will start when the user engages it in some way, such as choosing a ribbon button or opening the task pane.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a><span data-ttu-id="afd13-112">Получить текущее поведение при загрузке</span><span class="sxs-lookup"><span data-stu-id="afd13-112">Get the current load behavior</span></span>

<span data-ttu-id="afd13-113">Чтобы определить текущее поведение при запуске, запустите следующую функцию, которая возвращает `Office.StartupBehavior` объект.</span><span class="sxs-lookup"><span data-stu-id="afd13-113">To determine what the current startup behavior is, run the following function, which returns an `Office.StartupBehavior` object.</span></span>

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a><span data-ttu-id="afd13-114">Запуск кода при запуске документа</span><span class="sxs-lookup"><span data-stu-id="afd13-114">How to run code when the document opens</span></span>

<span data-ttu-id="afd13-115">Если надстройка настроена на загрузку при открытом документе, она будет запускаться немедленно.</span><span class="sxs-lookup"><span data-stu-id="afd13-115">When your add-in is configured to load on document open, it will run immediately.</span></span> <span data-ttu-id="afd13-116">Будет `Office.initialize` вызван обработник событий.</span><span class="sxs-lookup"><span data-stu-id="afd13-116">The `Office.initialize` event handler will be called.</span></span> <span data-ttu-id="afd13-117">Поместите код запуска в обработок событий или `Office.initialize` `Office.onReady` обработок событий.</span><span class="sxs-lookup"><span data-stu-id="afd13-117">Place your startup code in the `Office.initialize` or `Office.onReady` event handler.</span></span>

<span data-ttu-id="afd13-118">В следующем коде надстройки Excel показано, как зарегистрировать обработитель событий для событий изменения с активного листа.</span><span class="sxs-lookup"><span data-stu-id="afd13-118">The following Excel add-in code shows how to register an event handler for change events from the active worksheet.</span></span> <span data-ttu-id="afd13-119">Если вы настроите надстройку для загрузки при открытом документе, этот код зарегистрирует обработчик событий при его открытом документе.</span><span class="sxs-lookup"><span data-stu-id="afd13-119">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="afd13-120">События изменений можно обрабатывать до открытия области задач.</span><span class="sxs-lookup"><span data-stu-id="afd13-120">You can handle change events before the task pane is opened.</span></span>

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

<span data-ttu-id="afd13-121">В следующем коде надстройки PowerPoint показано, как зарегистрировать обработок событий для событий изменения выделения из документа PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="afd13-121">The following PowerPoint add-in code shows how to register an event handler for selection change events from the PowerPoint document.</span></span> <span data-ttu-id="afd13-122">Если вы настроите надстройку для загрузки при открытом документе, этот код зарегистрирует обработчик событий при его открытом документе.</span><span class="sxs-lookup"><span data-stu-id="afd13-122">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="afd13-123">События изменений можно обрабатывать до открытия области задач.</span><span class="sxs-lookup"><span data-stu-id="afd13-123">You can handle change events before the task pane is opened.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="afd13-124">См. также</span><span class="sxs-lookup"><span data-stu-id="afd13-124">See also</span></span>

- [<span data-ttu-id="afd13-125">Настройка надстройки Office для использования общей времени работы JavaScript</span><span class="sxs-lookup"><span data-stu-id="afd13-125">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="afd13-126">Совместное работу с данными и событиями между пользовательскими функциями Excel и учебником по области задач</span><span class="sxs-lookup"><span data-stu-id="afd13-126">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="afd13-127">Работа с событиями при помощи API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="afd13-127">Work with Events using the Excel JavaScript API</span></span>](../excel/excel-add-ins-events.md)
