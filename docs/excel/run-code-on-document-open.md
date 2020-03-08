---
title: Запуск кода в надстройке Excel при открытии документа (Предварительная версия)
description: Запуск кода в надстройке Excel при открытии документа.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 5b8c646a1154540244b1f5e0ac47ad8eaec1801f
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/26/2020
ms.locfileid: "42284188"
---
# <a name="run-code-in-your-excel-add-in-when-the-document-opens-preview"></a><span data-ttu-id="2dcfc-103">Запуск кода в надстройке Excel при открытии документа (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="2dcfc-103">Run code in your Excel add-in when the document opens (preview)</span></span>

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="2dcfc-104">Вы можете настроить надстройку Excel для загрузки и запуска кода сразу после открытия документа.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-104">You can configure your Excel add-in to load and run code as soon as the document is opened.</span></span> <span data-ttu-id="2dcfc-105">Это полезно, если необходимо зарегистрировать обработчики событий, предварительную загрузку данных для области задач, синхронизацию пользовательского интерфейса или выполнение других задач, чтобы надстройка стала видимой.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-105">This is useful if you need to register event handlers, preload data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.</span></span>

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a><span data-ttu-id="2dcfc-106">Настройка загрузки надстройки при открытии документа</span><span class="sxs-lookup"><span data-stu-id="2dcfc-106">Configure your add-in to load when the document opens</span></span>

<span data-ttu-id="2dcfc-107">Приведенный ниже код настраивает надстройку для загрузки и запуска при открытии документа.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-107">The following code configures your add-in to load and start running when the document is opened.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> <span data-ttu-id="2dcfc-108">`setStartupBehavior` Метод является асинхронным.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-108">The `setStartupBehavior` method is asynchronous.</span></span>

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a><span data-ttu-id="2dcfc-109">Настройка надстройки на отсутствие режима загрузки при открытии документа</span><span class="sxs-lookup"><span data-stu-id="2dcfc-109">Configure your add-in for no load behavior on document open</span></span>

<span data-ttu-id="2dcfc-110">Приведенный ниже код настраивает надстройку, не запускаясь при открытии документа.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-110">The following code configures your add-in not to start when the document is opened.</span></span> <span data-ttu-id="2dcfc-111">Вместо этого он запускается, когда пользователь применяет его каким-либо способом (например, для выбора кнопки на ленте или открытия области задач).</span><span class="sxs-lookup"><span data-stu-id="2dcfc-111">Instead it will start when the user engages it in some way (such as choosing a ribbon button, or opening the task pane.)</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a><span data-ttu-id="2dcfc-112">Получение текущего поведения при загрузке</span><span class="sxs-lookup"><span data-stu-id="2dcfc-112">Get the current load behavior</span></span>

<span data-ttu-id="2dcfc-113">Чтобы определить, каково текущее поведение при запуске, выполните следующую функцию, которая возвращает объект Office. Стартупбехавиор.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-113">To determine what the current startup behavior is, run the following function, which returns an Office.StartupBehavior object.</span></span>

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a><span data-ttu-id="2dcfc-114">Выполнение кода при открытии документа</span><span class="sxs-lookup"><span data-stu-id="2dcfc-114">How to run code when the document opens</span></span>

<span data-ttu-id="2dcfc-115">Когда ваша надстройка настроена на загрузку документа, он будет запущен немедленно.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-115">When your add-in is configured to load on document open, it will run immediately.</span></span> <span data-ttu-id="2dcfc-116">Будет `Office.initialize` вызван обработчик событий.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-116">The `Office.initialize` event handler will be called.</span></span> <span data-ttu-id="2dcfc-117">Поместите код запуска в обработчик `Office.initialize` событий.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-117">Place your startup code in the `Office.initialize` event handler.</span></span>

<span data-ttu-id="2dcfc-118">В приведенном ниже коде показано, как зарегистрировать обработчик событий для событий Changes с активного листа.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-118">The following code shows how to register an event handler for change events from the active worksheet.</span></span> <span data-ttu-id="2dcfc-119">Если вы настраиваете надстройку для загрузки при открытии документа, этот код регистрирует обработчик событий при открытии документа.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-119">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="2dcfc-120">События изменения можно обработать до открытия области задач.</span><span class="sxs-lookup"><span data-stu-id="2dcfc-120">You can handle change events before the task pane is opened.</span></span>


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

## <a name="see-also"></a><span data-ttu-id="2dcfc-121">См. также</span><span class="sxs-lookup"><span data-stu-id="2dcfc-121">See also</span></span>

- [<span data-ttu-id="2dcfc-122">Обмен данными и событиями между пользовательскими функциями и областью задач Excel</span><span class="sxs-lookup"><span data-stu-id="2dcfc-122">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)