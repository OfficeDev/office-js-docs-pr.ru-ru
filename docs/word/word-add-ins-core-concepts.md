---
title: Основные концепции программирования с помощью API JavaScript для Word
description: Создание надстроек для Word с помощью API JavaScript для Word.
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: 1e7a90d4be378ed9b2c1f30ebebd4a0beec45a11
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293095"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a><span data-ttu-id="61964-103">Основные концепции программирования с помощью API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="61964-103">Fundamental programming concepts with the Word JavaScript API</span></span>

<span data-ttu-id="61964-104">В этой статье описаны основные концепции использования [API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md) с целью создания надстроек для Word 2016 и более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="61964-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins for Word 2016 or later.</span></span>

## <a name="referencing-officejs"></a><span data-ttu-id="61964-105">Ссылки на Office.js</span><span class="sxs-lookup"><span data-stu-id="61964-105">Referencing Office.js</span></span>

<span data-ttu-id="61964-106">Файл Office.js можно получить из указанных ниже расположений.</span><span class="sxs-lookup"><span data-stu-id="61964-106">You can reference Office.js from the following locations:</span></span>

- <span data-ttu-id="61964-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`. Используйте этот ресурс для рабочих надстроек.</span><span class="sxs-lookup"><span data-stu-id="61964-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - use this resource for production add-ins.</span></span>

- <span data-ttu-id="61964-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`. Используйте этот ресурс для применения предварительных функций.</span><span class="sxs-lookup"><span data-stu-id="61964-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use this resource to try out preview features.</span></span>

## <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="61964-109">Наборы обязательных элементов API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="61964-109">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="61964-110">Наборы требований — это именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="61964-110">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="61964-111">Надстройки Office используют наборы обязательных элементов, указанных в манифесте, или проверки в среде выполнения, чтобы определить, поддерживает ли клиентское приложение Office необходимые API.</span><span class="sxs-lookup"><span data-stu-id="61964-111">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="61964-112">Подробнее о наборах обязательных элементов API JavaScript для Word см. в статье [Наборы обязательных элементов API JavaScript для Word](../reference/requirement-sets/word-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="61964-112">For detailed information about Word JavaScript API requirement sets, see [Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md).</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="61964-113">Запуск надстроек Word</span><span class="sxs-lookup"><span data-stu-id="61964-113">Running Word add-ins</span></span>

<span data-ttu-id="61964-114">Чтобы запустить надстройку, воспользуйтесь обработчиком событий `Office.initialize`Office.initialize.</span><span class="sxs-lookup"><span data-stu-id="61964-114">To run your add-in, use an `Office.initialize` event handler.</span></span> <span data-ttu-id="61964-115">Дополнительные сведения об инициализации надстроек см. в статье [Общие сведения об API](../develop/understanding-the-javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="61964-115">For more information about add-in initialization, see [Understanding the API](../develop/understanding-the-javascript-api-for-office.md).</span></span>

<span data-ttu-id="61964-116">Надстройки, предназначенные для Word 2016 и более поздних версий, могут использовать API для Word.</span><span class="sxs-lookup"><span data-stu-id="61964-116">Add-ins that target Word 2016 or later can use the Word-specific APIs.</span></span> <span data-ttu-id="61964-117">Они передают методу `Word.run()` логику взаимодействия с Word в качестве функции.</span><span class="sxs-lookup"><span data-stu-id="61964-117">They pass the Word-interaction logic as a function into the `Word.run()` method.</span></span> <span data-ttu-id="61964-118">Дополнительные сведения о том, как работать с документом Word в этой модели программирования, см. в статье [Использование модели API для определенных приложений](../develop/application-specific-api-model.md).</span><span class="sxs-lookup"><span data-stu-id="61964-118">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about how to interact with the Word document in this programming model.</span></span>

<span data-ttu-id="61964-119">В следующем примере показано, как инициализировать и запустить надстройку Word с помощью метода `Word.run()`.</span><span class="sxs-lookup"><span data-stu-id="61964-119">The following example shows how to initialize and run a Word add-in by using the `Word.run()` method.</span></span>

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

## <a name="see-also"></a><span data-ttu-id="61964-120">См. также</span><span class="sxs-lookup"><span data-stu-id="61964-120">See also</span></span>

- [<span data-ttu-id="61964-121">Обзор API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="61964-121">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="61964-122">Создание первой надстройки Word</span><span class="sxs-lookup"><span data-stu-id="61964-122">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="61964-123">Руководство по надстройкам Word</span><span class="sxs-lookup"><span data-stu-id="61964-123">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="61964-124">Справочник по API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="61964-124">Word JavaScript API reference</span></span>](/javascript/api/word)
