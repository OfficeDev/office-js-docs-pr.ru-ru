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
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a>Основные концепции программирования с помощью API JavaScript для Word

В этой статье описаны основные концепции использования [API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md) с целью создания надстроек для Word 2016 и более поздних версий.

## <a name="referencing-officejs"></a>Ссылки на Office.js

Файл Office.js можно получить из указанных ниже расположений.

- `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`. Используйте этот ресурс для рабочих надстроек.

- `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`. Используйте этот ресурс для применения предварительных функций.

## <a name="word-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Word

Наборы требований — это именованные группы элементов API. Надстройки Office используют наборы обязательных элементов, указанных в манифесте, или проверки в среде выполнения, чтобы определить, поддерживает ли клиентское приложение Office необходимые API. Подробнее о наборах обязательных элементов API JavaScript для Word см. в статье [Наборы обязательных элементов API JavaScript для Word](../reference/requirement-sets/word-api-requirement-sets.md).

## <a name="running-word-add-ins"></a>Запуск надстроек Word

Чтобы запустить надстройку, воспользуйтесь обработчиком событий `Office.initialize`Office.initialize. Дополнительные сведения об инициализации надстроек см. в статье [Общие сведения об API](../develop/understanding-the-javascript-api-for-office.md).

Надстройки, предназначенные для Word 2016 и более поздних версий, могут использовать API для Word. Они передают методу `Word.run()` логику взаимодействия с Word в качестве функции. Дополнительные сведения о том, как работать с документом Word в этой модели программирования, см. в статье [Использование модели API для определенных приложений](../develop/application-specific-api-model.md).

В следующем примере показано, как инициализировать и запустить надстройку Word с помощью метода `Word.run()`.

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

## <a name="see-also"></a>См. также

- [Обзор API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md)
- [Создание первой надстройки Word](../quickstarts/word-quickstart.md)
- [Руководство по надстройкам Word](../tutorials/word-tutorial.md)
- [Справочник по API JavaScript для Word](/javascript/api/word)
