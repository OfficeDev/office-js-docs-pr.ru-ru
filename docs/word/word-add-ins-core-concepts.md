---
title: Основные концепции программирования с помощью API JavaScript для Word
description: Создание надстроек для Word с помощью API JavaScript для Word.
ms.date: 07/05/2019
localization_priority: Priority
ms.openlocfilehash: 7849780c1aed48152355c3fdbf350d798b2de1f2
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325019"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a>Основные концепции программирования с помощью API JavaScript для Word

В этой статье описаны основные концепции использования [API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md) с целью создания надстроек для Word 2016 и более поздних версий.

## <a name="referencing-officejs"></a>Ссылки на Office.js

Файл Office.js можно получить из указанных ниже расположений.

- `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`. Используйте этот ресурс для рабочих надстроек.

- `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`. Используйте этот ресурс для применения предварительных функций.

## <a name="word-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Word

Наборы требований — это именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Подробнее о наборах обязательных элементов API JavaScript для Word см. в статье [Наборы обязательных элементов API JavaScript для Word](../reference/requirement-sets/word-api-requirement-sets.md).

## <a name="running-word-add-ins"></a>Запуск надстроек Word

Чтобы запустить надстройку, воспользуйтесь обработчиком событий `Office.initialize`Office.initialize. Дополнительные сведения об инициализации надстроек см. в статье [Общие сведения об API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

Надстройки для Word 2016 и более поздних версий запускаются передачей функции в метод `Word.run()`. Функции, передаваемой в метод `run`, обязательно должен быть присвоен контекстный аргумент. Этот [контекстный объект](/javascript/api/word/word.requestcontext) отличается от контекстного объекта, получаемого из объекта Office, но он также используется для взаимодействия со средой выполнения Word. Контекстный объект предоставляет доступ к объектной модели API JavaScript для Word. В следующем примере показано, как инициализировать и запустить надстройку Word с помощью метода `Word.run()`.

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

### <a name="asynchronous-nature-of-word-apis"></a>Асинхронный характер API Word

API JavaScript для Word загружается с помощью файла Office.js. Этот API изменяет способ взаимодействия с объектами, например с документами и абзацами. Вместо набора отдельных асинхронных API для получения и обновления каждого из этих объектов новый API JavaScript для Word предоставляет прокси-объекты JavaScript, которые соответствуют действующим объектам, используемым в Word. Вы можете взаимодействовать с этими прокси-объектами, синхронно считывая и записывая их свойства, а также вызывая синхронные методы для операций над ними. Эти взаимодействия с прокси-объектами не сразу реализуются в выполняющихся сценариях. Метод `context.sync` синхронизирует состояние запущенного JavaScript и реальных объектов в Office, выполняя поставленные в очередь инструкции и получая свойства загруженных объектов Word для их использования в сценарии.

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>Синхронизация документов Word с помощью прокси-объектов API JavaScript для Word

Объектная модель API JavaScript для Word нестрого связана с объектами в Word. Объекты API JavaScript для Word представляют собой прокси-объекты для объектов в документе Word. Действия, выполняемые над прокси-объектами, не будут реализованы в Word, пока не будет синхронизировано состояние документа. И наоборот, состояние документа Word не будет реализовано в прокси-объектах, пока оно не будет синхронизировано. Чтобы синхронизировать состояние документа, выполните метод `context.sync()`. В примере ниже показано, как создать прокси-объект основного текста и помещенную в очередь команду для загрузки свойства текста в прокси-объекте основного текста и как использовать метод `context.sync()` для синхронизации основного текста документа Word с прокси-объектом основного текста.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    body.load("text");

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a>Выполнение пакета команд

У прокси-объектов Word есть методы для доступа и обновления объектной модели. Эти методы выполняются последовательно в том порядке, в каком они были добавлены в пакет. При вызове метода `context.sync()` выполняются все команды, помещенные в очередь в пакете.

В следующем примере показано, как действует очередь команд. При вызове метода `context.sync()` в Word выполняется команда загрузки основного текста. Затем выполняется команда вставки текста в основной текст в Word. Результаты возвращаются в прокси-объект.основного текста. Значение свойства `body.text` в API JavaScript для Word представляет собой значение основного текста документа Word <u>перед тем, как</u> текст был вставлен в документ Word.

```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    body.load("text");

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="see-also"></a>См. также

- [Обзор API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md)
- [Создание первой надстройки Word](../quickstarts/word-quickstart.md)
- [Руководство по надстройкам Word](../tutorials/word-tutorial.md)
- [Справочник по API JavaScript для Word](/javascript/api/word)