---
title: Поиск текста в надстройке Word
description: ''
ms.date: 07/20/2018
localization_priority: Priority
ms.openlocfilehash: 3e97a9ff41ac2969eddafe8c5b4e762bcc70289b
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386766"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a>Поиск текста в надстройке Word 

Надстройки часто должны действовать на основе текста документа.
Функция поиска предоставлена для каждого элемента управления содержимым (включает объекты [Body](https://docs.microsoft.com/javascript/api/word/word.body), [Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph), [Range](https://docs.microsoft.com/javascript/api/word/word.range), [Table](https://docs.microsoft.com/javascript/api/word/word.table), [TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow), и базовый объект [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol)). Эта функция принимает строку (или подстановочное выражение), представляющую текст, который вы ищете, а также объект [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions). Она возвращает коллекцию диапазонов, соответствующих искомому тексту.

## <a name="search-options"></a>Параметры поиска
Параметры поиска представляют собой коллекцию логических значений, определяющих способ обработки параметра поиска. 

| Свойство     | Описание|
|:---------------|:----|
|ignorePunct|Возвращает или задает значение, которое указывает, следует ли пропустить все знаки препинания между словами. Соответствует установленному флажку "Не учитывать знаки препинания" в диалоговом окне "Найти и заменить".|
|ignoreSpace|Возвращает или задает значение, которое указывает, следует ли пропустить все пробелы между словами. Соответствует установленному флажку "Не учитывать пробелы" в диалоговом окне "Найти и заменить".|
|matchCase|Возвращает или задает значение, которое указывает, следует ли выполнять поиск с учетом регистра. Соответствует установленному флажку "Учитывать регистр" в диалоговом окне "Найти и заменить".|
|matchPrefix|Возвращает или задает значение, которое указывает, нужно ли учитывать слова, начинающиеся со строки поиска. Соответствует установленному флажку "Учитывать префикс" в диалоговом окне "Найти и заменить".|
|matchSuffix|Возвращает или задает значение, которое указывает, нужно ли учитывать слова, заканчивающиеся строкой поиска. Соответствует установленному флажку "Учитывать суффикс" в диалоговом окне "Найти и заменить".|
|matchWholeWord|Возвращает или задает значение, которое указывает, следует ли операции искать только целые слова, а не текст, являющийся частью большего слова. Соответствует установленному флажку "Только слово целиком" в диалоговом окне "Найти и заменить".|
|matchWildcards|Возвращает или задает значение, которое указывает, будет ли выполняться поиск с использованием специальных операторов поиска. Соответствует установленному флажку "Подстановочные знаки" в диалоговом окне "Найти и заменить".|

## <a name="wildcard-guidance"></a>Руководство по подстановочным знакам
В таблице ниже приведено руководство по подстановочным знакам поиска в API JavaScript для Word.

| Чтобы найти:         | Подстановочный знак |  Пример |
|:-----------------|:--------|:----------|
| Любой знак| ? |"л?с" находит "лес" и "лис". |
|Любая строка знаков| * |"к*т" находит "кот" и "компот".|
|Начало слова|< |"<(интер)" находит "интересный" и "интермедия", но не "заинтересованный".|
|Конец слова |> |"(ель)>" находит "ель" и "портфель", но не "ельник".|
|Один из указанных знаков|[ ] |"п[оы]л" находит "пол" и "пыл".|
|Любой символ из этого диапазона| [-] |"[б-с]оль" находит "боль" и "соль". Диапазон должен быть указан в алфавитном порядке.|
|Любой символ, кроме символов из диапазона, указанного в скобках|[!э-я] |"ко[!а-п]а" находит "кора" и "коса", но не "коза" или "кожа".|
|Точное количество повторений (n) предыдущего знака или выражения|{n} |"жарен{2}ый" находит "жаренный", но не "жареный".|
|Количество повторений предыдущего знака или выражения не менее n раз|{n,} |"жарен{1,}ый" находит и "жареный" и "жаренный".|
|Количество повторений предыдущего знака или выражения в диапазоне от n до m|{n,m} |10{1,3} находит 10, 100 и 1000.|
|Одно или несколько повторений предыдущего знака или выражения|@ |"жарен@ый" находит "жареный" и "жаренный".|

### <a name="escaping-the-special-characters"></a>Отмена специальных символов

Поиск с использованием подстановочных знаков по сути аналогичен поиску по регулярному выражению. В регулярных выражениях используются специальные знаки, в том числе '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!' и '@'. Если один из этих знаков входит в состав искомого строкового литерала, его необходимо отменить, чтобы приложение Word знало, что его необходимо рассматривать буквально, а не как часть логики регулярного выражения. Чтобы отменить символ при поиске с помощью пользовательского интерфейса Word, добавьте перед ним символ '\'. Чтобы отменить символ в программном коде, заключите его в символы '[]'. Например, выражение '[\*]\*' ищет все строки, начинающиеся с символа '\*', за которым следует любое количество других символов. 

## <a name="examples"></a>Примеры
В приведенных ниже примерах показаны распространенные сценарии.

### <a name="ignore-punctuation-search"></a>Поиск без учета знаков препинания

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document and ignore punctuation.
    var searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-prefix"></a>Поиск на основе префикса

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document based on a prefix.
    var searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-suffix"></a>Поиск на основе суффикса

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document for any string of characters after 'ly'.
    var searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'orange';
            searchResults.items[i].font.highlightColor = 'black';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-using-a-wildcard"></a>Поиск с использованием подстановочных знаков

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    var searchResults = context.document.body.search('to*n', {matchWildCards: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = 'pink';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

Дополнительные сведения можно найти в [Справочнике по API JavaScript для Word](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview).
