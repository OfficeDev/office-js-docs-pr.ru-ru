---
title: Используйте параметры поиска, чтобы найти текст в надстройке Word
description: ''
ms.date: 7/20/2018
ms.openlocfilehash: ca5c819edb7f3c183379d9df997e41eb56a4de51
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505372"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a>Используйте параметры поиска, чтобы найти текст в надстройке Word 

Надстройки часто должны действовать на основе текста документа. Функция поиска предоставлена для каждого элемента управления содержимым (включает объекты [Текст сообщения](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js) , [Параграф](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js), [Диапазон](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js), [Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js), [TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js)  и базовый объект [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js)). Эта функция принимает строку (или выражение wldcard), представляющую текст, который вы ищете, а также объект [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js). Он возвращает набор диапазонов, которые соответствуют искомому тексту.

## <a name="search-options"></a>Параметры поиска
Параметры поиска представляют собой набор логических значений, определяющих способ обработки параметра поиска. 

| Свойство     | Описание|
|:---------------|:----|
|ignorePunct|Возвращает или задает значение, указывающее, следует ли игнорировать все знаки пунктуации между словами. Соответствует флажку "Игнорировать знаки пунктуации" в диалоговом окне "Поиск и замена".|
|ignoreSpace|Возвращает или задает значение, указывающее, следует ли игнорировать все знаки пунктуации между словами. Соответствует флажку "Игнорировать символы пробела" в диалоговом окне "Поиск и замена".|
|matchCase|Возвращает или задает значение, указывающее, следует ли выполнять поиск с учетом регистра. Соответствует флажку "С учетом регистра" в диалоговом окне "Поиск и замена".|
|matchPrefix|Возвращает или задает значение, указывающее, следует ли сопоставлять слова, начинающиеся со строки поиска. Соответствует флажку "С учетом регистра" в диалоговом окне "Поиск и замена".|
|matchSuffix|Возвращает или задает значение, указывающее, следует ли сопоставлять слова, начинающиеся с строки поиска. Соответствует флажку "С учетом регистра" в диалоговом окне "Поиск и замена".|
|matchWholeWord|Возвращает или задает значение, указывающее, следует ли находить операции только целые слова, а не текст, который является частью большего слова. Соответствует флажку «Найти только целые слова» в диалоговом окне «Поиск и замена».|
|matchWildcards|Получает или задает значение, указывающее, будет ли поиск выполняться с использованием специальных поисковых операторов. Соответствует флажку "Использовать подстановочные знаки" в диалоговом окне "Поиск и замена".|

## <a name="wildcard-guidance"></a>Руководство по подстановочным знакам
В следующей таблице приведено руководство по поиску подстановочных знаков API Word JavaScript.

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
|Количество повторений предыдущего знака или выражения не менее n раз|{n,} |"жарен{1,}ый" находит "жареный" и "жаренный".|
|Количество повторений предыдущего знака или выражения в диапазоне от n до m|{n,m} |10{1,3} находит 10, 100 и 1000.|
|Одно или несколько повторений предыдущего знака или выражения|@ |"жарен@ый" находит "жареный" и "жаренный".|

### <a name="escaping-the-special-characters"></a>Специальные escape-символы

Поиск с использованием подстановочных знаков по сути аналогичен поиску по регулярному выражению. В регулярных выражениях используются специальные знаки, в том числе '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!' и '@'. Если один из этих знаков входит в состав искомого строкового литерала, его необходимо отменить, чтобы приложение Word знало, что его необходимо рассматривать буквально, а не как часть логики регулярного выражения. Чтобы отменить символ при поиске с помощью пользовательского интерфейса Word, добавьте перед ним символ '\'. Чтобы отменить символ в программном коде, заключите его в символы '[]'. Например, выражение '[\*]\*' ищет все строки, начинающиеся с символа '\*', за которым следует любое количество других символов. 

## <a name="examples"></a>Примеры
Следующие примеры демонстрируют общие сценарии.

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

Более подробную информацию можно найти в [Справочнике по API Word JavaScript](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js).