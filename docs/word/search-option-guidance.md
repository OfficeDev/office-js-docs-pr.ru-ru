---
title: Поиск текста в надстройке Word
description: Узнайте, как использовать параметры поиска в надстройке Word.
ms.date: 02/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: e8f9dd2605af9307a49fabfafdecb0df4e97fe9f
ms.sourcegitcommit: 5bf28c447c5b60e2cc7e7a2155db66cd9fe2ab6b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/04/2022
ms.locfileid: "65187345"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a>Поиск текста в надстройке Word

Надстройки часто должны действовать на основе текста документа.
Функция поиска предоставлена для каждого элемента управления содержимым (включает объекты [Body](/javascript/api/word/word.body), [Paragraph](/javascript/api/word/word.paragraph), [Range](/javascript/api/word/word.range), [Table](/javascript/api/word/word.table), [TableRow](/javascript/api/word/word.tablerow), и базовый объект [ContentControl](/javascript/api/word/word.contentcontrol)). Эта функция принимает строку (или подстановочное выражение), представляющую текст, который вы ищете, а также объект [SearchOptions](/javascript/api/word/word.searchoptions). Она возвращает коллекцию диапазонов, соответствующих искомому тексту.

## <a name="search-options"></a>Параметры поиска

Параметры поиска представляют собой коллекцию логических значений, определяющих способ обработки параметра поиска.

| Свойство       | Описание|
|:---------------|:----|
|ignorePunct|Возвращает или задает значение, которое указывает, следует ли пропустить все знаки препинания между словами. Соответствует флажку "Игнорировать знаки препинания" в диалоговом окне "Поиск и замена".|
|ignoreSpace|Возвращает или задает значение, которое указывает, следует ли пропустить все пробелы между словами. Соответствует флажку "Игнорировать пробелы" в диалоговом окне  "Поиск и замена".|
|matchCase|Возвращает или задает значение, указывающее, следует ли выполнять поиск с учетом регистра. Соответствует флажку "Совпадение регистра" в диалоговом окне **"Поиск** и замена".|
|matchPrefix|Возвращает или задает значение, которое указывает, нужно ли учитывать слова, начинающиеся со строки поиска. Соответствует флажку "Сопоставить префикс" в диалоговом окне " **Поиск** и замена".|
|matchSuffix|Возвращает или задает значение, которое указывает, нужно ли учитывать слова, заканчивающиеся строкой поиска. Соответствует флажку "Совпадение суффикса" в диалоговом окне " **Поиск** и замена".|
|matchWholeWord|Возвращает или задает значение, которое указывает, следует ли операции искать только целые слова, а не текст, являющийся частью большего слова. Соответствует флажку "Найти только целые слова" в диалоговом окне " **Найти** и заменить".|
|matchWildcards|Возвращает или задает значение, которое указывает, будет ли выполняться поиск с использованием специальных операторов поиска. Соответствует флажку "Использовать подстановочные знаки" в диалоговом окне " **Поиск** и замена".|

## <a name="wildcard-guidance"></a>Руководство по подстановочным знакам

В таблице ниже приведено руководство по подстановочным знакам поиска в API JavaScript для Word.

| Чтобы найти:         | Подстановочный знак |  Пример |
|:-----------------|:--------|:----------|
|Любой знак| ? |"л?с" находит "лес" и "лис". |
|Любая строка знаков| * |"к*т" находит "кот" и "компот".|
|Начало слова|< |"<(интер)" находит "интересный" и "интермедия", но не "заинтересованный".|
|Конец слова |> |"(ель)>" находит "ель" и "портфель", но не "ельник".|
|Один из указанных знаков|[ ] |"п[оы]л" находит "пол" и "пыл".|
|Любой символ из этого диапазона| [-] |"[б-с]оль" находит "боль" и "соль". Диапазон должен быть указан в алфавитном порядке.|
|Любой символ, кроме символов из диапазона, указанного в скобках|[!э-я] |"ко[!а-п]а" находит "кора" и "коса", но не "коза" или "кожа".|
|Ровно *n* вхождений предыдущего символа или выражения|{n} |"жарен{2}ый" находит "жаренный", но не "жареный".|
|По крайней *мере n* вхождений предыдущего символа или выражения|{n,} |"жарен{1,}ый" находит и "жареный" и "жаренный".|
|От *n* до *m* вхождений предыдущего символа или выражения|{n,m} |10{1,3} находит 10, 100 и 1000.|
|Одно или несколько повторений предыдущего знака или выражения|@ |"жарен@ый" находит "жареный" и "жаренный".|

### <a name="escaping-special-characters"></a>Экранирование специальных символов

Поиск с подстановочными знаками по сути такой же, как поиск по регулярному выражению. В регулярных выражениях есть специальные символы, включая "[", "]", "(", ")", "{", "}", "\*?", "?", "<", ">", "!" и "@". Если один из этих символов является частью строки литерала, которую ищет код, его необходимо экранировать, чтобы Word знал, что его следует рассматривать буквально, а не как часть логики регулярного выражения. Чтобы экранировать символ в поиске пользовательского интерфейса Word, перед ним должен быть символ обратной косой черты ("\\"), но чтобы экранировать его программными средствами, поместите его между символами []. Например, "[\*]\*" выполняет поиск любой строки, которая начинается с символа "\*", за которым следует любое количество других символов.

## <a name="examples"></a>Примеры

В приведенных ниже примерах показаны распространенные сценарии.

### <a name="ignore-punctuation-search"></a>Поиск без учета знаков препинания

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document and ignore punctuation.
    const searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-based-on-a-prefix"></a>Поиск на основе префикса

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document based on a prefix.
    const searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-based-on-a-suffix"></a>Поиск на основе суффикса

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document for any string of characters after 'ly'.
    const searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'orange';
        searchResults.items[i].font.highlightColor = 'black';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-using-a-wildcard"></a>Поиск с использованием подстановочных знаков

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    const searchResults = context.document.body.search('to*n', {matchWildcards: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = 'pink';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

Дополнительные сведения можно найти в [Справочнике по API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md).
