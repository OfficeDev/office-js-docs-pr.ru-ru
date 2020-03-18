---
title: Поиск текста в надстройке Word
description: Сведения об использовании параметров поиска в надстройке Word
ms.date: 09/27/2019
localization_priority: Normal
ms.openlocfilehash: 54ffa3e283f0ae4f43a8d47f7d8cc3a20ea14f6d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717322"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a><span data-ttu-id="99642-103">Поиск текста в надстройке Word</span><span class="sxs-lookup"><span data-stu-id="99642-103">Use search options to find text in your Word add-in</span></span>

<span data-ttu-id="99642-104">Надстройки часто должны действовать на основе текста документа.</span><span class="sxs-lookup"><span data-stu-id="99642-104">Add-ins frequently need to act based on the text of a document.</span></span>
<span data-ttu-id="99642-105">Функция поиска предоставлена для каждого элемента управления содержимым (включает объекты [Body](/javascript/api/word/word.body), [Paragraph](/javascript/api/word/word.paragraph), [Range](/javascript/api/word/word.range), [Table](/javascript/api/word/word.table), [TableRow](/javascript/api/word/word.tablerow), и базовый объект [ContentControl](/javascript/api/word/word.contentcontrol)).</span><span class="sxs-lookup"><span data-stu-id="99642-105">A search function is exposed by every content control (this includes [Body](/javascript/api/word/word.body), [Paragraph](/javascript/api/word/word.paragraph), [Range](/javascript/api/word/word.range), [Table](/javascript/api/word/word.table), [TableRow](/javascript/api/word/word.tablerow), and the base [ContentControl](/javascript/api/word/word.contentcontrol) object).</span></span> <span data-ttu-id="99642-106">Эта функция принимает строку (или подстановочное выражение), представляющую текст, который вы ищете, а также объект [SearchOptions](/javascript/api/word/word.searchoptions).</span><span class="sxs-lookup"><span data-stu-id="99642-106">This function takes in a string (or wildcard expression) representing the text you are searching for and a [SearchOptions](/javascript/api/word/word.searchoptions) object.</span></span> <span data-ttu-id="99642-107">Она возвращает коллекцию диапазонов, соответствующих искомому тексту.</span><span class="sxs-lookup"><span data-stu-id="99642-107">It returns a collection of ranges which match the search text.</span></span>

## <a name="search-options"></a><span data-ttu-id="99642-108">Параметры поиска</span><span class="sxs-lookup"><span data-stu-id="99642-108">Search options</span></span>

<span data-ttu-id="99642-109">Параметры поиска представляют собой коллекцию логических значений, определяющих способ обработки параметра поиска.</span><span class="sxs-lookup"><span data-stu-id="99642-109">The search options are a collection of boolean values defining how the search parameter should be treated.</span></span>

| <span data-ttu-id="99642-110">Свойство</span><span class="sxs-lookup"><span data-stu-id="99642-110">Property</span></span>       | <span data-ttu-id="99642-111">Описание</span><span class="sxs-lookup"><span data-stu-id="99642-111">Description</span></span>|
|:---------------|:----|
|<span data-ttu-id="99642-112">ignorePunct</span><span class="sxs-lookup"><span data-stu-id="99642-112">ignorePunct</span></span>|<span data-ttu-id="99642-113">Возвращает или задает значение, которое указывает, следует ли пропустить все знаки препинания между словами.</span><span class="sxs-lookup"><span data-stu-id="99642-113">Gets or sets a value indicating whether to ignore all punctuation characters between words.</span></span> <span data-ttu-id="99642-114">Соответствует установленному флажку "Не учитывать знаки препинания" в диалоговом окне "Найти и заменить".</span><span class="sxs-lookup"><span data-stu-id="99642-114">Corresponds to the "Ignore punctuation characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="99642-115">ignoreSpace</span><span class="sxs-lookup"><span data-stu-id="99642-115">ignoreSpace</span></span>|<span data-ttu-id="99642-116">Возвращает или задает значение, которое указывает, следует ли пропустить все пробелы между словами.</span><span class="sxs-lookup"><span data-stu-id="99642-116">Gets or sets a value indicating whether to ignore all whitespace between words.</span></span> <span data-ttu-id="99642-117">Соответствует установленному флажку "Не учитывать пробелы" в диалоговом окне "Найти и заменить".</span><span class="sxs-lookup"><span data-stu-id="99642-117">Corresponds to the "Ignore white-space characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="99642-118">matchCase</span><span class="sxs-lookup"><span data-stu-id="99642-118">matchCase</span></span>|<span data-ttu-id="99642-119">Возвращает или задает значение, которое указывает, следует ли выполнять поиск с учетом регистра.</span><span class="sxs-lookup"><span data-stu-id="99642-119">Gets or sets a value indicating whether to perform a case sensitive search.</span></span> <span data-ttu-id="99642-120">Соответствует установленному флажку "Учитывать регистр" в диалоговом окне "Найти и заменить".</span><span class="sxs-lookup"><span data-stu-id="99642-120">Corresponds to the "Match case" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="99642-121">matchPrefix</span><span class="sxs-lookup"><span data-stu-id="99642-121">matchPrefix</span></span>|<span data-ttu-id="99642-122">Возвращает или задает значение, которое указывает, нужно ли учитывать слова, начинающиеся со строки поиска.</span><span class="sxs-lookup"><span data-stu-id="99642-122">Gets or sets a value indicating whether to match words that begin with the search string.</span></span> <span data-ttu-id="99642-123">Соответствует установленному флажку "Учитывать префикс" в диалоговом окне "Найти и заменить".</span><span class="sxs-lookup"><span data-stu-id="99642-123">Corresponds to the "Match prefix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="99642-124">matchSuffix</span><span class="sxs-lookup"><span data-stu-id="99642-124">matchSuffix</span></span>|<span data-ttu-id="99642-125">Возвращает или задает значение, которое указывает, нужно ли учитывать слова, заканчивающиеся строкой поиска.</span><span class="sxs-lookup"><span data-stu-id="99642-125">Gets or sets a value indicating whether to match words that end with the search string.</span></span> <span data-ttu-id="99642-126">Соответствует установленному флажку "Учитывать суффикс" в диалоговом окне "Найти и заменить".</span><span class="sxs-lookup"><span data-stu-id="99642-126">Corresponds to the "Match suffix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="99642-127">matchWholeWord</span><span class="sxs-lookup"><span data-stu-id="99642-127">matchWholeWord</span></span>|<span data-ttu-id="99642-128">Возвращает или задает значение, которое указывает, следует ли операции искать только целые слова, а не текст, являющийся частью большего слова.</span><span class="sxs-lookup"><span data-stu-id="99642-128">Gets or sets a value indicating whether to find operation only entire words, not text that is part of a larger word.</span></span> <span data-ttu-id="99642-129">Соответствует установленному флажку "Только слово целиком" в диалоговом окне "Найти и заменить".</span><span class="sxs-lookup"><span data-stu-id="99642-129">Corresponds to the "Find whole words only" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="99642-130">matchWildcards</span><span class="sxs-lookup"><span data-stu-id="99642-130">matchWildcards</span></span>|<span data-ttu-id="99642-131">Возвращает или задает значение, которое указывает, будет ли выполняться поиск с использованием специальных операторов поиска.</span><span class="sxs-lookup"><span data-stu-id="99642-131">Gets or sets a value indicating whether the search will be performed using special search operators.</span></span> <span data-ttu-id="99642-132">Соответствует установленному флажку "Подстановочные знаки" в диалоговом окне "Найти и заменить".</span><span class="sxs-lookup"><span data-stu-id="99642-132">Corresponds to the "Use wildcards" check box in the Find and Replace dialog box.</span></span>|

## <a name="wildcard-guidance"></a><span data-ttu-id="99642-133">Руководство по подстановочным знакам</span><span class="sxs-lookup"><span data-stu-id="99642-133">Wildcard guidance</span></span>

<span data-ttu-id="99642-134">В таблице ниже приведено руководство по подстановочным знакам поиска в API JavaScript для Word.</span><span class="sxs-lookup"><span data-stu-id="99642-134">The following table provides guidance around the Word JavaScript API's search wildcards.</span></span>

| <span data-ttu-id="99642-135">Чтобы найти:</span><span class="sxs-lookup"><span data-stu-id="99642-135">To find:</span></span>         | <span data-ttu-id="99642-136">Подстановочный знак</span><span class="sxs-lookup"><span data-stu-id="99642-136">Wildcard</span></span> |  <span data-ttu-id="99642-137">Пример</span><span class="sxs-lookup"><span data-stu-id="99642-137">Sample</span></span> |
|:-----------------|:--------|:----------|
| <span data-ttu-id="99642-138">Любой знак</span><span class="sxs-lookup"><span data-stu-id="99642-138">Any single character</span></span>| <span data-ttu-id="99642-139">?</span><span class="sxs-lookup"><span data-stu-id="99642-139">?</span></span> |<span data-ttu-id="99642-140">"л?с" находит "лес" и "лис".</span><span class="sxs-lookup"><span data-stu-id="99642-140">s?t finds sat and set.</span></span> |
|<span data-ttu-id="99642-141">Любая строка знаков</span><span class="sxs-lookup"><span data-stu-id="99642-141">Any string of characters</span></span>| * |<span data-ttu-id="99642-142">"к\*т" находит "кот" и "компот".</span><span class="sxs-lookup"><span data-stu-id="99642-142">s\*d finds sad and started.</span></span>|
|<span data-ttu-id="99642-143">Начало слова</span><span class="sxs-lookup"><span data-stu-id="99642-143">The beginning of a word</span></span>|< |<span data-ttu-id="99642-144">"<(интер)" находит "интересный" и "интермедия", но не "заинтересованный".</span><span class="sxs-lookup"><span data-stu-id="99642-144"><(inter) finds interesting and intercept, but not splintered.</span></span>|
|<span data-ttu-id="99642-145">Конец слова</span><span class="sxs-lookup"><span data-stu-id="99642-145">The end of a word</span></span> |> |<span data-ttu-id="99642-146">"(ель)>" находит "ель" и "портфель", но не "ельник".</span><span class="sxs-lookup"><span data-stu-id="99642-146">(in)> finds in and within, but not interesting.</span></span>|
|<span data-ttu-id="99642-147">Один из указанных знаков</span><span class="sxs-lookup"><span data-stu-id="99642-147">One of the specified characters</span></span>|<span data-ttu-id="99642-148">[ ]</span><span class="sxs-lookup"><span data-stu-id="99642-148">[ ]</span></span> |<span data-ttu-id="99642-149">"п[оы]л" находит "пол" и "пыл".</span><span class="sxs-lookup"><span data-stu-id="99642-149">w[io]n finds win and won.</span></span>|
|<span data-ttu-id="99642-150">Любой символ из этого диапазона</span><span class="sxs-lookup"><span data-stu-id="99642-150">Any single character in this range</span></span>| <span data-ttu-id="99642-151">[-]</span><span class="sxs-lookup"><span data-stu-id="99642-151">[-]</span></span> |<span data-ttu-id="99642-p109">"[б-с]оль" находит "боль" и "соль". Диапазон должен быть указан в алфавитном порядке.</span><span class="sxs-lookup"><span data-stu-id="99642-p109">[r-t]ight finds right and sight. Ranges must be in ascending order.</span></span>|
|<span data-ttu-id="99642-154">Любой символ, кроме символов из диапазона, указанного в скобках</span><span class="sxs-lookup"><span data-stu-id="99642-154">Any single character except the characters in the range inside the brackets</span></span>|[!э-я] |<span data-ttu-id="99642-156">"ко[!а-п]а" находит "кора" и "коса", но не "коза" или "кожа".</span><span class="sxs-lookup"><span data-stu-id="99642-156">t[!a-m]ck finds tock and tuck, but not tack or tick.</span></span>|
|<span data-ttu-id="99642-157">Точное количество повторений (n) предыдущего знака или выражения</span><span class="sxs-lookup"><span data-stu-id="99642-157">Exactly n occurrences of the previous character or expression</span></span>|<span data-ttu-id="99642-158">{n}</span><span class="sxs-lookup"><span data-stu-id="99642-158">{n}</span></span> |<span data-ttu-id="99642-159">"жарен{2}ый" находит "жаренный", но не "жареный".</span><span class="sxs-lookup"><span data-stu-id="99642-159">fe{2}d finds feed but not fed.</span></span>|
|<span data-ttu-id="99642-160">Количество повторений предыдущего знака или выражения не менее n раз</span><span class="sxs-lookup"><span data-stu-id="99642-160">At least n occurrences of the previous character or expression</span></span>|<span data-ttu-id="99642-161">{n,}</span><span class="sxs-lookup"><span data-stu-id="99642-161">{n,}</span></span> |<span data-ttu-id="99642-162">"жарен{1,}ый" находит и "жареный" и "жаренный".</span><span class="sxs-lookup"><span data-stu-id="99642-162">fe{1,}d finds fed and feed.</span></span>|
|<span data-ttu-id="99642-163">Количество повторений предыдущего знака или выражения в диапазоне от n до m</span><span class="sxs-lookup"><span data-stu-id="99642-163">From n to m occurrences of the previous character or expression</span></span>|<span data-ttu-id="99642-164">{n,m}</span><span class="sxs-lookup"><span data-stu-id="99642-164">{n,m}</span></span> |<span data-ttu-id="99642-165">10{1,3} находит 10, 100 и 1000.</span><span class="sxs-lookup"><span data-stu-id="99642-165">10{1,3} finds 10, 100, and 1000.</span></span>|
|<span data-ttu-id="99642-166">Одно или несколько повторений предыдущего знака или выражения</span><span class="sxs-lookup"><span data-stu-id="99642-166">One or more occurrences of the previous character or expression</span></span>|@ |<span data-ttu-id="99642-167">"жарен@ый" находит "жареный" и "жаренный".</span><span class="sxs-lookup"><span data-stu-id="99642-167">lo@t finds lot and loot.</span></span>|

### <a name="escaping-the-special-characters"></a><span data-ttu-id="99642-168">Отмена специальных символов</span><span class="sxs-lookup"><span data-stu-id="99642-168">Escaping the special characters</span></span>

<span data-ttu-id="99642-p110">Поиск с использованием подстановочных знаков по сути аналогичен поиску по регулярному выражению. В регулярных выражениях используются специальные знаки, в том числе '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!' и '@'. Если один из этих знаков входит в состав искомого строкового литерала, его необходимо отменить, чтобы приложение Word знало, что его необходимо рассматривать буквально, а не как часть логики регулярного выражения. Чтобы отменить символ при поиске с помощью пользовательского интерфейса Word, добавьте перед ним символ '\'. Чтобы отменить символ в программном коде, заключите его в символы '[]'. Например, выражение '[\*]\*' ищет все строки, начинающиеся с символа '\*', за которым следует любое количество других символов.</span><span class="sxs-lookup"><span data-stu-id="99642-p110">Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a '\' character, but to escape it programmatically, put it between '[]' characters. For example, '[\*]\*' searches for any string that begins with a '\*' followed by any number of other characters.</span></span> 

## <a name="examples"></a><span data-ttu-id="99642-174">Примеры</span><span class="sxs-lookup"><span data-stu-id="99642-174">Examples</span></span>

<span data-ttu-id="99642-175">В приведенных ниже примерах показаны распространенные сценарии.</span><span class="sxs-lookup"><span data-stu-id="99642-175">The following examples demonstrate common scenarios.</span></span>

### <a name="ignore-punctuation-search"></a><span data-ttu-id="99642-176">Поиск без учета знаков препинания</span><span class="sxs-lookup"><span data-stu-id="99642-176">Ignore punctuation search</span></span>

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

### <a name="search-based-on-a-prefix"></a><span data-ttu-id="99642-177">Поиск на основе префикса</span><span class="sxs-lookup"><span data-stu-id="99642-177">Search based on a prefix</span></span>

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

### <a name="search-based-on-a-suffix"></a><span data-ttu-id="99642-178">Поиск на основе суффикса</span><span class="sxs-lookup"><span data-stu-id="99642-178">Search based on a suffix</span></span>

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

### <a name="search-using-a-wildcard"></a><span data-ttu-id="99642-179">Поиск с использованием подстановочных знаков</span><span class="sxs-lookup"><span data-stu-id="99642-179">Search using a wildcard</span></span>

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

<span data-ttu-id="99642-180">Дополнительные сведения можно найти в [Справочнике по API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md).</span><span class="sxs-lookup"><span data-stu-id="99642-180">More information can be found in the [Word JavaScript Reference API](../reference/overview/word-add-ins-reference-overview.md).</span></span>
