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
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a><span data-ttu-id="aeea2-102">Используйте параметры поиска, чтобы найти текст в надстройке Word</span><span class="sxs-lookup"><span data-stu-id="aeea2-102">Use search options to find text in your Word add-in</span></span> 

<span data-ttu-id="aeea2-p101">Надстройки часто должны действовать на основе текста документа. Функция поиска предоставлена для каждого элемента управления содержимым (включает объекты [Текст сообщения](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js) , [Параграф](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js), [Диапазон](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js), [Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js), [TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js)  и базовый объект [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js)). Эта функция принимает строку (или выражение wldcard), представляющую текст, который вы ищете, а также объект [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js). Он возвращает набор диапазонов, которые соответствуют искомому тексту.</span><span class="sxs-lookup"><span data-stu-id="aeea2-p101">Add-ins frequently need to act based on the text of a document. A search function is exposed by every content control (this includes [Body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js), [Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js), [Range](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js), [Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js), [TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js), and the base [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js) object). This function takes in a string (or wldcard expression) representing the text you are searching for and a [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js) object. It returns a collection of ranges which match the search text.</span></span>

## <a name="search-options"></a><span data-ttu-id="aeea2-107">Параметры поиска</span><span class="sxs-lookup"><span data-stu-id="aeea2-107">Search options</span></span>
<span data-ttu-id="aeea2-108">Параметры поиска представляют собой набор логических значений, определяющих способ обработки параметра поиска.</span><span class="sxs-lookup"><span data-stu-id="aeea2-108">The search options are a collection of boolean values defining how the search parameter should be treated.</span></span> 

| <span data-ttu-id="aeea2-109">Свойство</span><span class="sxs-lookup"><span data-stu-id="aeea2-109">Property</span></span>     | <span data-ttu-id="aeea2-110">Описание</span><span class="sxs-lookup"><span data-stu-id="aeea2-110">Description</span></span>|
|:---------------|:----|
|<span data-ttu-id="aeea2-111">ignorePunct</span><span class="sxs-lookup"><span data-stu-id="aeea2-111">ignorePunct</span></span>|<span data-ttu-id="aeea2-p102">Возвращает или задает значение, указывающее, следует ли игнорировать все знаки пунктуации между словами. Соответствует флажку "Игнорировать знаки пунктуации" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="aeea2-p102">Gets or sets a value indicating whether to ignore all punctuation characters between words. Corresponds to the "Ignore punctuation characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="aeea2-114">ignoreSpace</span><span class="sxs-lookup"><span data-stu-id="aeea2-114">ignoreSpace</span></span>|<span data-ttu-id="aeea2-p103">Возвращает или задает значение, указывающее, следует ли игнорировать все знаки пунктуации между словами. Соответствует флажку "Игнорировать символы пробела" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="aeea2-p103">Gets or sets a value indicating whether to ignore all whitespace between words. Corresponds to the "Ignore white-space characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="aeea2-117">matchCase</span><span class="sxs-lookup"><span data-stu-id="aeea2-117">matchCase</span></span>|<span data-ttu-id="aeea2-p104">Возвращает или задает значение, указывающее, следует ли выполнять поиск с учетом регистра. Соответствует флажку "С учетом регистра" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="aeea2-p104">Gets or sets a value indicating whether to perform a case sensitive search. Corresponds to the "Match case" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="aeea2-120">matchPrefix</span><span class="sxs-lookup"><span data-stu-id="aeea2-120">matchPrefix</span></span>|<span data-ttu-id="aeea2-p105">Возвращает или задает значение, указывающее, следует ли сопоставлять слова, начинающиеся со строки поиска. Соответствует флажку "С учетом регистра" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="aeea2-p105">Gets or sets a value indicating whether to match words that begin with the search string. Corresponds to the "Match prefix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="aeea2-123">matchSuffix</span><span class="sxs-lookup"><span data-stu-id="aeea2-123">matchSuffix</span></span>|<span data-ttu-id="aeea2-p106">Возвращает или задает значение, указывающее, следует ли сопоставлять слова, начинающиеся с строки поиска. Соответствует флажку "С учетом регистра" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="aeea2-p106">Gets or sets a value indicating whether to match words that end with the search string. Corresponds to the "Match suffix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="aeea2-126">matchWholeWord</span><span class="sxs-lookup"><span data-stu-id="aeea2-126">matchWholeWord</span></span>|<span data-ttu-id="aeea2-p107">Возвращает или задает значение, указывающее, следует ли находить операции только целые слова, а не текст, который является частью большего слова. Соответствует флажку «Найти только целые слова» в диалоговом окне «Поиск и замена».</span><span class="sxs-lookup"><span data-stu-id="aeea2-p107">Gets or sets a value indicating whether to find operation only entire words, not text that is part of a larger word. Corresponds to the "Find whole words only" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="aeea2-129">matchWildcards</span><span class="sxs-lookup"><span data-stu-id="aeea2-129">matchWildcards</span></span>|<span data-ttu-id="aeea2-p108">Получает или задает значение, указывающее, будет ли поиск выполняться с использованием специальных поисковых операторов. Соответствует флажку "Использовать подстановочные знаки" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="aeea2-p108">Gets or sets a value indicating whether the search will be performed using special search operators. Corresponds to the "Use wildcards" check box in the Find and Replace dialog box.</span></span>|

## <a name="wildcard-guidance"></a><span data-ttu-id="aeea2-132">Руководство по подстановочным знакам</span><span class="sxs-lookup"><span data-stu-id="aeea2-132">Wildcard Guidance</span></span>
<span data-ttu-id="aeea2-133">В следующей таблице приведено руководство по поиску подстановочных знаков API Word JavaScript.</span><span class="sxs-lookup"><span data-stu-id="aeea2-133">The following table provides guidance around the Word JavaScript API’s search wildcards.</span></span>

| <span data-ttu-id="aeea2-134">Чтобы найти:</span><span class="sxs-lookup"><span data-stu-id="aeea2-134">To find:</span></span>         | <span data-ttu-id="aeea2-135">Подстановочный знак</span><span class="sxs-lookup"><span data-stu-id="aeea2-135">Wildcard</span></span> |  <span data-ttu-id="aeea2-136">Пример</span><span class="sxs-lookup"><span data-stu-id="aeea2-136">Sample</span></span> |
|:-----------------|:--------|:----------|
| <span data-ttu-id="aeea2-137">Любой знак</span><span class="sxs-lookup"><span data-stu-id="aeea2-137">Any single character</span></span>| <span data-ttu-id="aeea2-138">?</span><span class="sxs-lookup"><span data-stu-id="aeea2-138">?</span></span> |<span data-ttu-id="aeea2-139">"л?с" находит "лес" и "лис".</span><span class="sxs-lookup"><span data-stu-id="aeea2-139">s?t finds sat and set.</span></span> |
|<span data-ttu-id="aeea2-140">Любая строка знаков</span><span class="sxs-lookup"><span data-stu-id="aeea2-140">Any string of characters</span></span>| * |<span data-ttu-id="aeea2-141">"к\*т" находит "кот" и "компот".</span><span class="sxs-lookup"><span data-stu-id="aeea2-141">s\*d finds sad and started.</span></span>|
|<span data-ttu-id="aeea2-142">Начало слова</span><span class="sxs-lookup"><span data-stu-id="aeea2-142">The beginning of a word</span></span>|< |<span data-ttu-id="aeea2-143">"<(интер)" находит "интересный" и "интермедия", но не "заинтересованный".</span><span class="sxs-lookup"><span data-stu-id="aeea2-143"><(inter) finds interesting and intercept, but not splintered.</span></span>|
|<span data-ttu-id="aeea2-144">Конец слова</span><span class="sxs-lookup"><span data-stu-id="aeea2-144">The end of a word</span></span> |> |<span data-ttu-id="aeea2-145">"(ель)>" находит "ель" и "портфель", но не "ельник".</span><span class="sxs-lookup"><span data-stu-id="aeea2-145">(in)> finds in and within, but not interesting.</span></span>|
|<span data-ttu-id="aeea2-146">Один из указанных знаков</span><span class="sxs-lookup"><span data-stu-id="aeea2-146">One of the specified characters</span></span>|<span data-ttu-id="aeea2-147">[ ]</span><span class="sxs-lookup"><span data-stu-id="aeea2-147">[ ]</span></span> |<span data-ttu-id="aeea2-148">"п[оы]л" находит "пол" и "пыл".</span><span class="sxs-lookup"><span data-stu-id="aeea2-148">w[io]n finds win and won.</span></span>|
|<span data-ttu-id="aeea2-149">Любой символ из этого диапазона</span><span class="sxs-lookup"><span data-stu-id="aeea2-149">Any single character in this range</span></span>| <span data-ttu-id="aeea2-150">[-]</span><span class="sxs-lookup"><span data-stu-id="aeea2-150">[-]</span></span> |<span data-ttu-id="aeea2-p109">"[б-с]оль" находит "боль" и "соль". Диапазон должен быть указан в алфавитном порядке.</span><span class="sxs-lookup"><span data-stu-id="aeea2-p109">[r-t]ight finds right and sight. Ranges must be in ascending order.</span></span>|
|<span data-ttu-id="aeea2-153">Любой символ, кроме символов из диапазона, указанного в скобках</span><span class="sxs-lookup"><span data-stu-id="aeea2-153">Any single character except the characters in the range inside the brackets</span></span>|[!э-я] |<span data-ttu-id="aeea2-155">"ко[!а-п]а" находит "кора" и "коса", но не "коза" или "кожа".</span><span class="sxs-lookup"><span data-stu-id="aeea2-155">t[!a-m]ck finds tock and tuck, but not tack or tick.</span></span>|
|<span data-ttu-id="aeea2-156">Точное количество повторений (n) предыдущего знака или выражения</span><span class="sxs-lookup"><span data-stu-id="aeea2-156">Exactly n occurrences of the previous character or expression</span></span>|<span data-ttu-id="aeea2-157">{n}</span><span class="sxs-lookup"><span data-stu-id="aeea2-157">{n}</span></span> |<span data-ttu-id="aeea2-158">"жарен{2}ый" находит "жаренный", но не "жареный".</span><span class="sxs-lookup"><span data-stu-id="aeea2-158">fe{2}d finds feed but not fed.</span></span>|
|<span data-ttu-id="aeea2-159">Количество повторений предыдущего знака или выражения не менее n раз</span><span class="sxs-lookup"><span data-stu-id="aeea2-159">At least n occurrences of the previous character or expression</span></span>|<span data-ttu-id="aeea2-160">{n,}</span><span class="sxs-lookup"><span data-stu-id="aeea2-160">{n,}</span></span> |<span data-ttu-id="aeea2-161">"жарен{1,}ый" находит "жареный" и "жаренный".</span><span class="sxs-lookup"><span data-stu-id="aeea2-161">fe{1,}d finds fed and feed.</span></span>|
|<span data-ttu-id="aeea2-162">Количество повторений предыдущего знака или выражения в диапазоне от n до m</span><span class="sxs-lookup"><span data-stu-id="aeea2-162">From n to m occurrences of the previous character or expression</span></span>|<span data-ttu-id="aeea2-163">{n,m}</span><span class="sxs-lookup"><span data-stu-id="aeea2-163">{n,m}</span></span> |<span data-ttu-id="aeea2-164">10{1,3} находит 10, 100 и 1000.</span><span class="sxs-lookup"><span data-stu-id="aeea2-164">10{1,3} finds 10, 100, and 1000.</span></span>|
|<span data-ttu-id="aeea2-165">Одно или несколько повторений предыдущего знака или выражения</span><span class="sxs-lookup"><span data-stu-id="aeea2-165">One or more occurrences of the previous character or expression</span></span>|@ |<span data-ttu-id="aeea2-166">"жарен@ый" находит "жареный" и "жаренный".</span><span class="sxs-lookup"><span data-stu-id="aeea2-166">lo@t finds lot and loot.</span></span>|

### <a name="escaping-the-special-characters"></a><span data-ttu-id="aeea2-167">Специальные escape-символы</span><span class="sxs-lookup"><span data-stu-id="aeea2-167">Escaping the special characters</span></span>

<span data-ttu-id="aeea2-p110">Поиск с использованием подстановочных знаков по сути аналогичен поиску по регулярному выражению. В регулярных выражениях используются специальные знаки, в том числе '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!' и '@'. Если один из этих знаков входит в состав искомого строкового литерала, его необходимо отменить, чтобы приложение Word знало, что его необходимо рассматривать буквально, а не как часть логики регулярного выражения. Чтобы отменить символ при поиске с помощью пользовательского интерфейса Word, добавьте перед ним символ '\'. Чтобы отменить символ в программном коде, заключите его в символы '[]'. Например, выражение '[\*]\*' ищет все строки, начинающиеся с символа '\*', за которым следует любое количество других символов.</span><span class="sxs-lookup"><span data-stu-id="aeea2-p110">Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a '\' character, but to escape it programmatically, put it between '[]' characters. For example, '[\*]\*' searches for any string that begins with a '\*' followed by any number of other characters.</span></span> 

## <a name="examples"></a><span data-ttu-id="aeea2-173">Примеры</span><span class="sxs-lookup"><span data-stu-id="aeea2-173">Examples</span></span>
<span data-ttu-id="aeea2-174">Следующие примеры демонстрируют общие сценарии.</span><span class="sxs-lookup"><span data-stu-id="aeea2-174">The following examples demonstrate common scenarios.</span></span>

### <a name="ignore-punctuation-search"></a><span data-ttu-id="aeea2-175">Поиск без учета знаков препинания</span><span class="sxs-lookup"><span data-stu-id="aeea2-175">Ignore punctuation search</span></span>

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

### <a name="search-based-on-a-prefix"></a><span data-ttu-id="aeea2-176">Поиск на основе префикса</span><span class="sxs-lookup"><span data-stu-id="aeea2-176">Search based on a prefix</span></span>

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

### <a name="search-based-on-a-suffix"></a><span data-ttu-id="aeea2-177">Поиск на основе суффикса</span><span class="sxs-lookup"><span data-stu-id="aeea2-177">Search based on a suffix</span></span>

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

### <a name="search-using-a-wildcard"></a><span data-ttu-id="aeea2-178">Поиск с использованием подстановочных знаков</span><span class="sxs-lookup"><span data-stu-id="aeea2-178">Search using a wildcard</span></span>

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

<span data-ttu-id="aeea2-179">Более подробную информацию можно найти в [Справочнике по API Word JavaScript](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="aeea2-179">More information can be found in the [Word JavaScript Reference API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js).</span></span>