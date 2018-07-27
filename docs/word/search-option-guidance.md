---
title: Используйте параметры поиска, чтобы найти текст в надстройке Word
description: ''
ms.date: 7/20/2018
ms.openlocfilehash: 9dcd5e42de9cc0816797a4a14b40a0e3e376f158
ms.sourcegitcommit: eea7f2b1679cf9a209d35880b906e311bdf1359c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2018
ms.locfileid: "21254863"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a><span data-ttu-id="76d71-102">Используйте параметры поиска, чтобы найти текст в надстройке Word</span><span class="sxs-lookup"><span data-stu-id="76d71-102">Use search options to find text in your Word add-in</span></span> 

<span data-ttu-id="76d71-103">Надстройки часто должны действовать на основе текста документа.</span><span class="sxs-lookup"><span data-stu-id="76d71-103">Add-ins frequently need to act based on the text of a document.</span></span>
<span data-ttu-id="76d71-104">Функция поиска предоставлена для каждого элемента управления содержимым (включает объекты [Текст сообщения](https://dev.office.com/reference/add-ins/word/body), [Параграф](https://dev.office.com/reference/add-ins/word/paragraph), [Диапазон](https://dev.office.com/reference/add-ins/word/range), [Таблица](https://dev.office.com/reference/add-ins/word/table), [TableRow](https://dev.office.com/reference/add-ins/word/tablerow), и базовый объект[ContentControl](https://dev.office.com/reference/add-ins/word/contentcontrol)).</span><span class="sxs-lookup"><span data-stu-id="76d71-104">A search function is exposed by every content control (this includes [Body](https://dev.office.com/reference/add-ins/word/body), [Paragraph](https://dev.office.com/reference/add-ins/word/paragraph), [Range](https://dev.office.com/reference/add-ins/word/range), [Table](https://dev.office.com/reference/add-ins/word/table), [TableRow](https://dev.office.com/reference/add-ins/word/tablerow), and the base [ContentControl](https://dev.office.com/reference/add-ins/word/contentcontrol) object).</span></span> <span data-ttu-id="76d71-105">Эта функция принимает строку (или выражение wldcard), представляющую текст, который вы ищете, а также объект[SearchOptions](https://dev.office.com/reference/add-ins/word/searchoptions).</span><span class="sxs-lookup"><span data-stu-id="76d71-105">This function takes in a string (or wldcard expression) representing the text you are searching for and a [SearchOptions](https://dev.office.com/reference/add-ins/word/searchoptions) object.</span></span> <span data-ttu-id="76d71-106">Он возвращает набор диапазонов, которые соответствуют искомому тексту.</span><span class="sxs-lookup"><span data-stu-id="76d71-106">It returns a collection of ranges which match the search text.</span></span>

## <a name="search-options"></a><span data-ttu-id="76d71-107">Параметры поиска</span><span class="sxs-lookup"><span data-stu-id="76d71-107">Search options</span></span>
<span data-ttu-id="76d71-108">Параметры поиска представляют собой набор логических значений, определяющих способ обработки параметра поиска.</span><span class="sxs-lookup"><span data-stu-id="76d71-108">The search options are a collection of boolean values defining how the search parameter should be treated.</span></span> 

| <span data-ttu-id="76d71-109">Свойство</span><span class="sxs-lookup"><span data-stu-id="76d71-109">Property</span></span>     | <span data-ttu-id="76d71-110">Описание</span><span class="sxs-lookup"><span data-stu-id="76d71-110">Description</span></span>|
|:---------------|:----|
|<span data-ttu-id="76d71-111">ignorePunct</span><span class="sxs-lookup"><span data-stu-id="76d71-111">ignorePunct</span></span>|<span data-ttu-id="76d71-112">Возвращает или задает значение, указывающее, следует ли игнорировать все знаки пунктуации между словами.</span><span class="sxs-lookup"><span data-stu-id="76d71-112">Gets or sets a value indicating whether to ignore all punctuation characters between words.</span></span> <span data-ttu-id="76d71-113">Соответствует флажку "Игнорировать знаки пунктуации" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="76d71-113">Corresponds to the "Ignore punctuation characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="76d71-114">ignoreSpace</span><span class="sxs-lookup"><span data-stu-id="76d71-114">ignoreSpace</span></span>|<span data-ttu-id="76d71-115">Возвращает или задает значение, указывающее, следует ли игнорировать все знаки пунктуации между словами.</span><span class="sxs-lookup"><span data-stu-id="76d71-115">Gets or sets a value indicating whether to ignore all whitespace between words.</span></span> <span data-ttu-id="76d71-116">Соответствует флажку "Игнорировать символы пробела" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="76d71-116">Corresponds to the "Ignore white-space characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="76d71-117">matchCase</span><span class="sxs-lookup"><span data-stu-id="76d71-117">matchCase</span></span>|<span data-ttu-id="76d71-118">Возвращает или задает значение, указывающее, следует ли выполнять поиск с учетом регистра.</span><span class="sxs-lookup"><span data-stu-id="76d71-118">Gets or sets a value indicating whether to perform a case sensitive search.</span></span> <span data-ttu-id="76d71-119">Соответствует флажку "С учетом регистра" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="76d71-119">Corresponds to the Sounds like check box in the Find and Replace dialog box</span></span>|
|<span data-ttu-id="76d71-120">matchPrefix</span><span class="sxs-lookup"><span data-stu-id="76d71-120">matchPrefix</span></span>|<span data-ttu-id="76d71-121">Возвращает или задает значение, указывающее, следует ли сопоставлять слова, начинающиеся со строки поиска.</span><span class="sxs-lookup"><span data-stu-id="76d71-121">Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="76d71-122">Соответствует флажку "С учетом регистра" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="76d71-122">Corresponds to the Sounds like check box in the Find and Replace dialog box</span></span>|
|<span data-ttu-id="76d71-123">matchSuffix</span><span class="sxs-lookup"><span data-stu-id="76d71-123">matchSuffix</span></span>|<span data-ttu-id="76d71-124">Возвращает или задает значение, указывающее, следует ли сопоставлять слова, начинающиеся с строки поиска.</span><span class="sxs-lookup"><span data-stu-id="76d71-124">Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="76d71-125">Соответствует флажку "С учетом регистра" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="76d71-125">Corresponds to the Sounds like check box in the Find and Replace dialog box</span></span>|
|<span data-ttu-id="76d71-126">matchWholeWord</span><span class="sxs-lookup"><span data-stu-id="76d71-126">matchWholeWord</span></span>|<span data-ttu-id="76d71-127">Возвращает или задает значение, указывающее, следует ли находить операции только целые слова, а не текст, который является частью большего слова.</span><span class="sxs-lookup"><span data-stu-id="76d71-127">Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="76d71-128">Соответствует флажку «Найти только целые слова» в диалоговом окне «Поиск и замена».</span><span class="sxs-lookup"><span data-stu-id="76d71-128">Corresponds to the Sounds like check box in the Find and Replace dialog box</span></span>|
|<span data-ttu-id="76d71-129">matchWildcards</span><span class="sxs-lookup"><span data-stu-id="76d71-129">matchWildcards</span></span>|<span data-ttu-id="76d71-130">Получает или задает значение, указывающее, будет ли поиск выполняться с использованием специальных поисковых операторов.</span><span class="sxs-lookup"><span data-stu-id="76d71-130">Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.</span></span> <span data-ttu-id="76d71-131">Соответствует флажку "Использовать подстановочные знаки" в диалоговом окне "Поиск и замена".</span><span class="sxs-lookup"><span data-stu-id="76d71-131">Corresponds to the Sounds like check box in the Find and Replace dialog box</span></span>|

## <a name="wildcard-guidance"></a><span data-ttu-id="76d71-132">Руководство по подстановочным знакам</span><span class="sxs-lookup"><span data-stu-id="76d71-132">Wildcard Guidance</span></span>
<span data-ttu-id="76d71-133">В следующей таблице приведено руководство по поиску подстановочных знаков API Word JavaScript.</span><span class="sxs-lookup"><span data-stu-id="76d71-133">The following table provides guidance around the Word JavaScript API’s search wildcards.</span></span>

| <span data-ttu-id="76d71-134">Чтобы найти:</span><span class="sxs-lookup"><span data-stu-id="76d71-134">To find:</span></span>         | <span data-ttu-id="76d71-135">Подстановочный знак</span><span class="sxs-lookup"><span data-stu-id="76d71-135">Wildcard</span></span> |  <span data-ttu-id="76d71-136">Пример</span><span class="sxs-lookup"><span data-stu-id="76d71-136">Sample</span></span> |
|:-----------------|:--------|:----------|
| <span data-ttu-id="76d71-137">Любой знак</span><span class="sxs-lookup"><span data-stu-id="76d71-137">Any single character</span></span>| <span data-ttu-id="76d71-138">?</span><span class="sxs-lookup"><span data-stu-id="76d71-138"></span></span> |<span data-ttu-id="76d71-139">"л?с" находит "лес" и "лис".</span><span class="sxs-lookup"><span data-stu-id="76d71-139">s?t finds sat and set.</span></span> |
|<span data-ttu-id="76d71-140">Любая строка знаков</span><span class="sxs-lookup"><span data-stu-id="76d71-140">Any string of characters</span></span>| * |<span data-ttu-id="76d71-141">"к\*т" находит "кот" и "компот".</span><span class="sxs-lookup"><span data-stu-id="76d71-141">s\*d finds sad and started.</span></span>|
|<span data-ttu-id="76d71-142">Начало слова</span><span class="sxs-lookup"><span data-stu-id="76d71-142">The beginning of a word</span></span>|< |<span data-ttu-id="76d71-143">"<(интер)" находит "интересный" и "интермедия", но не "заинтересованный".</span><span class="sxs-lookup"><span data-stu-id="76d71-143"><(inter) finds interesting and intercept, but not splintered.</span></span>|
|<span data-ttu-id="76d71-144">Конец слова</span><span class="sxs-lookup"><span data-stu-id="76d71-144">The end of a word</span></span> |> |<span data-ttu-id="76d71-145">"(ель)>" находит "ель" и "портфель", но не "ельник".</span><span class="sxs-lookup"><span data-stu-id="76d71-145">(in)> finds in and within, but not interesting.</span></span>|
|<span data-ttu-id="76d71-146">Один из указанных знаков</span><span class="sxs-lookup"><span data-stu-id="76d71-146">One of the specified characters</span></span>|<span data-ttu-id="76d71-147">[ ]</span><span class="sxs-lookup"><span data-stu-id="76d71-147"></span></span> |<span data-ttu-id="76d71-148">"п[оы]л" находит "пол" и "пыл".</span><span class="sxs-lookup"><span data-stu-id="76d71-148">w[io]n finds win and won.</span></span>|
|<span data-ttu-id="76d71-149">Любой символ из этого диапазона</span><span class="sxs-lookup"><span data-stu-id="76d71-149">Any single character in this range</span></span>| <span data-ttu-id="76d71-150">[-]</span><span class="sxs-lookup"><span data-stu-id="76d71-150"></span></span> |<span data-ttu-id="76d71-p109">"[б-с]оль" находит "боль" и "соль". Диапазон должен быть указан в алфавитном порядке.</span><span class="sxs-lookup"><span data-stu-id="76d71-p109">[r-t]ight finds right and sight. Ranges must be in ascending order.</span></span>|
|<span data-ttu-id="76d71-153">Любой символ, кроме символов из диапазона, указанного в скобках</span><span class="sxs-lookup"><span data-stu-id="76d71-153">Any single character except the characters in the range inside the brackets</span></span>|[!э-я] |<span data-ttu-id="76d71-155">"ко[!а-п]а" находит "кора" и "коса", но не "коза" или "кожа".</span><span class="sxs-lookup"><span data-stu-id="76d71-155">t[!a-m]ck finds tock and tuck, but not tack or tick.</span></span>|
|<span data-ttu-id="76d71-156">Точное количество повторений (n) предыдущего знака или выражения</span><span class="sxs-lookup"><span data-stu-id="76d71-156">Exactly n occurrences of the previous character or expression</span></span>|<span data-ttu-id="76d71-157">{n}</span><span class="sxs-lookup"><span data-stu-id="76d71-157">{n}</span></span> |<span data-ttu-id="76d71-158">"жарен{2}ый" находит "жаренный", но не "жареный".</span><span class="sxs-lookup"><span data-stu-id="76d71-158">fe{2}d finds feed but not fed.</span></span>|
|<span data-ttu-id="76d71-159">Количество повторений предыдущего знака или выражения не менее n раз</span><span class="sxs-lookup"><span data-stu-id="76d71-159">At least n occurrences of the previous character or expression</span></span>|<span data-ttu-id="76d71-160">{n,}</span><span class="sxs-lookup"><span data-stu-id="76d71-160">{n,}</span></span> |<span data-ttu-id="76d71-161">"жарен{1,} ый" находит "жареный" и "жаренный".</span><span class="sxs-lookup"><span data-stu-id="76d71-161">fe{1,}d finds fed and feed.</span></span>|
|<span data-ttu-id="76d71-162">Количество повторений предыдущего знака или выражения в диапазоне от n до m</span><span class="sxs-lookup"><span data-stu-id="76d71-162">From n to m occurrences of the previous character or expression</span></span>|<span data-ttu-id="76d71-163">{n,m}</span><span class="sxs-lookup"><span data-stu-id="76d71-163">{n,m}</span></span> |<span data-ttu-id="76d71-164">10{1,3} находит 10, 100 и 1000.</span><span class="sxs-lookup"><span data-stu-id="76d71-164">10{1,3} finds 10, 100, and 1000.</span></span>|
|<span data-ttu-id="76d71-165">Одно или несколько повторений предыдущего знака или выражения</span><span class="sxs-lookup"><span data-stu-id="76d71-165">One or more occurrences of the previous character or expression</span></span>|@ |<span data-ttu-id="76d71-166">"жарен@ый" находит "жареный" и "жаренный".</span><span class="sxs-lookup"><span data-stu-id="76d71-166">lo@t finds lot and loot.</span></span>|

### <a name="escaping-the-special-characters"></a><span data-ttu-id="76d71-167">Отмена специальных символов</span><span class="sxs-lookup"><span data-stu-id="76d71-167">Escaping the special characters</span></span>

<span data-ttu-id="76d71-p110">Поиск с использованием подстановочных знаков по сути аналогичен поиску по регулярному выражению. В регулярных выражениях используются специальные знаки, в том числе '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!' и '@'. Если один из этих знаков входит в состав искомого строкового литерала, его необходимо отменить, чтобы приложение Word знало, что его необходимо рассматривать буквально, а не как часть логики регулярного выражения. Чтобы отменить символ при поиске с помощью пользовательского интерфейса Word, добавьте перед ним символ '\'. Чтобы отменить символ в программном коде, заключите его в символы '[]'. Например, выражение '[\*]\*' ищет все строки, начинающиеся с символа '\*', за которым следует любое количество других символов.</span><span class="sxs-lookup"><span data-stu-id="76d71-p110">Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a '\' character, but to escape it programmatically, put it between '[]' characters. For example, '[\*]\*' searches for any string that begins with a '\*' followed by any number of other characters.</span></span> 

## <a name="examples"></a><span data-ttu-id="76d71-173">Примеры</span><span class="sxs-lookup"><span data-stu-id="76d71-173">Examples</span></span>
<span data-ttu-id="76d71-174">Следующие примеры демонстрируют общие сценарии.</span><span class="sxs-lookup"><span data-stu-id="76d71-174">The following examples demonstrate common scenarios.</span></span>

### <a name="ignore-punctuation-search"></a><span data-ttu-id="76d71-175">Поиск без учета знаков препинания</span><span class="sxs-lookup"><span data-stu-id="76d71-175">Ignore punctuation search</span></span>

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

### <a name="search-based-on-a-prefix"></a><span data-ttu-id="76d71-176">Поиск на основе префикса</span><span class="sxs-lookup"><span data-stu-id="76d71-176">Search based on a prefix</span></span>

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

### <a name="search-based-on-a-suffix"></a><span data-ttu-id="76d71-177">Поиск на основе суффикса</span><span class="sxs-lookup"><span data-stu-id="76d71-177">Search based on a suffix</span></span>

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

### <a name="search-using-a-wildcard"></a><span data-ttu-id="76d71-178">Поиск с использованием подстановочных знаков</span><span class="sxs-lookup"><span data-stu-id="76d71-178">Search using a wildcard</span></span>

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

<span data-ttu-id="76d71-179">Более подробную информацию можно найти в [Справочнике по API Word JavaScript](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="76d71-179">More information can be found in the [Word JavaScript Reference API](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview).</span></span>