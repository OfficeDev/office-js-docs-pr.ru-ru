---
title: Избегайте использования метода context.sync в циклах
description: Узнайте, как использовать раздельный цикл и сопоставленные шаблоны объектов, чтобы не вызывать context.sync в цикле.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 64cfd5cd350746ba07e1a98986a4bd7811431475
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349142"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a><span data-ttu-id="ef38f-103">Избегайте использования метода context.sync в циклах</span><span class="sxs-lookup"><span data-stu-id="ef38f-103">Avoid using the context.sync method in loops</span></span>

> [!NOTE]
> <span data-ttu-id="ef38f-104">В этой статье предполагается, что вы не на самом начале работы по крайней мере с одним из четырех API javaScript Office для Excel, Word, OneNote и Visio, которые используют пакетную систему для взаимодействия с Office &mdash; &mdash; документом.</span><span class="sxs-lookup"><span data-stu-id="ef38f-104">This article assumes that you're beyond the beginning stage of working with at least one of the four application-specific Office JavaScript APIs&mdash;for Excel, Word, OneNote, and Visio&mdash;that use a batch system to interact with the Office document.</span></span> <span data-ttu-id="ef38f-105">В частности, необходимо знать, что такое вызов, и знать, что такое `context.sync` объект коллекции.</span><span class="sxs-lookup"><span data-stu-id="ef38f-105">In particular, you should know what a call of `context.sync` does and you should know what a collection object is.</span></span> <span data-ttu-id="ef38f-106">Если вы еще не на этом этапе, начните с Office [API JavaScript](../develop/understanding-the-javascript-api-for-office.md) и документации, связанной с "конкретным приложением" в этой статье.</span><span class="sxs-lookup"><span data-stu-id="ef38f-106">If you're not at that stage, please start with [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md) and the documentation linked to under "application-specific" in that article.</span></span>

<span data-ttu-id="ef38f-107">Для некоторых сценариев программирования в Office надстройки, которые используют одну из моделей API, определенных приложениям (для Excel, Word, OneNote и Visio), коду необходимо прочитать, написать или обработать некоторые свойства от каждого члена объекта коллекции.</span><span class="sxs-lookup"><span data-stu-id="ef38f-107">For some programming scenarios in Office Add-ins that use one of the application-specific API models (for Excel, Word, OneNote, and Visio), your code needs to read, write, or process some property from every member of a collection object.</span></span> <span data-ttu-id="ef38f-108">Например, надстройка Excel, которая должна получать значения каждой ячейки в определенном столбце таблицы или надстройку Word, которая должна выделить каждый экземпляр строки в документе.</span><span class="sxs-lookup"><span data-stu-id="ef38f-108">For example, an Excel add-in that needs to get the values of every cell in a particular table column or a Word add-in that needs to highlight every instance of a string in the document.</span></span> <span data-ttu-id="ef38f-109">Необходимо итерировать над участниками в свойстве объекта коллекции; но по соображениям производительности следует избегать вызовов во всех `items` `context.sync` итерациях цикла.</span><span class="sxs-lookup"><span data-stu-id="ef38f-109">You need to iterate over the members in the `items` property of the collection object; but, for performance reasons, you need to avoid calling `context.sync` in every iteration of the loop.</span></span> <span data-ttu-id="ef38f-110">Каждый вызов `context.sync` — это круговая поездка из надстройки в Office документа.</span><span class="sxs-lookup"><span data-stu-id="ef38f-110">Every call of `context.sync` is a round trip from the add-in to the Office document.</span></span> <span data-ttu-id="ef38f-111">Повторяемая круговая поездка повредит производительности, особенно если надстройка работает в Office в Интернете, так как круглая поездка идет через Интернет.</span><span class="sxs-lookup"><span data-stu-id="ef38f-111">Repeated round trips hurt performance, especially if the add-in is running in Office on the web because the round trips go across the internet.</span></span>

> [!NOTE]
> <span data-ttu-id="ef38f-112">Во всех примерах этой статьи используются циклы, но описанные практики применимы к любому заявлению цикла, которое может итерировать через массив, включая `for` следующие:</span><span class="sxs-lookup"><span data-stu-id="ef38f-112">All examples in this article use `for` loops but the practices described apply to any loop statement that can iterate through an array, including the following:</span></span>
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> <span data-ttu-id="ef38f-113">Они также применяются к любому методу массива, к которому функция передается и применяется к элементу массива, включая следующие:</span><span class="sxs-lookup"><span data-stu-id="ef38f-113">They also apply to any array method to which a function is passed and applied to the items in the array, including the following:</span></span>
>
> - `Array.every`
> - `Array.forEach`
> - `Array.filter`
> - `Array.find`
> - `Array.findIndex`
> - `Array.map`
> - `Array.reduce`
> - `Array.reduceRight`
> - `Array.some`

## <a name="writing-to-the-document"></a><span data-ttu-id="ef38f-114">Запись в документ</span><span class="sxs-lookup"><span data-stu-id="ef38f-114">Writing to the document</span></span>

<span data-ttu-id="ef38f-115">В простейшем случае вы пишете только членам объекта коллекции, не читая их свойства.</span><span class="sxs-lookup"><span data-stu-id="ef38f-115">In the simplest case, you are only writing to members of a collection object, not reading their properties.</span></span> <span data-ttu-id="ef38f-116">Например, следующий код выделяет желтым цветом каждый экземпляр "" в документе Word.</span><span class="sxs-lookup"><span data-stu-id="ef38f-116">For example, the following code highlights in yellow every instance of "the" in a Word document.</span></span>

> [!NOTE]
> <span data-ttu-id="ef38f-117">Как правило, перед закрытием "}" символа метода приложения (например, и т. `context.sync` `run` `Excel.run` д.) необходимо поставить окончательный `Word.run` результат.</span><span class="sxs-lookup"><span data-stu-id="ef38f-117">It is generally a good practice to put have a final `context.sync` just before the closing "}" character of the application `run` method (such as `Excel.run`, `Word.run`, etc.).</span></span> <span data-ttu-id="ef38f-118">Это происходит потому, что метод делает скрытый вызов как последнее, что он делает, если, и только если есть очереди команд, которые еще не `run` `context.sync` были синхронизированы.</span><span class="sxs-lookup"><span data-stu-id="ef38f-118">This is because the `run` method makes a hidden call of `context.sync` as the last thing it does if, and only if, there are queued commands that have not yet been synchronized.</span></span> <span data-ttu-id="ef38f-119">Тот факт, что этот вызов скрыт, может привести к путанице, поэтому обычно рекомендуется добавить явный `context.sync` .</span><span class="sxs-lookup"><span data-stu-id="ef38f-119">The fact that this call is hidden can be confusing, so we generally recommend that you add the explicit `context.sync`.</span></span> <span data-ttu-id="ef38f-120">Однако, учитывая, что в этой статье речь идет о минимизации вызовов, это на самом деле более запутанным, чтобы добавить совершенно `context.sync` ненужный окончательный `context.sync` .</span><span class="sxs-lookup"><span data-stu-id="ef38f-120">However, given that this article is about minimizing calls of `context.sync`, it is actually more confusing to add an entirely unnecessary final `context.sync`.</span></span> <span data-ttu-id="ef38f-121">Таким образом, в этой статье мы оставьте его, когда нет несинхронизированных команд в конце `run` .</span><span class="sxs-lookup"><span data-stu-id="ef38f-121">So, in this article, we leave it out when there are no unsynchronized commands at the end of the `run`.</span></span>

```javascript
Word.run(async function (context) {
    let startTime, endTime;
    const docBody = context.document.body;

    // search() returns an array of Ranges.
    const searchResults = docBody.search('the', { matchWholeWord: true });
    context.load(searchResults, 'items');
    await context.sync();

    // Record the system time.
    startTime = performance.now();

    for (var i = 0; i < searchResults.items.length; i++) {
      searchResults.items[i].font.highlightColor = '#FFFF00';

      await context.sync(); // SYNCHRONIZE IN EACH ITERATION
    }
    
    // await context.sync(); // SYNCHRONIZE AFTER THE LOOP

    // Record the system time again then calculate how long the operation took.
    endTime = performance.now();
    console.log("The operation took: " + (endTime - startTime) + " milliseconds.");
  })
}
```

<span data-ttu-id="ef38f-122">Предыдущий код занял 1 полную секунду для завершения в документе с 200 экземплярами "" в Word на Windows.</span><span class="sxs-lookup"><span data-stu-id="ef38f-122">The preceding code took 1 full second to complete in a document with 200 instances of "the" in Word on Windows.</span></span> <span data-ttu-id="ef38f-123">Но когда строка внутри цикла комментируется и одна и та же строка сразу после того, как цикл некомментационный, операция занимает только `await context.sync();` 1/10 секунды.</span><span class="sxs-lookup"><span data-stu-id="ef38f-123">But when the `await context.sync();` line inside the loop is commented out and the same line just after the loop is uncommented, the operation took only a 1/10th of a second.</span></span> <span data-ttu-id="ef38f-124">В Word в Интернете (с edge в качестве браузера) потребовалось 3 полных секунды с синхронизацией внутри цикла и только 6/10ths секунды с синхронизацией после цикла, примерно в пять раз быстрее.</span><span class="sxs-lookup"><span data-stu-id="ef38f-124">In Word on the web (with Edge as the browser), it took 3 full seconds with the synchronization inside the loop and only 6/10ths of a second with the synchronization after the loop, about five times faster.</span></span> <span data-ttu-id="ef38f-125">В документе с 2000 экземплярами "the" потребовалось (в Word в Интернете) 80 секунд с синхронизацией внутри цикла и только 4 секунды с синхронизацией после цикла, примерно в 20 раз быстрее.</span><span class="sxs-lookup"><span data-stu-id="ef38f-125">In a document with 2000 instances of "the", it took (in Word on the web) 80 seconds with the synchronization inside the loop and only 4 seconds with the synchronization after the loop, about 20 times faster.</span></span>

> [!NOTE]
> <span data-ttu-id="ef38f-126">Стоит задаться вопросом, будет ли синхронизация внутри цикла выполняться быстрее, если синхронизация выполняется параллельно, что можно сделать, просто удалив ключевое слово с передней панели `await` `context.sync()` .</span><span class="sxs-lookup"><span data-stu-id="ef38f-126">It's worth asking whether the synchronize-inside-the-loop version would execute faster if the synchronizations ran concurrently, which could be done by simply removing the `await` keyword from the front of the `context.sync()`.</span></span> <span data-ttu-id="ef38f-127">Это приведет к тому, что время выполнения инициирует синхронизацию, а затем сразу же запустит следующую итерацию цикла, не дожидаясь завершения синхронизации.</span><span class="sxs-lookup"><span data-stu-id="ef38f-127">This would cause the runtime to initiate the synchronization and then immediately start the next iteration of the loop without waiting for the synchronization to complete.</span></span> <span data-ttu-id="ef38f-128">Однако это решение не так хорошо, как перемещение из цикла полностью по `context.sync` этим причинам:</span><span class="sxs-lookup"><span data-stu-id="ef38f-128">However, this is not as good a solution as moving the `context.sync` out of the loop entirely for these reasons:</span></span>
>
> - <span data-ttu-id="ef38f-129">Так же, как команды в пакетном задании синхронизации стоят в очереди, сами пакетные задания стоят в очереди в Office, но Office поддерживает не более 50 пакетных заданий в очереди.</span><span class="sxs-lookup"><span data-stu-id="ef38f-129">Just as the commands in a synchronization batch job are queued, the batch jobs themselves are queued in Office, but Office supports no more than 50 batch jobs in the queue.</span></span> <span data-ttu-id="ef38f-130">Еще больше вызывает ошибки.</span><span class="sxs-lookup"><span data-stu-id="ef38f-130">Any more triggers errors.</span></span> <span data-ttu-id="ef38f-131">Таким образом, если в цикле находится более 50 итераций, существует вероятность превышения размера очереди.</span><span class="sxs-lookup"><span data-stu-id="ef38f-131">So, if there are more than 50 iterations in a loop, there is a chance that the queue size is exceeded.</span></span> <span data-ttu-id="ef38f-132">Чем больше количество итераций, тем больше вероятность этого.</span><span class="sxs-lookup"><span data-stu-id="ef38f-132">The greater the number of iterations, the greater the chance of this happening.</span></span> 
> - <span data-ttu-id="ef38f-133">"Одновременно" не означает одновременно.</span><span class="sxs-lookup"><span data-stu-id="ef38f-133">"Concurrently" does not mean simultaneously.</span></span> <span data-ttu-id="ef38f-134">Выполнение нескольких операций синхронизации по-прежнему займет больше времени, чем одно.</span><span class="sxs-lookup"><span data-stu-id="ef38f-134">It would still take longer to execute multiple synchronization operations than to execute one.</span></span>
> - <span data-ttu-id="ef38f-135">Одновечерные операции не гарантируются для выполнения в том же порядке, в котором они были запущены.</span><span class="sxs-lookup"><span data-stu-id="ef38f-135">Concurrent operations are not guaranteed to complete in the same order in which they started.</span></span> <span data-ttu-id="ef38f-136">В предыдущем примере неважно, какой порядок выделяет слово "the", но существуют сценарии, в которых важно, чтобы элементы в коллекции обрабатывались в порядке.</span><span class="sxs-lookup"><span data-stu-id="ef38f-136">In the preceding example, it doesn't matter what order the  word "the" gets highlighted, but there are scenarios where it's important that the items in the collection be processed in order.</span></span>

## <a name="reading-values-from-the-document-with-the-split-loop-pattern"></a><span data-ttu-id="ef38f-137">Чтение значений из документа с шаблоном раздельного цикла</span><span class="sxs-lookup"><span data-stu-id="ef38f-137">Reading values from the document with the split loop pattern</span></span>

<span data-ttu-id="ef38f-138">Избежать s внутри цикла становится все сложнее, когда код должен прочитать свойство элементов коллекции по мере обработки `context.sync` каждого из них. </span><span class="sxs-lookup"><span data-stu-id="ef38f-138">Avoiding `context.sync`s inside a loop becomes more challenging when the code must *read* a property of the collection items as it processes each one.</span></span> <span data-ttu-id="ef38f-139">Предположим, что коду необходимо итерировать все элементы управления контентом в документе Word и войти в текст первого абзаца, связанного с каждым элементом управления.</span><span class="sxs-lookup"><span data-stu-id="ef38f-139">Suppose your code needs to iterate all the content controls in a Word document and log the text of the first paragraph associated with each control.</span></span> <span data-ttu-id="ef38f-140">Ваши инстинкты программирования могут привести к циклу управления, загрузке свойства каждого (первого) абзаца, вызову для заполнения объекта прокси-абзаца текстом из документа, а затем войти в `text` `context.sync` него.</span><span class="sxs-lookup"><span data-stu-id="ef38f-140">Your programming instincts might lead you to loop over the controls, load the `text` property of each (first) paragraph, call `context.sync` to populate the proxy paragraph object with the text from the document, and then log it.</span></span> <span data-ttu-id="ef38f-141">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="ef38f-141">The following is an example.</span></span>

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load('items');
    await context.sync();

    for (let i = 0; i < contentControls.items.length; i++) {
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst();
      paragraph.load('text');
      await context.sync();
      console.log(paragraph.text);
    }
});
```

<span data-ttu-id="ef38f-142">В этом сценарии, чтобы избежать использования цикла, следует использовать шаблон, который мы называем `context.sync` **шаблоном раздельного цикла.**</span><span class="sxs-lookup"><span data-stu-id="ef38f-142">In this scenario, to avoid having a `context.sync` in a loop, you should use a pattern we call the **split loop** pattern.</span></span> <span data-ttu-id="ef38f-143">Давайте рассмотрим конкретный пример шаблона, прежде чем мы примем официальное описание этого шаблона.</span><span class="sxs-lookup"><span data-stu-id="ef38f-143">Let's see a concrete example of the pattern before we get to a formal description of it.</span></span> <span data-ttu-id="ef38f-144">Вот как шаблон раздельного цикла можно применить к предыдущему фрагменту кода.</span><span class="sxs-lookup"><span data-stu-id="ef38f-144">Here's how the split loop pattern can be applied to the preceding code snippet.</span></span> <span data-ttu-id="ef38f-145">Обратите внимание на следующие аспекты этого кода.</span><span class="sxs-lookup"><span data-stu-id="ef38f-145">Note the following about this code.</span></span>

- <span data-ttu-id="ef38f-146">В настоящее время существует два цикла, и между ними происходит, поэтому внутри них нет `context.sync` `context.sync` ни одной петли.</span><span class="sxs-lookup"><span data-stu-id="ef38f-146">There are now two loops and the `context.sync` comes between them, so there's no `context.sync` inside either loop.</span></span>
- <span data-ttu-id="ef38f-147">Первый цикл итерирует элементы в объекте коллекции и загружает свойство так же, как и исходный цикл, но первый цикл не может войти в текст абзаца, так как он больше не содержит свойства `text` `context.sync` прокси-объекта. `text` `paragraph`</span><span class="sxs-lookup"><span data-stu-id="ef38f-147">The first loop iterates through the items in the collection object and loads the `text` property just as the original loop did, but the first loop cannot log the paragraph text because it no longer contains a `context.sync` to populate the `text` property of the `paragraph` proxy object.</span></span> <span data-ttu-id="ef38f-148">Вместо этого объект `paragraph` добавляется в массив.</span><span class="sxs-lookup"><span data-stu-id="ef38f-148">Instead, it adds the `paragraph` object to an array.</span></span>
- <span data-ttu-id="ef38f-149">Второй цикл итерирует массив, созданный первым циклом, и регистрирует `text` каждый `paragraph` элемент.</span><span class="sxs-lookup"><span data-stu-id="ef38f-149">The second loop iterates through the array that was created by the first loop, and logs the `text` of each `paragraph` item.</span></span> <span data-ttu-id="ef38f-150">Это возможно, так как между двумя циклами заполняются `context.sync` все `text` свойства.</span><span class="sxs-lookup"><span data-stu-id="ef38f-150">This is possible because the `context.sync` that came between the two loops populated all the `text` properties.</span></span>

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load("items");
    await context.sync();

    const firstParagraphsOfCCs = [];
    for (let i = 0; i < contentControls.items.length; i++) {
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst();
      paragraph.load('text');
      firstParagraphsOfCCs.push(paragraph);
    }

    await context.sync();

    for (let i = 0; i < firstParagraphsOfCCs.length; i++) {
      console.log(firstParagraphsOfCCs[i].text);
    }
});
```

<span data-ttu-id="ef38f-151">В предыдущем примере предлагается следующая процедура превращения цикла, содержающего шаблон `context.sync` раздельного цикла.</span><span class="sxs-lookup"><span data-stu-id="ef38f-151">The preceding example suggests the following procedure for turning a loop that contains a `context.sync` into the split loop pattern.</span></span>

1. <span data-ttu-id="ef38f-152">Замените цикл двумя циклами.</span><span class="sxs-lookup"><span data-stu-id="ef38f-152">Replace the loop with two loops.</span></span>
2. <span data-ttu-id="ef38f-153">Создайте первый цикл итерации над коллекцией и добавьте каждый элемент в массив, а также загрузите любое свойство элемента, который необходимо прочитать коду.</span><span class="sxs-lookup"><span data-stu-id="ef38f-153">Create a first loop to iterate over the collection and add each item to an array while also loading any property of the item that your code needs to read.</span></span>
3. <span data-ttu-id="ef38f-154">После первого цикла необходимо вызвать заполнение объектов прокси с `context.sync` любыми загруженными свойствами.</span><span class="sxs-lookup"><span data-stu-id="ef38f-154">Following the first loop, call `context.sync` to populate the proxy objects with any loaded properties.</span></span>
4. <span data-ttu-id="ef38f-155">Выполните второй цикл, чтобы итерировать массив, созданный в первом цикле, и прочитать `context.sync` загруженные свойства.</span><span class="sxs-lookup"><span data-stu-id="ef38f-155">Follow the `context.sync` with a second loop to iterate over the array created in the first loop and read the loaded properties.</span></span>

## <a name="processing-objects-in-the-document-with-the-correlated-objects-pattern"></a><span data-ttu-id="ef38f-156">Обработка объектов в документе с сопоставленным шаблоном объектов</span><span class="sxs-lookup"><span data-stu-id="ef38f-156">Processing objects in the document with the correlated objects pattern</span></span>

<span data-ttu-id="ef38f-157">Рассмотрим более сложный сценарий, при котором обработка элементов в коллекции требует данных, которых нет в самих элементов.</span><span class="sxs-lookup"><span data-stu-id="ef38f-157">Let's consider a more complex scenario where processing the items in the collection requires data that isn't in the items themselves.</span></span> <span data-ttu-id="ef38f-158">Сценарий предполагает надстройки Word, которая работает на документах, созданных из шаблона с некоторым шаблонным текстом.</span><span class="sxs-lookup"><span data-stu-id="ef38f-158">The scenario envisions a Word add-in that operates on documents created from a template with some boilerplate text.</span></span> <span data-ttu-id="ef38f-159">В тексте рассеян один или несколько экземпляров строки следующих задатков: "{Coordinator}", "{Deputy}" и "{Manager}".</span><span class="sxs-lookup"><span data-stu-id="ef38f-159">Scattered in the text are one or more instances of the following placeholder strings: "{Coordinator}", "{Deputy}", and "{Manager}".</span></span> <span data-ttu-id="ef38f-160">Надстройка заменяет каждого замещаемого лица именем определенного человека.</span><span class="sxs-lookup"><span data-stu-id="ef38f-160">The add-in replaces each placeholder with some person's name.</span></span> <span data-ttu-id="ef38f-161">Пользовательский интерфейс надстройки не важен для этой статьи.</span><span class="sxs-lookup"><span data-stu-id="ef38f-161">The UI of the add-in is not important to this article.</span></span> <span data-ttu-id="ef38f-162">Например, в нем может быть области задач с тремя текстовыми полями, каждая из которых помечена одним из держателей.</span><span class="sxs-lookup"><span data-stu-id="ef38f-162">For example, it could have a task pane with three text boxes, each labeled with one of the placeholders.</span></span> <span data-ttu-id="ef38f-163">Пользователь вводит имя в каждом текстовом окне, а затем нажмет кнопку **Заменить.**</span><span class="sxs-lookup"><span data-stu-id="ef38f-163">The user enters a name in each text box and then presses a **Replace** button.</span></span> <span data-ttu-id="ef38f-164">Обработчик кнопки создает массив, который совмещает имена с держателями, а затем заменяет каждого держателя на назначенное имя.</span><span class="sxs-lookup"><span data-stu-id="ef38f-164">The handler for the button creates an array that maps the names to the placeholders, and then replaces each placeholder with the assigned name.</span></span> 

<span data-ttu-id="ef38f-165">Чтобы поэкспериментировать с кодом, не нужно создавать надстройку с этим пользовательским интерфейсом.</span><span class="sxs-lookup"><span data-stu-id="ef38f-165">You don't need to actually produce an add-in with this UI to experiment with the code.</span></span> <span data-ttu-id="ef38f-166">Вы можете использовать Script Lab [для](../overview/explore-with-script-lab.md) прототипа важного кода.</span><span class="sxs-lookup"><span data-stu-id="ef38f-166">You can use the [Script Lab tool](../overview/explore-with-script-lab.md) to prototype the important code.</span></span> <span data-ttu-id="ef38f-167">Чтобы создать массив сопоставления, используйте следующее утверждение назначения.</span><span class="sxs-lookup"><span data-stu-id="ef38f-167">Use the following assignment statement to create the mapping array.</span></span>

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

<span data-ttu-id="ef38f-168">В следующем коде показано, как можно заменить каждого держателя на назначенное имя, если вы использовали `context.sync` внутренние циклы.</span><span class="sxs-lookup"><span data-stu-id="ef38f-168">The following code shows how you might replace each placeholder with its assigned name if you used `context.sync` inside loops.</span></span>

```javascript
Word.run(async (context) => {

    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildCards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');

      await context.sync(); 

      for (let j = 0; j < searchResults.items.length; j++) {
        searchResults.items[j].insertText(jobMapping[i].person, Word.InsertLocation.replace);

        await context.sync();
      }
    }
});
```

<span data-ttu-id="ef38f-169">В предыдущем коде есть внешний и внутренний цикл.</span><span class="sxs-lookup"><span data-stu-id="ef38f-169">In the preceding code, there is an outer and an inner loop.</span></span> <span data-ttu-id="ef38f-170">Каждый из них содержит `context.sync` .</span><span class="sxs-lookup"><span data-stu-id="ef38f-170">Each of them contains a `context.sync`.</span></span> <span data-ttu-id="ef38f-171">На основе самого первого фрагмента кода в этой статье вы, вероятно, увидите, что внутренний цикл можно просто `context.sync` перемещать после внутреннего цикла.</span><span class="sxs-lookup"><span data-stu-id="ef38f-171">Based on the very first code snippet in this article, you probably see that the `context.sync` in the inner loop can simply be moved after the inner loop.</span></span> <span data-ttu-id="ef38f-172">Но это по-прежнему оставляет код с (два из них на `context.sync` самом деле) во внешнем цикле.</span><span class="sxs-lookup"><span data-stu-id="ef38f-172">But that would still leave the code with a `context.sync` (two of them actually) in the outer loop.</span></span> <span data-ttu-id="ef38f-173">В следующем коде показано, как можно удалить `context.sync` из циклов.</span><span class="sxs-lookup"><span data-stu-id="ef38f-173">The following code shows how you can remove `context.sync` from the loops.</span></span> <span data-ttu-id="ef38f-174">Ниже мы обсудим код.</span><span class="sxs-lookup"><span data-stu-id="ef38f-174">We discuss the code below.</span></span>

```javascript
Word.run(async (context) => {

    const allSearchResults = [];
    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildCards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');
      let correlatedSearchResult = {
        rangesMatchingJob: searchResults,
        personAssignedToJob: jobMapping[i].person
      }
      allSearchResults.push(correlatedSearchResult);
    }

    await context.sync()

    for (let i = 0; i < allSearchResults.length; i++) {
      let correlatedObject = allSearchResults[i];

      for (let j = 0; j < correlatedObject.rangesMatchingJob.items.length; j++) {
        let targetRange = correlatedObject.rangesMatchingJob.items[j];
        let name = correlatedObject.personAssignedToJob;
        targetRange.insertText(name, Word.InsertLocation.replace);
      }
    }

    await context.sync();
});
```

<span data-ttu-id="ef38f-175">Обратите внимание, что в коде используется шаблон раздельного цикла:</span><span class="sxs-lookup"><span data-stu-id="ef38f-175">Note the code uses the split loop pattern:</span></span>

- <span data-ttu-id="ef38f-176">Внешний цикл из предыдущего примера был разделен на два.</span><span class="sxs-lookup"><span data-stu-id="ef38f-176">The outer loop from the preceding example has been split into two.</span></span> <span data-ttu-id="ef38f-177">(Второй цикл имеет внутренний цикл, который ожидается, так как код итерирует над набором заданий (или держателей) и в этом наборе он итерирует над диапазонами совпадений.)</span><span class="sxs-lookup"><span data-stu-id="ef38f-177">(The second loop has an inner loop, which is expected because the code is iterating over a set of jobs (or placeholders) and within that set it is iterating over the matching ranges.)</span></span>
- <span data-ttu-id="ef38f-178">Существует после `context.sync` каждого крупного цикла, но не внутри `context.sync` цикла.</span><span class="sxs-lookup"><span data-stu-id="ef38f-178">There is a `context.sync` after each major loop, but no `context.sync` inside any loop.</span></span>
- <span data-ttu-id="ef38f-179">Второй основной цикл итерирует массив, созданный в первом цикле.</span><span class="sxs-lookup"><span data-stu-id="ef38f-179">The second major loop iterates through an array that is created in the first loop.</span></span>

<span data-ttu-id="ef38f-180">Но массив, созданный в  первом цикле, не содержит только объект Office, как это сделал первый цикл в разделе Чтение значений из документа с шаблоном [раздельного цикла](#reading-values-from-the-document-with-the-split-loop-pattern).</span><span class="sxs-lookup"><span data-stu-id="ef38f-180">But the array created in the first loop does *not* contain only an Office object as the first loop did in the section [Reading values from the document with the split loop pattern](#reading-values-from-the-document-with-the-split-loop-pattern).</span></span> <span data-ttu-id="ef38f-181">Это происходит из-за того, что некоторые сведения, необходимые для обработки объектов диапазона Word, не находятся в самих объектах Range, а приходят из `jobMapping` массива.</span><span class="sxs-lookup"><span data-stu-id="ef38f-181">This is because some of the information needed to process the Word Range objects is not in the Range objects themselves but instead comes from the `jobMapping` array.</span></span>

<span data-ttu-id="ef38f-182">Таким образом, объекты массива, созданные в первом цикле, — это настраиваемые объекты, которые имеют два свойства.</span><span class="sxs-lookup"><span data-stu-id="ef38f-182">So, the objects in the array created in the first loop are custom objects that have two properties.</span></span> <span data-ttu-id="ef38f-183">Первый — это массив диапазонов Word, которые соответствуют определенному названию задания (то есть строке задатки), а второй — строке, которая предоставляет имя человека, назначенного на задание.</span><span class="sxs-lookup"><span data-stu-id="ef38f-183">The first is an array of Word Ranges that match a specific job title (that is, a placeholder string) and the second is a string that provides the name of the person assigned to the job.</span></span> <span data-ttu-id="ef38f-184">Это упрощает написание окончательного цикла и упрощает чтение, так как вся информация, необходимая для обработки данного диапазона, содержится в том же настраиваемом объекте, который содержит диапазон.</span><span class="sxs-lookup"><span data-stu-id="ef38f-184">This makes the final loop easy to write and easy to read because all of the information needed to process a given range is contained in the same custom object that contains the range.</span></span> <span data-ttu-id="ef38f-185">Имя, которое должно заменить _**correlatedObject**.rangesMatchingJob.items[j],_ является другим свойством того же объекта: _**correlatedObject**.personAssignedToJob_.</span><span class="sxs-lookup"><span data-stu-id="ef38f-185">The name that should replace _**correlatedObject**.rangesMatchingJob.items[j]_ is the other property of the same object: _**correlatedObject**.personAssignedToJob_.</span></span>

<span data-ttu-id="ef38f-186">Этот вариант шаблона раздельного цикла мы называем **шаблоном коррелирующих объектов.**</span><span class="sxs-lookup"><span data-stu-id="ef38f-186">We call this variation of the split loop pattern the **correlated objects** pattern.</span></span> <span data-ttu-id="ef38f-187">Общая идея состоит в том, что первый цикл создает массив настраиваемого объекта.</span><span class="sxs-lookup"><span data-stu-id="ef38f-187">The general idea is that the first loop creates an array of custom objects.</span></span> <span data-ttu-id="ef38f-188">Каждый объект имеет свойство, значение которого является одним из элементов в объекте Office коллекции (или массиве таких элементов).</span><span class="sxs-lookup"><span data-stu-id="ef38f-188">Each object has a property whose value is one of the items in an Office collection object (or an array of such items).</span></span> <span data-ttu-id="ef38f-189">Настраиваемый объект имеет другие свойства, каждое из которых предоставляет сведения, необходимые для обработки Office объектов в финальном цикле.</span><span class="sxs-lookup"><span data-stu-id="ef38f-189">The custom object has other properties, each of which provides information needed to process the Office objects in the final loop.</span></span> <span data-ttu-id="ef38f-190">См. в разделе [Другие примеры](#other-examples-of-these-patterns) этих шаблонов для ссылки на пример, в котором настраиваемый объект соотносится более чем с двумя свойствами.</span><span class="sxs-lookup"><span data-stu-id="ef38f-190">See the section [Other examples of these patterns](#other-examples-of-these-patterns) for a link to an example where the custom correlating object has more than two properties.</span></span>

<span data-ttu-id="ef38f-191">Еще один нюанс: иногда для создания массива настраиваемого коррелирующих объектов требуется несколько циклов.</span><span class="sxs-lookup"><span data-stu-id="ef38f-191">One further caveat: sometimes it takes more than one loop just to create the array of custom correlating objects.</span></span> <span data-ttu-id="ef38f-192">Это может произойти, если необходимо прочитать свойство каждого члена одного объекта Office только для сбора сведений, которые будут использоваться для обработки другого объекта коллекции.</span><span class="sxs-lookup"><span data-stu-id="ef38f-192">This can happen if you need to read a property of each member of one Office collection object just to gather information that will be used to process another collection object.</span></span> <span data-ttu-id="ef38f-193">(Например, код должен читать заголовки всех столбцов в таблице Excel, так как надстройка будет применять формат номеров к ячейкам некоторых столбцов на основе заголовка этого столбца.) Но вы всегда можете сохранить `context.sync` s между циклами, а не в цикле.</span><span class="sxs-lookup"><span data-stu-id="ef38f-193">(For example, your code needs to read the titles of all the columns in an Excel table because your add-in is going to apply a number format to the cells of some columns based on that column's title.) But you can always keep the `context.sync`s between the loops, rather than in a loop.</span></span> <span data-ttu-id="ef38f-194">В примере [см. раздел](#other-examples-of-these-patterns) Другие примеры этих шаблонов.</span><span class="sxs-lookup"><span data-stu-id="ef38f-194">See the section [Other examples of these patterns](#other-examples-of-these-patterns) for an example.</span></span>

## <a name="other-examples-of-these-patterns"></a><span data-ttu-id="ef38f-195">Другие примеры этих шаблонов</span><span class="sxs-lookup"><span data-stu-id="ef38f-195">Other examples of these patterns</span></span>

- <span data-ttu-id="ef38f-196">В очень простом примере для Excel, использующих циклы, см. принятый ответ на этот вопрос о переполнении стека: возможно ли стоять в очереди несколько `Array.forEach` [context.load перед context.sync?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)</span><span class="sxs-lookup"><span data-stu-id="ef38f-196">For a very simple example for Excel that uses `Array.forEach` loops, see the accepted answer to this Stack Overflow question: [Is it possible to queue more than one context.load before context.sync?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)</span></span>
- <span data-ttu-id="ef38f-197">Простой пример для Word, которое использует циклы и не использует синтаксис, см. в принятом ответе на этот вопрос Переполнение стека: Итерирование всех абзацев с управлением контентом с Office `Array.forEach` `async` / `await` [API JavaScript](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).</span><span class="sxs-lookup"><span data-stu-id="ef38f-197">For a simple example for Word that uses `Array.forEach` loops and doesn't use `async`/`await` syntax, see the accepted answer to this Stack Overflow question: [Iterating over all paragraphs with content controls with Office JavaScript API](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).</span></span>
- <span data-ttu-id="ef38f-198">Пример Word, который написан в TypeScript, см. в примере Word [Add-in Angular2 Style Checker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), особенно файлword.doc[ ument.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts).</span><span class="sxs-lookup"><span data-stu-id="ef38f-198">For an example for Word that is written in TypeScript, see the sample [Word Add-in Angular2 Style Checker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), especially the file [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts).</span></span> <span data-ttu-id="ef38f-199">Он имеет смесь `for` и `Array.forEach` циклы.</span><span class="sxs-lookup"><span data-stu-id="ef38f-199">It has a mixture of `for` and `Array.forEach` loops.</span></span>
- <span data-ttu-id="ef38f-200">Для продвинутого примера Word [импортировать этот gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) в средство [Script Lab.](../overview/explore-with-script-lab.md)</span><span class="sxs-lookup"><span data-stu-id="ef38f-200">For an advanced Word sample, import [this gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) into the [Script Lab tool](../overview/explore-with-script-lab.md).</span></span> <span data-ttu-id="ef38f-201">Для контекста с помощью gist см. принятый ответ на вопрос Stack Overflow Document, который не синхронизируется [после замены текста.](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text)</span><span class="sxs-lookup"><span data-stu-id="ef38f-201">For context in using the gist, see the accepted answer to the Stack Overflow question [Document not in sync after replace text](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text).</span></span> <span data-ttu-id="ef38f-202">В этом примере создается настраиваемый тип объекта, который имеет три свойства.</span><span class="sxs-lookup"><span data-stu-id="ef38f-202">This sample creates a custom correlating object type that has three properties.</span></span> <span data-ttu-id="ef38f-203">Для построения массива коррелирующих объектов используется в общей сложности три цикла и еще два цикла для окончательной обработки.</span><span class="sxs-lookup"><span data-stu-id="ef38f-203">It uses a total of three loops to construct the array of correlated objects, and two more loops to do the final processing.</span></span> <span data-ttu-id="ef38f-204">Существует сочетание циклов `for` `Array.forEach` и циклов.</span><span class="sxs-lookup"><span data-stu-id="ef38f-204">There are a mixture of `for` and `Array.forEach` loops.</span></span>
- <span data-ttu-id="ef38f-205">Хотя не является строго примером шаблонов раздельного цикла или соотносимых объектов, существует расширенный пример Excel, который показывает, как преобразовать набор значений ячейки в другие валюты с помощью одного . `context.sync`</span><span class="sxs-lookup"><span data-stu-id="ef38f-205">Although not strictly an example of the split loop or correlated objects patterns, there is an advanced Excel sample that shows how to convert a set of cell values to other currencies with just a single `context.sync`.</span></span> <span data-ttu-id="ef38f-206">Чтобы попробовать его, откройте [Script Lab и](../overview/explore-with-script-lab.md) перейдите к примеру **конвертера** валюты.</span><span class="sxs-lookup"><span data-stu-id="ef38f-206">To try it, open the [Script Lab tool](../overview/explore-with-script-lab.md) and navigate to the **Currency Converter** sample.</span></span>

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a><span data-ttu-id="ef38f-207">Когда не *следует использовать* шаблоны в этой статье?</span><span class="sxs-lookup"><span data-stu-id="ef38f-207">When should you *not* use the patterns in this article?</span></span>

<span data-ttu-id="ef38f-208">Excel не может прочитать более 5 МБ данных в заданный вызов `context.sync` .</span><span class="sxs-lookup"><span data-stu-id="ef38f-208">Excel cannot read more than 5 MB of data in a given call of `context.sync`.</span></span> <span data-ttu-id="ef38f-209">Если это ограничение превышено, будет выброшена ошибка.</span><span class="sxs-lookup"><span data-stu-id="ef38f-209">If this limit is exceeded, an error is thrown.</span></span> <span data-ttu-id="ef38f-210">(Дополнительные сведения см. в разделе Excel [](resource-limits-and-performance-optimization.md#excel-add-ins) надстройки" ограничений ресурсов и оптимизации производительности для Office надстройки.) Очень редко к этому ограничению подходят, но если есть вероятность, что это произойдет с вашей  надстройкой, то код не должен загружать все данные в один цикл и следовать циклу с `context.sync` помощью .</span><span class="sxs-lookup"><span data-stu-id="ef38f-210">(See the "Excel add-ins section" of [Resource limits and performance optimization for Office Add-ins](resource-limits-and-performance-optimization.md#excel-add-ins) for more information.) It's very rare that this limit is approached, but if there's a chance that this will happen with your add-in, then your code should *not* load all the data in a single loop and follow the loop with a `context.sync`.</span></span> <span data-ttu-id="ef38f-211">Но все равно следует избегать каждой итерации цикла над `context.sync` объектом коллекции.</span><span class="sxs-lookup"><span data-stu-id="ef38f-211">But you still should avoid having a `context.sync` in every iteration of a loop over a collection object.</span></span> <span data-ttu-id="ef38f-212">Вместо этого определите подсети элементов в коллекции и петлю по каждому подмножество в свою очередь, с между `context.sync` циклами.</span><span class="sxs-lookup"><span data-stu-id="ef38f-212">Instead, define subsets of the items in the collection and loop over each subset in turn, with a `context.sync` between the loops.</span></span> <span data-ttu-id="ef38f-213">Вы можете структурировать это с помощью внешнего цикла, который итерирует над подсетями и содержит в каждой из этих `context.sync` внешних итераций.</span><span class="sxs-lookup"><span data-stu-id="ef38f-213">You could structure this with an outer loop that iterates over the subsets and contains the `context.sync` in each of these outer iterations.</span></span>
