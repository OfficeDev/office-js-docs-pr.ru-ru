---
title: Распространенные проблемы кодирования и неожиданное поведение платформы
description: Список проблем платформы API JavaScript для Office, часто встречающихся разработчиками.
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: a4d7a09c1645bea181060157d933036d1924044f
ms.sourcegitcommit: 88d81aa2d707105cf0eb55d9774b2e7cf468b03a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/13/2019
ms.locfileid: "38301934"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="b27c6-103">Распространенные проблемы кодирования и неожиданное поведение платформы</span><span class="sxs-lookup"><span data-stu-id="b27c6-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="b27c6-104">В этой статье описываются аспекты API JavaScript для Office, которые могут привести к непредвиденному поведению или требуют определенных шаблонов кодирования для достижения желаемого результата.</span><span class="sxs-lookup"><span data-stu-id="b27c6-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="b27c6-105">Если возникла проблема, связанная с этим списком, сообщите нам об этом с помощью формы отзыва в нижней части статьи.</span><span class="sxs-lookup"><span data-stu-id="b27c6-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="b27c6-106">Общие API и API Outlook не основаны на обещаниях</span><span class="sxs-lookup"><span data-stu-id="b27c6-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="b27c6-107">[Общие API](/javascript/api/office) (которые не привязаны к определенному ведущему приложению Office) и [API Outlook](/javascript/api/outlook) используют модель программирования на основе обратных вызовов.</span><span class="sxs-lookup"><span data-stu-id="b27c6-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="b27c6-108">Для взаимодействия с базовым документом Office требуется асинхронный вызов чтения или записи, указывающий обратный вызов, который должен выполняться при завершении операции.</span><span class="sxs-lookup"><span data-stu-id="b27c6-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes.</span></span> <span data-ttu-id="b27c6-109">Пример этого шаблона приведен в статье [Document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="b27c6-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="b27c6-110">Эти общие API и методы API Outlook не возвращают [обещаний](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span><span class="sxs-lookup"><span data-stu-id="b27c6-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="b27c6-111">Таким образом, вы не можете использовать параметр [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) , чтобы приостановить выполнение до завершения асинхронной операции.</span><span class="sxs-lookup"><span data-stu-id="b27c6-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="b27c6-112">Если требуется `await` поведение, можно создать оболочку вызова метода в явно созданном обещании.</span><span class="sxs-lookup"><span data-stu-id="b27c6-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

```js
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

> [!NOTE]
> <span data-ttu-id="b27c6-113">Справочная документация содержит реализацию [файла. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)в оболочке для обещания.</span><span class="sxs-lookup"><span data-stu-id="b27c6-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-must-be-set-with-json-structs"></a><span data-ttu-id="b27c6-114">Некоторые свойства должны быть заданы с помощью структуры JSON</span><span class="sxs-lookup"><span data-stu-id="b27c6-114">Some properties must be set with JSON structs</span></span>

> [!NOTE]
> <span data-ttu-id="b27c6-115">Этот раздел относится только к API, предназначенным для ведущего приложения, для Excel и Word.</span><span class="sxs-lookup"><span data-stu-id="b27c6-115">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="b27c6-116">Некоторые свойства должны быть заданы как структуры JSON, а не как задавать отдельные вложенные свойства.</span><span class="sxs-lookup"><span data-stu-id="b27c6-116">Some properties must be set as JSON structs, instead of setting their individual subproperties.</span></span> <span data-ttu-id="b27c6-117">Один из примеров этого примера находится в файле [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="b27c6-117">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="b27c6-118">Свойство должно быть задано с помощью одного объекта Пажелайаутзумоптионс, как показано ниже: [](/javascript/api/excel/excel.pagelayoutzoomoptions) `zoom`</span><span class="sxs-lookup"><span data-stu-id="b27c6-118">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="b27c6-119">В предыдущем примере вы ***не*** сможете напрямую присвоить `zoom` значение: `sheet.pageLayout.zoom.scale = 200;`.</span><span class="sxs-lookup"><span data-stu-id="b27c6-119">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="b27c6-120">Этот оператор выдает ошибку, `zoom` так как не загружен.</span><span class="sxs-lookup"><span data-stu-id="b27c6-120">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="b27c6-121">Даже если `zoom` были загружены, набор масштабов не вступит в силу.</span><span class="sxs-lookup"><span data-stu-id="b27c6-121">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="b27c6-122">Все операции контекста выполняются `zoom`, обновляя прокси-объект в надстройке и перезаписывая локально заданные значения.</span><span class="sxs-lookup"><span data-stu-id="b27c6-122">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="b27c6-123">Это поведение отличается от [свойств навигации](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) , таких как [Range. Format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="b27c6-123">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="b27c6-124">Свойства `format` можно задать с помощью навигации по объектам, как показано ниже:</span><span class="sxs-lookup"><span data-stu-id="b27c6-124">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="b27c6-125">Можно определить свойство, для которого должны быть заданы вложенные свойства структуры JSON, путем проверки модификатора "только чтение".</span><span class="sxs-lookup"><span data-stu-id="b27c6-125">You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier.</span></span> <span data-ttu-id="b27c6-126">Все свойства, доступные только для чтения, могут иметь нередактируемые вложенные свойства, не предназначенные только для чтения.</span><span class="sxs-lookup"><span data-stu-id="b27c6-126">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="b27c6-127">Записываемые свойства, `PageLayout.zoom` такие как, должны быть заданы с помощью структуры JSON.</span><span class="sxs-lookup"><span data-stu-id="b27c6-127">Writeable properties like `PageLayout.zoom` must be set with a JSON struct.</span></span> <span data-ttu-id="b27c6-128">В сводке:</span><span class="sxs-lookup"><span data-stu-id="b27c6-128">In summary:</span></span>

- <span data-ttu-id="b27c6-129">Свойство только для чтения: вложенные свойства можно задать с помощью навигации.</span><span class="sxs-lookup"><span data-stu-id="b27c6-129">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="b27c6-130">Записываемое свойство: вложенные свойства должны быть заданы с помощью структуры JSON (и не могут быть заданы с помощью навигации).</span><span class="sxs-lookup"><span data-stu-id="b27c6-130">Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).</span></span>

## <a name="excel-data-transfer-limits"></a><span data-ttu-id="b27c6-131">Пределы переноса данных Excel</span><span class="sxs-lookup"><span data-stu-id="b27c6-131">Excel data transfer limits</span></span>

<span data-ttu-id="b27c6-132">При создании надстройки Excel учитывайте следующие ограничения размера при взаимодействии с книгой:</span><span class="sxs-lookup"><span data-stu-id="b27c6-132">If you're building an Excel add-in, be aware of the following size limitations when interacting with the workbook:</span></span>

- <span data-ttu-id="b27c6-133">В Excel в Интернете действует ограничение в 5 МБ на размер полезных данных запросов и откликов.</span><span class="sxs-lookup"><span data-stu-id="b27c6-133">Excel on the web has a payload size limit for requests and responses of 5MB.</span></span> <span data-ttu-id="b27c6-134">При превышении этого ограничения возникает ошибка `RichAPI.Error`.</span><span class="sxs-lookup"><span data-stu-id="b27c6-134">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>
- <span data-ttu-id="b27c6-135">Диапазон ограничен 5 000 000 ячейками для операций Get.</span><span class="sxs-lookup"><span data-stu-id="b27c6-135">A range is limited to five million cells for get operations.</span></span>

<span data-ttu-id="b27c6-136">Если ожидается, что вводимые пользователем данные превышают эти ограничения, обязательно проверьте данные перед вызовом `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="b27c6-136">If you expect user input to exceed these limits, be sure to check the data before calling `context.sync()`.</span></span> <span data-ttu-id="b27c6-137">При необходимости разделите операцию на небольшие части.</span><span class="sxs-lookup"><span data-stu-id="b27c6-137">Split the operation into smaller pieces as needed.</span></span> <span data-ttu-id="b27c6-138">Не забудьте позвонить `context.sync()` по каждой подоперации, чтобы избежать повторного пакетной операции.</span><span class="sxs-lookup"><span data-stu-id="b27c6-138">Be sure to call `context.sync()` for each sub-operation to avoid those operations getting batched together again.</span></span>

<span data-ttu-id="b27c6-139">Эти ограничения обычно превышаются с помощью больших диапазонов.</span><span class="sxs-lookup"><span data-stu-id="b27c6-139">These limitations are typically exceeded by large ranges.</span></span> <span data-ttu-id="b27c6-140">Надстройка может использовать [RangeAreas](/javascript/api/excel/excel.rangeareas) для стратегических обновлений ячеек в пределах большого диапазона.</span><span class="sxs-lookup"><span data-stu-id="b27c6-140">Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range.</span></span> <span data-ttu-id="b27c6-141">Для получения дополнительных сведений просмотрите [работу с несколькими диапазонами в](../excel/excel-add-ins-multiple-ranges.md) надстройках Excel.</span><span class="sxs-lookup"><span data-stu-id="b27c6-141">See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="b27c6-142">Установка свойств, предназначенных только для чтения</span><span class="sxs-lookup"><span data-stu-id="b27c6-142">Setting read-only properties</span></span>

<span data-ttu-id="b27c6-143">[Определения TypeScript](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) для Office JS указывают, какие свойства объекта доступны только для чтения.</span><span class="sxs-lookup"><span data-stu-id="b27c6-143">The [TypeScript definitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="b27c6-144">Если вы попытаетесь установить свойство, доступное только для чтения, операция записи завершится с ошибкой без уведомления и не выдается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="b27c6-144">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="b27c6-145">В следующем примере ошибочно попытаются задать свойство, доступное только для чтения, [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="b27c6-145">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a><span data-ttu-id="b27c6-146">См. также</span><span class="sxs-lookup"><span data-stu-id="b27c6-146">See also</span></span>

- <span data-ttu-id="b27c6-147">[OfficeDev/Office-JS](https://github.com/OfficeDev/office-js/issues): место для создания отчетов и просмотра проблем с платформой надстроек Office и API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b27c6-147">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="b27c6-148">[Переполнение стека](https://stackoverflow.com/questions/tagged/office-js): место для Ask и просмотра вопросов по программированию, посвященных API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="b27c6-148">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="b27c6-149">При публикации в стеке обязательно примените к вопросу тег "Office — JS".</span><span class="sxs-lookup"><span data-stu-id="b27c6-149">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="b27c6-150">[UserVoice](https://officespdev.uservoice.com/): в этом месте вы можете предложить новые функции для платформы надстроек Office и API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="b27c6-150">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
