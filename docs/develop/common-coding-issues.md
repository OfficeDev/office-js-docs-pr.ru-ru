---
title: Распространенные проблемы кодирования и неожиданное поведение платформы
description: Список проблем платформы API JavaScript для Office, часто встречающихся разработчиками.
ms.date: 10/29/2019
localization_priority: Normal
ms.openlocfilehash: 8cea95e3214585ba8e0b77535916f9c564dde9df
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902200"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="016bc-103">Распространенные проблемы кодирования и неожиданное поведение платформы</span><span class="sxs-lookup"><span data-stu-id="016bc-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="016bc-104">В этой статье описываются аспекты API JavaScript для Office, которые могут привести к непредвиденному поведению или требуют определенных шаблонов кодирования для достижения желаемого результата.</span><span class="sxs-lookup"><span data-stu-id="016bc-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="016bc-105">Если возникла проблема, связанная с этим списком, сообщите нам об этом с помощью формы отзыва в нижней части статьи.</span><span class="sxs-lookup"><span data-stu-id="016bc-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="some-properties-must-be-set-with-json-structs"></a><span data-ttu-id="016bc-106">Некоторые свойства должны быть заданы с помощью структуры JSON</span><span class="sxs-lookup"><span data-stu-id="016bc-106">Some properties must be set with JSON structs</span></span>

> [!NOTE]
> <span data-ttu-id="016bc-107">Этот раздел относится только к API, предназначенным для ведущего приложения, для Excel и Word.</span><span class="sxs-lookup"><span data-stu-id="016bc-107">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="016bc-108">Некоторые свойства должны быть заданы как структуры JSON, а не как задавать отдельные вложенные свойства.</span><span class="sxs-lookup"><span data-stu-id="016bc-108">Some properties must be set as JSON structs, instead of setting their individual subproperties.</span></span> <span data-ttu-id="016bc-109">Один из примеров этого примера находится в файле [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="016bc-109">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="016bc-110">Свойство должно быть задано с помощью одного объекта Пажелайаутзумоптионс, как показано ниже: [](/javascript/api/excel/excel.pagelayoutzoomoptions) `zoom`</span><span class="sxs-lookup"><span data-stu-id="016bc-110">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="016bc-111">В предыдущем примере вы ***не*** сможете напрямую присвоить `zoom` значение: `sheet.pageLayout.zoom.scale = 200;`.</span><span class="sxs-lookup"><span data-stu-id="016bc-111">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="016bc-112">Этот оператор выдает ошибку, `zoom` так как не загружен.</span><span class="sxs-lookup"><span data-stu-id="016bc-112">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="016bc-113">Даже если `zoom` были загружены, набор масштабов не вступит в силу.</span><span class="sxs-lookup"><span data-stu-id="016bc-113">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="016bc-114">Все операции контекста выполняются `zoom`, обновляя прокси-объект в надстройке и перезаписывая локально заданные значения.</span><span class="sxs-lookup"><span data-stu-id="016bc-114">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="016bc-115">Это поведение отличается от [свойств навигации](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) , таких как [Range. Format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="016bc-115">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="016bc-116">Свойства `format` можно задать с помощью навигации по объектам, как показано ниже:</span><span class="sxs-lookup"><span data-stu-id="016bc-116">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="016bc-117">Можно определить свойство, для которого должны быть заданы вложенные свойства структуры JSON, путем проверки модификатора "только чтение".</span><span class="sxs-lookup"><span data-stu-id="016bc-117">You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier.</span></span> <span data-ttu-id="016bc-118">Все свойства, доступные только для чтения, могут иметь нередактируемые вложенные свойства, не предназначенные только для чтения.</span><span class="sxs-lookup"><span data-stu-id="016bc-118">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="016bc-119">Записываемые свойства, `PageLayout.zoom` такие как, должны быть заданы с помощью структуры JSON.</span><span class="sxs-lookup"><span data-stu-id="016bc-119">Writeable properties like `PageLayout.zoom` must be set with a JSON struct.</span></span> <span data-ttu-id="016bc-120">В сводке:</span><span class="sxs-lookup"><span data-stu-id="016bc-120">In summary:</span></span>

- <span data-ttu-id="016bc-121">Свойство только для чтения: вложенные свойства можно задать с помощью навигации.</span><span class="sxs-lookup"><span data-stu-id="016bc-121">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="016bc-122">Записываемое свойство: вложенные свойства должны быть заданы с помощью структуры JSON (и не могут быть заданы с помощью навигации).</span><span class="sxs-lookup"><span data-stu-id="016bc-122">Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="016bc-123">Установка свойств, предназначенных только для чтения</span><span class="sxs-lookup"><span data-stu-id="016bc-123">Setting read-only properties</span></span>

<span data-ttu-id="016bc-124">[Определения TypeScript](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) для Office JS указывают, какие свойства объекта доступны только для чтения.</span><span class="sxs-lookup"><span data-stu-id="016bc-124">The [TypeScript definitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="016bc-125">Если вы попытаетесь установить свойство, доступное только для чтения, операция записи завершится с ошибкой без уведомления и не выдается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="016bc-125">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="016bc-126">В следующем примере ошибочно попытаются задать свойство, доступное только для чтения, [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="016bc-126">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a><span data-ttu-id="016bc-127">См. также</span><span class="sxs-lookup"><span data-stu-id="016bc-127">See also</span></span>

- <span data-ttu-id="016bc-128">[OfficeDev/Office-JS](https://github.com/OfficeDev/office-js/issues): место для создания отчетов и просмотра проблем с платформой надстроек Office и API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="016bc-128">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="016bc-129">[Переполнение стека](https://stackoverflow.com/questions/tagged/office-js): место для Ask и просмотра вопросов по программированию, посвященных API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="016bc-129">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="016bc-130">При публикации в стеке обязательно примените к вопросу тег "Office — JS".</span><span class="sxs-lookup"><span data-stu-id="016bc-130">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="016bc-131">[UserVoice](https://officespdev.uservoice.com/): в этом месте вы можете предложить новые функции для платформы надстроек Office и API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="016bc-131">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
