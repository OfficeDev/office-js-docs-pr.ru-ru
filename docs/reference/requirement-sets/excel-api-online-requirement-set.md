---
title: Набор обязательных элементов API JavaScript для Excel Online
description: Сведения о наборе требований Ексцелапионлине.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 16c96f413424d5fc85a21419fb72cf6580c1ac18
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996531"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="5e640-103">Набор обязательных элементов API JavaScript для Excel Online</span><span class="sxs-lookup"><span data-stu-id="5e640-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="5e640-104">`ExcelApiOnline`Набор требований является особым набором требований, включающим функции, доступные только для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="5e640-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="5e640-105">API в этом наборе обязательных элементов считаются рабочими API (не подчиняются недокументированным изменениям поведения или структурным изменениям) для Excel в веб-приложении.</span><span class="sxs-lookup"><span data-stu-id="5e640-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web application.</span></span> <span data-ttu-id="5e640-106">`ExcelApiOnline` считаются API-интерфейсами Preview для других платформ (Windows, Mac, iOS) и могут не поддерживаться ни одной из этих платформ.</span><span class="sxs-lookup"><span data-stu-id="5e640-106">`ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="5e640-107">Если API в `ExcelApiOnline` наборе обязательных элементов поддерживаются на всех платформах, они будут добавлены к следующему набору обязательных требований ( `ExcelApi 1.[NEXT]` ).</span><span class="sxs-lookup"><span data-stu-id="5e640-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="5e640-108">После того как новое требование будет общедоступным, эти API будут удалены из `ExcelApiOnline` .</span><span class="sxs-lookup"><span data-stu-id="5e640-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="5e640-109">Это можно считать похожим процессом повышения роли API, который перемещается из бета-версии в выпуск.</span><span class="sxs-lookup"><span data-stu-id="5e640-109">Think of this as a similar promotion process as an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5e640-110">`ExcelApiOnline` является надмножеством набора последних пронумерованных требований.</span><span class="sxs-lookup"><span data-stu-id="5e640-110">`ExcelApiOnline` is superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5e640-111">`ExcelApiOnline 1.1` — Единственная версия интерфейсов API, предназначенных только для интерактивного подключения.</span><span class="sxs-lookup"><span data-stu-id="5e640-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="5e640-112">Это связано с тем, что Excel в Интернете всегда будет иметь одну версию, доступную для пользователей с последней версией.</span><span class="sxs-lookup"><span data-stu-id="5e640-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

## <a name="recommended-usage"></a><span data-ttu-id="5e640-113">Рекомендуемое использование</span><span class="sxs-lookup"><span data-stu-id="5e640-113">Recommended usage</span></span>

<span data-ttu-id="5e640-114">Так как `ExcelApiOnline` API поддерживаются только в Excel в Интернете, надстройка должна проверить, поддерживается ли набор требований, прежде чем вызывать эти API.</span><span class="sxs-lookup"><span data-stu-id="5e640-114">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="5e640-115">Это позволяет избежать вызова Интернет-API на другой платформе.</span><span class="sxs-lookup"><span data-stu-id="5e640-115">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="5e640-116">После того как API находится в наборе требований к нескольким платформам, необходимо удалить или изменить `isSetSupported` проверку.</span><span class="sxs-lookup"><span data-stu-id="5e640-116">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="5e640-117">При этом функция надстройки будет включена на других платформах.</span><span class="sxs-lookup"><span data-stu-id="5e640-117">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="5e640-118">При внесении изменений обязательно протестируйте эту функцию на этих платформах.</span><span class="sxs-lookup"><span data-stu-id="5e640-118">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5e640-119">В манифесте невозможно указать `ExcelApiOnline 1.1` требования для активации.</span><span class="sxs-lookup"><span data-stu-id="5e640-119">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="5e640-120">Значение не является допустимым для использования в [элементе Set](../manifest/set.md).</span><span class="sxs-lookup"><span data-stu-id="5e640-120">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="5e640-121">Список API</span><span class="sxs-lookup"><span data-stu-id="5e640-121">API list</span></span>

| <span data-ttu-id="5e640-122">Класс</span><span class="sxs-lookup"><span data-stu-id="5e640-122">Class</span></span> | <span data-ttu-id="5e640-123">Поля</span><span class="sxs-lookup"><span data-stu-id="5e640-123">Fields</span></span> | <span data-ttu-id="5e640-124">Описание</span><span class="sxs-lookup"><span data-stu-id="5e640-124">Description</span></span> |
|:---|:---|:---|
|[<span data-ttu-id="5e640-125">Range</span><span class="sxs-lookup"><span data-stu-id="5e640-125">Range</span></span>](/javascript/api/excel/excel.range)|[<span data-ttu-id="5e640-126">Жетмержедареас ()</span><span class="sxs-lookup"><span data-stu-id="5e640-126">getMergedAreas()</span></span>](/javascript/api/excel/excel.range#getmergedareas--)|<span data-ttu-id="5e640-127">Возвращает объект RangeAreas, представляющий Объединенные области в этом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="5e640-127">Returns a RangeAreas object that represents the merged areas in this range.</span></span>|

## <a name="see-also"></a><span data-ttu-id="5e640-128">См. также</span><span class="sxs-lookup"><span data-stu-id="5e640-128">See also</span></span>

- [<span data-ttu-id="5e640-129">Справочная документация по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="5e640-129">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [<span data-ttu-id="5e640-130">Предварительные версии API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="5e640-130">Excel JavaScript preview APIs</span></span>](excel-preview-apis.md)
- [<span data-ttu-id="5e640-131">Наборы обязательных элементов API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="5e640-131">Excel JavaScript API requirement sets</span></span>](excel-api-requirement-sets.md)
