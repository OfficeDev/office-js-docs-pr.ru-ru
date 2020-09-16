---
title: Набор обязательных элементов API JavaScript для Excel Online
description: Сведения о наборе требований Ексцелапионлине.
ms.date: 09/15/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 29f5826ba2adbf18b79033b83254b046210015fe
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819807"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="d8da8-103">Набор обязательных элементов API JavaScript для Excel Online</span><span class="sxs-lookup"><span data-stu-id="d8da8-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="d8da8-104">`ExcelApiOnline`Набор требований является особым набором требований, включающим функции, доступные только для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="d8da8-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="d8da8-105">API в этом наборе обязательных элементов считаются рабочими API (не подчиняются недокументированным изменениям поведения или структурным изменениям) для Excel в веб-приложении.</span><span class="sxs-lookup"><span data-stu-id="d8da8-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web application.</span></span> <span data-ttu-id="d8da8-106">`ExcelApiOnline` считаются API-интерфейсами Preview для других платформ (Windows, Mac, iOS) и могут не поддерживаться ни одной из этих платформ.</span><span class="sxs-lookup"><span data-stu-id="d8da8-106">`ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="d8da8-107">Если API в `ExcelApiOnline` наборе обязательных элементов поддерживаются на всех платформах, они будут добавлены к следующему набору обязательных требований ( `ExcelApi 1.[NEXT]` ).</span><span class="sxs-lookup"><span data-stu-id="d8da8-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="d8da8-108">После того как новое требование будет общедоступным, эти API будут удалены из `ExcelApiOnline` .</span><span class="sxs-lookup"><span data-stu-id="d8da8-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="d8da8-109">Это можно считать похожим процессом повышения роли API, который перемещается из бета-версии в выпуск.</span><span class="sxs-lookup"><span data-stu-id="d8da8-109">Think of this as a similar promotion process as an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d8da8-110">`ExcelApiOnline` является надмножеством набора последних пронумерованных требований.</span><span class="sxs-lookup"><span data-stu-id="d8da8-110">`ExcelApiOnline` is superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d8da8-111">`ExcelApiOnline 1.1` — Единственная версия интерфейсов API, предназначенных только для интерактивного подключения.</span><span class="sxs-lookup"><span data-stu-id="d8da8-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="d8da8-112">Это связано с тем, что Excel в Интернете всегда будет иметь одну версию, доступную для пользователей с последней версией.</span><span class="sxs-lookup"><span data-stu-id="d8da8-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

## <a name="recommended-usage"></a><span data-ttu-id="d8da8-113">Рекомендуемое использование</span><span class="sxs-lookup"><span data-stu-id="d8da8-113">Recommended usage</span></span>

<span data-ttu-id="d8da8-114">Так как `ExcelApiOnline` API поддерживаются только в Excel в Интернете, надстройка должна проверить, поддерживается ли набор требований, прежде чем вызывать эти API.</span><span class="sxs-lookup"><span data-stu-id="d8da8-114">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="d8da8-115">Это позволяет избежать вызова Интернет-API на другой платформе.</span><span class="sxs-lookup"><span data-stu-id="d8da8-115">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="d8da8-116">После того как API находится в наборе требований к нескольким платформам, необходимо удалить или изменить `isSetSupported` проверку.</span><span class="sxs-lookup"><span data-stu-id="d8da8-116">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="d8da8-117">При этом функция надстройки будет включена на других платформах.</span><span class="sxs-lookup"><span data-stu-id="d8da8-117">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="d8da8-118">При внесении изменений обязательно протестируйте эту функцию на этих платформах.</span><span class="sxs-lookup"><span data-stu-id="d8da8-118">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d8da8-119">В манифесте невозможно указать `ExcelApiOnline 1.1` требования для активации.</span><span class="sxs-lookup"><span data-stu-id="d8da8-119">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="d8da8-120">Значение не является допустимым для использования в [элементе Set](../manifest/set.md).</span><span class="sxs-lookup"><span data-stu-id="d8da8-120">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="d8da8-121">Список API</span><span class="sxs-lookup"><span data-stu-id="d8da8-121">API list</span></span>

<span data-ttu-id="d8da8-122">В настоящее время интерфейсы API в наборе обязательных элементов отсутствуют `ExcelApiOnline` .</span><span class="sxs-lookup"><span data-stu-id="d8da8-122">There are currently no APIs in the `ExcelApiOnline` requirement set.</span></span> <span data-ttu-id="d8da8-123">Все интерфейсы API, которые ранее были частью этого набора, превышены до набора обязательных наборов требований и доступны на всех платформах.</span><span class="sxs-lookup"><span data-stu-id="d8da8-123">All the APIs that were previously a part of this set have graduated to a numbered requirement set and are available across all platforms.</span></span>

## <a name="see-also"></a><span data-ttu-id="d8da8-124">См. также</span><span class="sxs-lookup"><span data-stu-id="d8da8-124">See also</span></span>

- [<span data-ttu-id="d8da8-125">Справочная документация по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="d8da8-125">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [<span data-ttu-id="d8da8-126">Предварительные версии API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="d8da8-126">Excel JavaScript preview APIs</span></span>](excel-preview-apis.md)
- [<span data-ttu-id="d8da8-127">Наборы обязательных элементов API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="d8da8-127">Excel JavaScript API requirement sets</span></span>](excel-api-requirement-sets.md)
