---
title: Набор обязательных элементов API JavaScript для Excel Online
description: Сведения о наборе требований Ексцелапионлине
ms.date: 11/19/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e583c9832f04e17dc1c82d38d056fe2749888a77
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757494"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="d6a0c-103">Набор обязательных элементов API JavaScript для Excel Online</span><span class="sxs-lookup"><span data-stu-id="d6a0c-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="d6a0c-104">Набор `ExcelApiOnline` требований является особым набором требований, включающим функции, доступные только для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="d6a0c-105">API в этом наборе обязательных элементов считаются рабочими API (не подчиняются недокументированным изменениям поведения или структурным изменениям) для Excel на веб-узле.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web host.</span></span> <span data-ttu-id="d6a0c-106">`ExcelApiOnline`считаются API-интерфейсами Preview для других платформ (Windows, Mac, iOS) и могут не поддерживаться ни одной из этих платформ.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-106">`ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="d6a0c-107">Если API в `ExcelApiOnline` наборе обязательных элементов поддерживаются на всех платформах, они будут добавлены к следующему набору`ExcelApi 1.[NEXT]`обязательных требований ().</span><span class="sxs-lookup"><span data-stu-id="d6a0c-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="d6a0c-108">После того как новое требование будет общедоступным, эти API будут `ExcelApiOnline`удалены из.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="d6a0c-109">Это можно считать похожим процессом повышения роли API, который перемещается из бета-версии в выпуск.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-109">Think of this as a similar promotion process as an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d6a0c-110">`ExcelApiOnline`является надмножеством набора последних пронумерованных требований.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-110">`ExcelApiOnline` is superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d6a0c-111">`ExcelApiOnline 1.1`— Единственная версия интерфейсов API, предназначенных только для интерактивного подключения.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="d6a0c-112">Это связано с тем, что Excel в Интернете всегда будет иметь одну версию, доступную для пользователей с последней версией.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

## <a name="recommended-usage"></a><span data-ttu-id="d6a0c-113">Рекомендуемое использование</span><span class="sxs-lookup"><span data-stu-id="d6a0c-113">Recommended usage</span></span>

<span data-ttu-id="d6a0c-114">Так `ExcelApiOnline` как API поддерживаются только в Excel в Интернете, надстройка должна проверить, поддерживается ли набор требований, прежде чем вызывать эти API.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-114">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="d6a0c-115">Это позволяет избежать вызова Интернет-API на другой платформе.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-115">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="d6a0c-116">После того как API находится в наборе требований к нескольким платформам, необходимо удалить или изменить `isSetSupported` проверку.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-116">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="d6a0c-117">При этом функция надстройки будет включена на других платформах.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-117">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="d6a0c-118">При внесении изменений обязательно протестируйте эту функцию на этих платформах.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-118">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d6a0c-119">В манифесте невозможно `ExcelApiOnline 1.1` указать требования для активации.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-119">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="d6a0c-120">Значение не является допустимым для использования в [элементе Set](../manifest/set.md).</span><span class="sxs-lookup"><span data-stu-id="d6a0c-120">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="d6a0c-121">Список API</span><span class="sxs-lookup"><span data-stu-id="d6a0c-121">API list</span></span>

<span data-ttu-id="d6a0c-122">В настоящее время интерфейсы API, доступные только в Интернете, отсутствуют.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-122">There are currently no online-only APIs.</span></span> <span data-ttu-id="d6a0c-123">Возврат к последующим добавлению новых компонентов в Excel в Интернете и поддержка API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="d6a0c-123">Check back as new features are added to Excel on the web and supported by the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="d6a0c-124">См. также</span><span class="sxs-lookup"><span data-stu-id="d6a0c-124">See also</span></span>

- [<span data-ttu-id="d6a0c-125">Справочная документация по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="d6a0c-125">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="d6a0c-126">Предварительные версии API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="d6a0c-126">Excel JavaScript preview APIs</span></span>](./excel-preview-apis.md)
- [<span data-ttu-id="d6a0c-127">Наборы обязательных элементов API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="d6a0c-127">Excel JavaScript API requirement sets</span></span>](./excel-api-requirement-sets.md)