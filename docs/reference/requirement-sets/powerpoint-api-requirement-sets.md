---
title: Наборы обязательных элементов API JavaScript для PowerPoint
description: Узнайте больше о наборах требований PowerPoint JavaScript API
ms.date: 03/11/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: a82d73087b19fbce12f571a2bad61e866ab62f86
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611332"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="d2468-103">Наборы обязательных элементов API JavaScript для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d2468-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="d2468-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="d2468-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="d2468-107">В приведенной ниже таблице перечислены наборы обязательных элементов для PowerPoint, ведущие приложения Office, которые их поддерживают, а также версии сборок или даты выхода.</span><span class="sxs-lookup"><span data-stu-id="d2468-107">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="d2468-108">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="d2468-108">Requirement set</span></span>  |  <span data-ttu-id="d2468-109">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="d2468-109">Office on Windows</span></span><br><span data-ttu-id="d2468-110">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2468-110">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="d2468-111">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="d2468-111">Office on iPad</span></span><br><span data-ttu-id="d2468-112">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2468-112">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="d2468-113">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d2468-113">Office on Mac</span></span><br><span data-ttu-id="d2468-114">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2468-114">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="d2468-115">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d2468-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="d2468-116">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="d2468-116">PowerPointApi 1.1</span></span> | <span data-ttu-id="d2468-117">Версия 1810 (сборка 11001.20074) или более поздняя</span><span class="sxs-lookup"><span data-stu-id="d2468-117">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="d2468-118">2.17 или более поздняя</span><span class="sxs-lookup"><span data-stu-id="d2468-118">2.17 or later</span></span> | <span data-ttu-id="d2468-119">16.19 или более поздняя</span><span class="sxs-lookup"><span data-stu-id="d2468-119">16.19 or later</span></span> | <span data-ttu-id="d2468-120">Октябрь 2018 г.</span><span class="sxs-lookup"><span data-stu-id="d2468-120">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="d2468-121">Номера версий и сборок Office</span><span class="sxs-lookup"><span data-stu-id="d2468-121">Office versions and build numbers</span></span>

<span data-ttu-id="d2468-122">Дополнительные сведения о номерах версий и сборок Office см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="d2468-122">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="d2468-123">API JavaScript для PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="d2468-123">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="d2468-124">API JavaScript для PowerPoint 1.1 включает один API для создания новой презентации.</span><span class="sxs-lookup"><span data-stu-id="d2468-124">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="d2468-125">Дополнительные сведения об API см. в статье [API JavaScript для PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="d2468-125">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="d2468-126">Проверка поддержки обязательных элементов в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="d2468-126">Runtime requirement support check</span></span>

<span data-ttu-id="d2468-127">В среде выполнения надстройки могут проверять, поддерживает ли ведущее приложение набор обязательных элементов API, выполняя следующую проверку.</span><span class="sxs-lookup"><span data-stu-id="d2468-127">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="d2468-128">Проверка поддержки обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="d2468-128">Manifest-based requirement support check</span></span>

<span data-ttu-id="d2468-129">Используйте элемент `Requirements` в манифесте надстройки, чтобы указать ключевые наборы обязательных элементов или элементы API, которые должна использовать надстройка.</span><span class="sxs-lookup"><span data-stu-id="d2468-129">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="d2468-130">Если платформа или ведущее приложение Office не поддерживает наборы обязательных элементов или элементы API, указанные в элементе `Requirements`, надстройка не будет работать в этом ведущем приложении или на этой платформе, а также не будет отображаться в разделе "Мои надстройки".</span><span class="sxs-lookup"><span data-stu-id="d2468-130">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="d2468-131">Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="d2468-131">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="d2468-132">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="d2468-132">Office Common API requirement sets</span></span>

<span data-ttu-id="d2468-133">Большинство функций надстройки PowerPoint определяются набором обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="d2468-133">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="d2468-134">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="d2468-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d2468-135">См. также</span><span class="sxs-lookup"><span data-stu-id="d2468-135">See also</span></span>

- [<span data-ttu-id="d2468-136">Справочная документация по API JavaScript для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d2468-136">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="d2468-137">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="d2468-137">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="d2468-138">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d2468-138">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="d2468-139">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d2468-139">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
