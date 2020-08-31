---
title: Наборы обязательных элементов API JavaScript для PowerPoint
description: Узнайте больше о наборах обязательных элементов API JavaScript для PowerPoint
ms.date: 07/10/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: b2b5d4b7b5a0677812f227b6a32683c35bbf1662
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293508"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="64e34-103">Наборы обязательных элементов API JavaScript для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="64e34-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="64e34-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="64e34-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="64e34-107">В таблице ниже перечислены наборы обязательных элементов для PowerPoint, клиентские приложения Office, которые их поддерживают, а также версии сборок или даты выхода.</span><span class="sxs-lookup"><span data-stu-id="64e34-107">The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="64e34-108">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="64e34-108">Requirement set</span></span>  |  <span data-ttu-id="64e34-109">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="64e34-109">Office on Windows</span></span><br><span data-ttu-id="64e34-110">(подключено к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="64e34-110">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="64e34-111">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="64e34-111">Office on iPad</span></span><br><span data-ttu-id="64e34-112">(подключено к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="64e34-112">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="64e34-113">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="64e34-113">Office on Mac</span></span><br><span data-ttu-id="64e34-114">(подключено к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="64e34-114">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="64e34-115">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="64e34-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="64e34-116">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="64e34-116">PowerPointApi 1.1</span></span> | <span data-ttu-id="64e34-117">Версия 1810 (сборка 11001.20074) или более поздняя</span><span class="sxs-lookup"><span data-stu-id="64e34-117">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="64e34-118">2.17 или более поздняя</span><span class="sxs-lookup"><span data-stu-id="64e34-118">2.17 or later</span></span> | <span data-ttu-id="64e34-119">16.19 или более поздняя</span><span class="sxs-lookup"><span data-stu-id="64e34-119">16.19 or later</span></span> | <span data-ttu-id="64e34-120">Октябрь 2018 г.</span><span class="sxs-lookup"><span data-stu-id="64e34-120">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="64e34-121">Номера версий и сборок Office</span><span class="sxs-lookup"><span data-stu-id="64e34-121">Office versions and build numbers</span></span>

<span data-ttu-id="64e34-122">Дополнительные сведения о номерах версий и сборок Office см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="64e34-122">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="64e34-123">API JavaScript для PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="64e34-123">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="64e34-124">API JavaScript для PowerPoint 1.1 включает один API для создания новой презентации.</span><span class="sxs-lookup"><span data-stu-id="64e34-124">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="64e34-125">Дополнительные сведения об API см. в статье [API JavaScript для PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="64e34-125">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="64e34-126">Проверка поддержки обязательных элементов в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="64e34-126">Runtime requirement support check</span></span>

<span data-ttu-id="64e34-127">В среде выполнения надстройки могут проверять, поддерживает ли конкретное приложение набор обязательных элементов API с помощью следующей проверки.</span><span class="sxs-lookup"><span data-stu-id="64e34-127">At runtime, add-ins can check if a particular application supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="64e34-128">Проверка поддержки обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="64e34-128">Manifest-based requirement support check</span></span>

<span data-ttu-id="64e34-129">Используйте элемент `Requirements` в манифесте надстройки, чтобы указать ключевые наборы обязательных элементов или элементы API, которые должна использовать надстройка.</span><span class="sxs-lookup"><span data-stu-id="64e34-129">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="64e34-130">Если платформа или приложение Office не поддерживает наборы обязательных элементов или элементы API, указанные в элементе `Requirements`, надстройка не будет работать в этом приложении или на этой платформе, а также не будет отображаться в разделе "Мои надстройки".</span><span class="sxs-lookup"><span data-stu-id="64e34-130">If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="64e34-131">Ниже показана надстройка, которая загружается во всех клиентских приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="64e34-131">The following code example shows an add-in that loads in all Office client applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="64e34-132">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="64e34-132">Office Common API requirement sets</span></span>

<span data-ttu-id="64e34-133">Большинство функций надстройки PowerPoint определяются набором обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="64e34-133">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="64e34-134">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="64e34-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="64e34-135">См. также</span><span class="sxs-lookup"><span data-stu-id="64e34-135">See also</span></span>

- [<span data-ttu-id="64e34-136">Справочная документация по API JavaScript для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="64e34-136">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="64e34-137">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="64e34-137">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="64e34-138">Указание приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="64e34-138">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="64e34-139">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="64e34-139">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
