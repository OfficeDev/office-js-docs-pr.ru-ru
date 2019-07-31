---
title: Наборы обязательных элементов API JavaScript для PowerPoint
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 4f64654a4130cc0d4bf96d9c59e364e77c808748
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/31/2019
ms.locfileid: "35941151"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="56a3d-102">Наборы обязательных элементов API JavaScript для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="56a3d-102">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="56a3d-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="56a3d-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="56a3d-106">В следующей таблице перечислены наборы требований PowerPoint, ведущие приложения Office, которые поддерживают эти наборы требований, а также версии сборки или Дата доступности.</span><span class="sxs-lookup"><span data-stu-id="56a3d-106">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="56a3d-107">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="56a3d-107">Requirement set</span></span>  |  <span data-ttu-id="56a3d-108">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="56a3d-108">Office on Windows</span></span><br><span data-ttu-id="56a3d-109">(подключено к подписке Office 365)</span><span class="sxs-lookup"><span data-stu-id="56a3d-109">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="56a3d-110">Office на iPad</span><span class="sxs-lookup"><span data-stu-id="56a3d-110">Office on iPad</span></span><br><span data-ttu-id="56a3d-111">(подключено к подписке Office 365)</span><span class="sxs-lookup"><span data-stu-id="56a3d-111">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="56a3d-112">Office на Mac</span><span class="sxs-lookup"><span data-stu-id="56a3d-112">Office on Mac</span></span><br><span data-ttu-id="56a3d-113">(подключено к подписке Office 365)</span><span class="sxs-lookup"><span data-stu-id="56a3d-113">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="56a3d-114">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="56a3d-114">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="56a3d-115">Поверпоинтапи 1,1</span><span class="sxs-lookup"><span data-stu-id="56a3d-115">PowerPointApi 1.1</span></span> | <span data-ttu-id="56a3d-116">Версия 1810 (сборка 11001,20074) или более поздняя</span><span class="sxs-lookup"><span data-stu-id="56a3d-116">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="56a3d-117">2.17 или более поздняя</span><span class="sxs-lookup"><span data-stu-id="56a3d-117">2.17 or later</span></span> | <span data-ttu-id="56a3d-118">16,19 или более поздняя версия</span><span class="sxs-lookup"><span data-stu-id="56a3d-118">16.19 or later</span></span> | <span data-ttu-id="56a3d-119">Октябрь 2018 г.</span><span class="sxs-lookup"><span data-stu-id="56a3d-119">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="56a3d-120">Номера версий и сборок Office</span><span class="sxs-lookup"><span data-stu-id="56a3d-120">Office versions and build numbers</span></span>

<span data-ttu-id="56a3d-121">Более подробную информацию о версиях Office и номерах сборок можно узнать в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="56a3d-121">For more information about Office versions and build numbers, see:</span></span>

- <span data-ttu-id="56a3d-122">[Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);</span><span class="sxs-lookup"><span data-stu-id="56a3d-122">[Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span></span>
- <span data-ttu-id="56a3d-123">[Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);</span><span class="sxs-lookup"><span data-stu-id="56a3d-123">[What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)</span></span>
- <span data-ttu-id="56a3d-124">[Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);</span><span class="sxs-lookup"><span data-stu-id="56a3d-124">[Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span></span>

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="56a3d-125">API JavaScript для PowerPoint 1,1</span><span class="sxs-lookup"><span data-stu-id="56a3d-125">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="56a3d-126">API JavaScript для PowerPoint 1,1 содержит один API для создания новой презентации.</span><span class="sxs-lookup"><span data-stu-id="56a3d-126">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="56a3d-127">Дополнительные сведения об API можно найти в статье [API JavaScript для PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="56a3d-127">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="56a3d-128">Проверка поддержки требований в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="56a3d-128">Runtime requirement support check</span></span>

<span data-ttu-id="56a3d-129">В среде выполнения надстройки могут проверять, поддерживает ли конкретный узел набор обязательных элементов API, выполнив следующие действия.</span><span class="sxs-lookup"><span data-stu-id="56a3d-129">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="56a3d-130">Проверка поддержки обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="56a3d-130">Manifest-based requirement support check</span></span>

<span data-ttu-id="56a3d-131">Используйте `Requirements` элемент в манифесте надстройки, чтобы указать критические наборы требований или элементы API, которые должна использовать надстройка.</span><span class="sxs-lookup"><span data-stu-id="56a3d-131">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="56a3d-132">Если ведущее приложение или платформа Office не поддерживает наборы требований или элементы API, указанные в `Requirements` элементе, надстройка не будет запускаться на этом узле или платформе и не будет отображаться в папке "Мои надстройки".</span><span class="sxs-lookup"><span data-stu-id="56a3d-132">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="56a3d-133">Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="56a3d-133">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="56a3d-134">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="56a3d-134">Office Common API requirement sets</span></span>

<span data-ttu-id="56a3d-135">Большинство функциональных возможностей надстройки PowerPoint берутся из общего набора API.</span><span class="sxs-lookup"><span data-stu-id="56a3d-135">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="56a3d-136">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="56a3d-136">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="56a3d-137">См. также</span><span class="sxs-lookup"><span data-stu-id="56a3d-137">See also</span></span>

- [<span data-ttu-id="56a3d-138">Справочная документация по API JavaScript для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="56a3d-138">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="56a3d-139">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="56a3d-139">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="56a3d-140">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="56a3d-140">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="56a3d-141">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="56a3d-141">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
