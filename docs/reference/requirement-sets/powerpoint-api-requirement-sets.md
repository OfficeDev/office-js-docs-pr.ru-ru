---
title: Наборы обязательных элементов API JavaScript для PowerPoint
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 5bba2354cabba3c3ccd4ddf38d3e03c25a32b8a9
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950959"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="75bed-102">Наборы обязательных элементов API JavaScript для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="75bed-102">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="75bed-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="75bed-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="75bed-106">В приведенной ниже таблице перечислены наборы обязательных элементов для PowerPoint, ведущие приложения Office, которые их поддерживают, а также версии сборок или даты выхода.</span><span class="sxs-lookup"><span data-stu-id="75bed-106">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="75bed-107">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="75bed-107">Requirement set</span></span>  |  <span data-ttu-id="75bed-108">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="75bed-108">Office on Windows</span></span><br><span data-ttu-id="75bed-109">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75bed-109">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="75bed-110">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="75bed-110">Office on iPad</span></span><br><span data-ttu-id="75bed-111">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75bed-111">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="75bed-112">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="75bed-112">Office on Mac</span></span><br><span data-ttu-id="75bed-113">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75bed-113">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="75bed-114">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="75bed-114">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="75bed-115">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="75bed-115">PowerPointApi 1.1</span></span> | <span data-ttu-id="75bed-116">Версия 1810 (сборка 11001.20074) или более поздняя</span><span class="sxs-lookup"><span data-stu-id="75bed-116">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="75bed-117">2.17 или более поздняя</span><span class="sxs-lookup"><span data-stu-id="75bed-117">2.17 or later</span></span> | <span data-ttu-id="75bed-118">16.19 или более поздняя</span><span class="sxs-lookup"><span data-stu-id="75bed-118">16.19 or later</span></span> | <span data-ttu-id="75bed-119">Октябрь 2018 г.</span><span class="sxs-lookup"><span data-stu-id="75bed-119">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="75bed-120">Номера версий и сборок Office</span><span class="sxs-lookup"><span data-stu-id="75bed-120">Office versions and build numbers</span></span>

<span data-ttu-id="75bed-121">Дополнительные сведения о номерах версий и сборок Office см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="75bed-121">For more information about Office versions and build numbers, see:</span></span>

- <span data-ttu-id="75bed-122">[Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);</span><span class="sxs-lookup"><span data-stu-id="75bed-122">[Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span></span>
- <span data-ttu-id="75bed-123">[Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);</span><span class="sxs-lookup"><span data-stu-id="75bed-123">[What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)</span></span>
- <span data-ttu-id="75bed-124">[Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);</span><span class="sxs-lookup"><span data-stu-id="75bed-124">[Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span></span>

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="75bed-125">API JavaScript для PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="75bed-125">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="75bed-126">API JavaScript для PowerPoint 1.1 включает один API для создания новой презентации.</span><span class="sxs-lookup"><span data-stu-id="75bed-126">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="75bed-127">Дополнительные сведения об API см. в статье [API JavaScript для PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="75bed-127">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="75bed-128">Проверка поддержки обязательных элементов в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="75bed-128">Runtime requirement support check</span></span>

<span data-ttu-id="75bed-129">В среде выполнения надстройки могут проверять, поддерживает ли ведущее приложение набор обязательных элементов API, выполняя следующую проверку.</span><span class="sxs-lookup"><span data-stu-id="75bed-129">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="75bed-130">Проверка поддержки обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="75bed-130">Manifest-based requirement support check</span></span>

<span data-ttu-id="75bed-131">Используйте элемент `Requirements` в манифесте надстройки, чтобы указать ключевые наборы обязательных элементов или элементы API, которые должна использовать надстройка.</span><span class="sxs-lookup"><span data-stu-id="75bed-131">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="75bed-132">Если платформа или ведущее приложение Office не поддерживает наборы обязательных элементов или элементы API, указанные в элементе `Requirements`, надстройка не будет работать в этом ведущем приложении или на этой платформе, а также не будет отображаться в разделе "Мои надстройки".</span><span class="sxs-lookup"><span data-stu-id="75bed-132">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="75bed-133">Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="75bed-133">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="75bed-134">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="75bed-134">Office Common API requirement sets</span></span>

<span data-ttu-id="75bed-135">Большинство функций надстройки PowerPoint определяются набором обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="75bed-135">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="75bed-136">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="75bed-136">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="75bed-137">См. также</span><span class="sxs-lookup"><span data-stu-id="75bed-137">See also</span></span>

- [<span data-ttu-id="75bed-138">Справочная документация по API JavaScript для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="75bed-138">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="75bed-139">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="75bed-139">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="75bed-140">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="75bed-140">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="75bed-141">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="75bed-141">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
