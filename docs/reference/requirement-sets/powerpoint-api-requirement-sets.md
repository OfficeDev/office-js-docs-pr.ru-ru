---
title: Наборы обязательных элементов API JavaScript для PowerPoint
description: Узнайте больше о наборах обязательных элементов API JavaScript для PowerPoint
ms.date: 10/26/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: cf9ab510e4b35a140c77ee958279cb85a2189fa2
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774730"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="82b05-103">Наборы обязательных элементов API JavaScript для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="82b05-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="82b05-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="82b05-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="82b05-107">В таблице ниже перечислены наборы обязательных элементов для PowerPoint, клиентские приложения Office, которые их поддерживают, а также версии сборок или даты выхода.</span><span class="sxs-lookup"><span data-stu-id="82b05-107">The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="82b05-108">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="82b05-108">Requirement set</span></span>  |  <span data-ttu-id="82b05-109">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="82b05-109">Office on Windows</span></span><br><span data-ttu-id="82b05-110">(подключено к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="82b05-110">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="82b05-111">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="82b05-111">Office on iPad</span></span><br><span data-ttu-id="82b05-112">(подключено к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="82b05-112">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="82b05-113">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="82b05-113">Office on Mac</span></span><br><span data-ttu-id="82b05-114">(подключено к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="82b05-114">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="82b05-115">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="82b05-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| [<span data-ttu-id="82b05-116">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="82b05-116">Preview</span></span>](powerpoint-preview-apis.md)  | <span data-ttu-id="82b05-117">Используйте последнюю версию Office, чтобы воспользоваться предварительными версиями API (может потребоваться присоединиться к [программе предварительной оценки Office](https://insider.office.com)).</span><span class="sxs-lookup"><span data-stu-id="82b05-117">Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://insider.office.com)).</span></span> |
| <span data-ttu-id="82b05-118">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="82b05-118">PowerPointApi 1.1</span></span> | <span data-ttu-id="82b05-119">Версия 1810 (сборка 11001.20074) или более поздняя</span><span class="sxs-lookup"><span data-stu-id="82b05-119">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="82b05-120">2.17 или более поздняя</span><span class="sxs-lookup"><span data-stu-id="82b05-120">2.17 or later</span></span> | <span data-ttu-id="82b05-121">16.19 или более поздняя</span><span class="sxs-lookup"><span data-stu-id="82b05-121">16.19 or later</span></span> | <span data-ttu-id="82b05-122">Октябрь 2018 г.</span><span class="sxs-lookup"><span data-stu-id="82b05-122">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="82b05-123">Номера версий и сборок Office</span><span class="sxs-lookup"><span data-stu-id="82b05-123">Office versions and build numbers</span></span>

<span data-ttu-id="82b05-124">Дополнительные сведения о номерах версий и сборок Office см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="82b05-124">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="82b05-125">API JavaScript для PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="82b05-125">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="82b05-126">API JavaScript для PowerPoint 1.1 содержит [единый API для создания новых презентаций](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-).</span><span class="sxs-lookup"><span data-stu-id="82b05-126">PowerPoint JavaScript API 1.1 contains a [single API to create a new presentation](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-).</span></span> <span data-ttu-id="82b05-127">Сведения об этом API см. в разделе [Создание презентации](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).</span><span class="sxs-lookup"><span data-stu-id="82b05-127">For details about the API, see [Create a presentation](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).</span></span>

## <a name="how-to-use-powerpoint-requirement-sets-at-runtime-and-in-the-manifest"></a><span data-ttu-id="82b05-128">Использование наборов обязательных элементов PowerPoint в среде выполнения и в манифесте</span><span class="sxs-lookup"><span data-stu-id="82b05-128">How to use PowerPoint requirement sets at runtime and in the manifest</span></span>

> [!NOTE]
> <span data-ttu-id="82b05-129">В этом разделе предполагается, что вы знакомы с общими сведениями о наборах обязательных элементов, изложенными в статьях [Версии и наборы обязательных элементов Office](../../develop/office-versions-and-requirement-sets.md) и [Указание приложений и обязательных элементов API Office](../../develop/specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="82b05-129">This section assumes you're familiar with the overview of requirement sets at [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md) and [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md).</span></span>

<span data-ttu-id="82b05-130">Наборы требований — это именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="82b05-130">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="82b05-131">Надстройка Office может выполнить проверку в среде выполнения или использовать указанные в манифесте наборы обязательных элементов, чтобы определить, поддерживает ли приложение Office необходимые надстройке API.</span><span class="sxs-lookup"><span data-stu-id="82b05-131">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="82b05-132">Проверка поддержки наборов обязательных элементов в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="82b05-132">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="82b05-133">В следующем примере кода показано, как определить, поддерживает ли приложение Office, в котором запускается надстройка, указанный набор обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="82b05-133">The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="82b05-134">Определение поддержки наборов обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="82b05-134">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="82b05-135">С помощью [элемента Requirements](../manifest/requirements.md) в манифесте надстройки можно указать минимальные наборы обязательных элементов и/или методы API, необходимые надстройке для активации.</span><span class="sxs-lookup"><span data-stu-id="82b05-135">You can use the [Requirements element](../manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="82b05-136">Если приложение или платформа Office не поддерживает наборы обязательных элементов или методы API, указанные в элементе манифеста `Requirements`, надстройка не будет работать в этом приложении или на этой платформе и не будет отображать список надстроек, показанный в разделе **Мои надстройки** . Если вашей надстройке для полной функциональности необходим определенный набор обязательных элементов, но она может быть полезна пользователям даже на тех платформах, которые не поддерживают этот набор, мы рекомендуем проверить поддержку обязательных элементов в среде выполнения как описано выше, а не прописывать поддержку набора обязательных элементов в манифесте.</span><span class="sxs-lookup"><span data-stu-id="82b05-136">If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins** . If your add-in requires a specific requirement set for full functionality, but it can provide value even to users on platforms that do not support the requirement set, we recommend that you check for requirement support at runtime as described above, instead of defining requirement set support in the manifest.</span></span>

<span data-ttu-id="82b05-137">В следующем примере кода показан элемент `Requirements` в манифесте надстройки, где указано, что надстройка должна загружаться во всех клиентских приложениях Office, поддерживающих набор обязательных элементов PowerPointApi версии 1.1 или более поздней.</span><span class="sxs-lookup"><span data-stu-id="82b05-137">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support PowerPointApi requirement set version 1.1 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="82b05-138">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="82b05-138">Office Common API requirement sets</span></span>

<span data-ttu-id="82b05-139">Большинство функций надстройки PowerPoint определяются набором обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="82b05-139">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="82b05-140">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="82b05-140">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="82b05-141">См. также</span><span class="sxs-lookup"><span data-stu-id="82b05-141">See also</span></span>

- [<span data-ttu-id="82b05-142">Справочная документация по API JavaScript для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="82b05-142">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="82b05-143">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="82b05-143">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="82b05-144">Указание приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="82b05-144">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="82b05-145">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="82b05-145">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
