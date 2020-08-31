---
title: Наборы обязательных элементов API JavaScript для OneNote
description: Узнайте больше о наборах обязательных элементов API JavaScript для OneNote.
ms.date: 08/24/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: c8cadacac640cbe710c9894a65ee780267066afc
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293529"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="73c23-103">Наборы обязательных элементов API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="73c23-103">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="73c23-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="73c23-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="73c23-107">В приведенной ниже таблице перечислены наборы обязательных элементов для OneNote, клиентские приложения Office, которые их поддерживают, а также версии сборок или даты выхода.</span><span class="sxs-lookup"><span data-stu-id="73c23-107">The following table lists the OneNote requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="73c23-108">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="73c23-108">Requirement set</span></span>  |  <span data-ttu-id="73c23-109">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="73c23-109">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="73c23-110">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="73c23-110">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)  | <span data-ttu-id="73c23-111">Сентябрь 2016 г.</span><span class="sxs-lookup"><span data-stu-id="73c23-111">September 2016</span></span> |  

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="73c23-112">API JavaScript для OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="73c23-112">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="73c23-113">API JavaScript для OneNote 1.1 — первая версия этого API.</span><span class="sxs-lookup"><span data-stu-id="73c23-113">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="73c23-114">Дополнительные сведения об этом API см. в статье [Обзор создания кода с помощью API JavaScript для OneNote](../../onenote/onenote-add-ins-programming-overview.md).</span><span class="sxs-lookup"><span data-stu-id="73c23-114">For details about the API, see the [OneNote JavaScript API programming overview](../../onenote/onenote-add-ins-programming-overview.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="73c23-115">Проверка поддержки обязательных элементов в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="73c23-115">Runtime requirement support check</span></span>

<span data-ttu-id="73c23-116">В среде выполнения надстройки могут проверять, поддерживает ли конкретное приложение Office набор обязательных элементов API с помощью следующей проверки.</span><span class="sxs-lookup"><span data-stu-id="73c23-116">At runtime, add-ins can check if a particular Office application supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="73c23-117">Проверка поддержки обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="73c23-117">Manifest-based requirement support check</span></span>

<span data-ttu-id="73c23-118">Используйте элемент `Requirements` в манифесте надстройки, чтобы указать ключевые наборы обязательных элементов или элементы API, которые должна использовать надстройка.</span><span class="sxs-lookup"><span data-stu-id="73c23-118">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="73c23-119">Если платформа или приложение Office не поддерживает наборы обязательных элементов или элементы API, указанные в элементе `Requirements`, надстройка не будет работать в этом приложении или на этой платформе, а также не будет отображаться в разделе "Мои надстройки".</span><span class="sxs-lookup"><span data-stu-id="73c23-119">If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="73c23-120">Ниже показана надстройка, которая загружается во всех клиентских приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="73c23-120">The following code example shows an add-in that loads in all Office client applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="73c23-121">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="73c23-121">Office Common API requirement sets</span></span>

<span data-ttu-id="73c23-122">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="73c23-122">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="73c23-123">См. также</span><span class="sxs-lookup"><span data-stu-id="73c23-123">See also</span></span>

- [<span data-ttu-id="73c23-124">Справочная документация по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="73c23-124">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="73c23-125">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="73c23-125">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="73c23-126">Указание приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="73c23-126">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="73c23-127">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="73c23-127">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
