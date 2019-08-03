---
title: Наборы обязательных элементов API JavaScript для OneNote
description: ''
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 3a1e5133b36af612156fb272651f1775e916a0fe
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064874"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="76ef2-102">Наборы обязательных элементов API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="76ef2-102">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="76ef2-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="76ef2-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="76ef2-106">В приведенной ниже таблице перечислены наборы обязательных элементов для OneNote, ведущие приложения Office, которые их поддерживают, а также версии сборок или даты выхода.</span><span class="sxs-lookup"><span data-stu-id="76ef2-106">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="76ef2-107">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="76ef2-107">Requirement set</span></span>  |  <span data-ttu-id="76ef2-108">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="76ef2-108">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="76ef2-109">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="76ef2-109">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)  | <span data-ttu-id="76ef2-110">Сентябрь 2016 г.</span><span class="sxs-lookup"><span data-stu-id="76ef2-110">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="76ef2-111">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="76ef2-111">Office Common API requirement sets</span></span>

<span data-ttu-id="76ef2-112">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="76ef2-112">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="76ef2-113">API JavaScript для OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="76ef2-113">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="76ef2-114">API JavaScript для OneNote 1.1 — первая версия этого API.</span><span class="sxs-lookup"><span data-stu-id="76ef2-114">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="76ef2-115">Дополнительные сведения об этом API см. в статье [Обзор создания кода с помощью API JavaScript для OneNote](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span><span class="sxs-lookup"><span data-stu-id="76ef2-115">For details about the API, see the [OneNote JavaScript API programming overview](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="76ef2-116">Проверка поддержки требований в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="76ef2-116">Runtime requirement support check</span></span>

<span data-ttu-id="76ef2-117">В среде выполнения надстройки могут проверять, поддерживает ли конкретный узел набор обязательных элементов API, выполнив следующие действия.</span><span class="sxs-lookup"><span data-stu-id="76ef2-117">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="76ef2-118">Проверка поддержки обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="76ef2-118">Manifest-based requirement support check</span></span>

<span data-ttu-id="76ef2-119">Используйте `Requirements` элемент в манифесте надстройки, чтобы указать критические наборы требований или элементы API, которые должна использовать надстройка.</span><span class="sxs-lookup"><span data-stu-id="76ef2-119">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="76ef2-120">Если ведущее приложение или платформа Office не поддерживает наборы требований или элементы API, указанные в `Requirements` элементе, надстройка не будет запускаться на этом узле или платформе и не будет отображаться в папке "Мои надстройки".</span><span class="sxs-lookup"><span data-stu-id="76ef2-120">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="76ef2-121">Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="76ef2-121">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="76ef2-122">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="76ef2-122">Office Common API requirement sets</span></span>

<span data-ttu-id="76ef2-123">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="76ef2-123">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="76ef2-124">См. также</span><span class="sxs-lookup"><span data-stu-id="76ef2-124">See also</span></span>

- [<span data-ttu-id="76ef2-125">Справочная документация по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="76ef2-125">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="76ef2-126">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="76ef2-126">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="76ef2-127">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="76ef2-127">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="76ef2-128">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="76ef2-128">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
