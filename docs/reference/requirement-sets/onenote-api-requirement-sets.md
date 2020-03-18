---
title: Наборы обязательных элементов API JavaScript для OneNote
description: Узнайте больше о наборах требований OneNote JavaScript API
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 7717ff20fe4a7f29621a30df7d01d122111021db
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717448"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="3bc1b-103">Наборы обязательных элементов API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="3bc1b-103">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="3bc1b-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3bc1b-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="3bc1b-107">В приведенной ниже таблице перечислены наборы обязательных элементов для OneNote, ведущие приложения Office, которые их поддерживают, а также версии сборок или даты выхода.</span><span class="sxs-lookup"><span data-stu-id="3bc1b-107">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="3bc1b-108">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="3bc1b-108">Requirement set</span></span>  |  <span data-ttu-id="3bc1b-109">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="3bc1b-109">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="3bc1b-110">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="3bc1b-110">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)  | <span data-ttu-id="3bc1b-111">Сентябрь 2016 г.</span><span class="sxs-lookup"><span data-stu-id="3bc1b-111">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="3bc1b-112">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="3bc1b-112">Office Common API requirement sets</span></span>

<span data-ttu-id="3bc1b-113">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3bc1b-113">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="3bc1b-114">API JavaScript для OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="3bc1b-114">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="3bc1b-115">API JavaScript для OneNote 1.1 — первая версия этого API.</span><span class="sxs-lookup"><span data-stu-id="3bc1b-115">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="3bc1b-116">Дополнительные сведения об этом API см. в статье [Обзор создания кода с помощью API JavaScript для OneNote](../../onenote/onenote-add-ins-programming-overview.md).</span><span class="sxs-lookup"><span data-stu-id="3bc1b-116">For details about the API, see the [OneNote JavaScript API programming overview](../../onenote/onenote-add-ins-programming-overview.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="3bc1b-117">Проверка поддержки требований в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="3bc1b-117">Runtime requirement support check</span></span>

<span data-ttu-id="3bc1b-118">В среде выполнения надстройки могут проверять, поддерживает ли ведущее приложение набор обязательных элементов API, выполняя следующую проверку.</span><span class="sxs-lookup"><span data-stu-id="3bc1b-118">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="3bc1b-119">Проверка поддержки обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="3bc1b-119">Manifest-based requirement support check</span></span>

<span data-ttu-id="3bc1b-120">Используйте элемент `Requirements` в манифесте надстройки, чтобы указать ключевые наборы обязательных элементов или элементы API, которые должна использовать надстройка.</span><span class="sxs-lookup"><span data-stu-id="3bc1b-120">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="3bc1b-121">Если платформа или ведущее приложение Office не поддерживает наборы обязательных элементов или элементы API, указанные в элементе `Requirements`, надстройка не будет работать в этом ведущем приложении или на этой платформе, а также не будет отображаться в разделе "Мои надстройки".</span><span class="sxs-lookup"><span data-stu-id="3bc1b-121">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="3bc1b-122">Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="3bc1b-122">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="3bc1b-123">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="3bc1b-123">Office Common API requirement sets</span></span>

<span data-ttu-id="3bc1b-124">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3bc1b-124">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3bc1b-125">См. также</span><span class="sxs-lookup"><span data-stu-id="3bc1b-125">See also</span></span>

- [<span data-ttu-id="3bc1b-126">Справочная документация по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="3bc1b-126">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="3bc1b-127">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="3bc1b-127">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="3bc1b-128">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="3bc1b-128">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="3bc1b-129">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="3bc1b-129">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
