---
title: Наборы обязательных элементов API JavaScript для OneNote
description: ''
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: e1012b337b3713f57a5d3df7f7c7ccbcf509b5aa
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940857"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="15c60-102">Наборы обязательных элементов API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="15c60-102">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="15c60-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="15c60-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="15c60-106">В приведенной ниже таблице перечислены наборы обязательных элементов для OneNote, ведущие приложения Office, которые их поддерживают, а также версии сборок или даты выхода.</span><span class="sxs-lookup"><span data-stu-id="15c60-106">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="15c60-107">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="15c60-107">Requirement set</span></span>  |  <span data-ttu-id="15c60-108">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="15c60-108">Office on the web</span></span> |
|:-----|:-----|
| <span data-ttu-id="15c60-109">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="15c60-109">OneNoteApi 1.1</span></span>  | <span data-ttu-id="15c60-110">Сентябрь 2016 г.</span><span class="sxs-lookup"><span data-stu-id="15c60-110">September 2016</span></span> |

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="15c60-111">API JavaScript для OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="15c60-111">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="15c60-112">API JavaScript для OneNote 1.1 — первая версия этого API.</span><span class="sxs-lookup"><span data-stu-id="15c60-112">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="15c60-113">Дополнительные сведения об этом API см. в статье [Обзор создания кода с помощью API JavaScript для OneNote](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span><span class="sxs-lookup"><span data-stu-id="15c60-113">For details about the API, see the [OneNote JavaScript API programming overview](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="15c60-114">Проверка поддержки требований в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="15c60-114">Runtime requirement support check</span></span>

<span data-ttu-id="15c60-115">В среде выполнения надстройки могут проверять, поддерживает ли конкретный узел набор обязательных элементов API, выполнив следующие действия.</span><span class="sxs-lookup"><span data-stu-id="15c60-115">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1') === true) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="15c60-116">Проверка поддержки обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="15c60-116">Manifest-based requirement support check</span></span>

<span data-ttu-id="15c60-117">Используйте `Requirements` элемент в манифесте надстройки, чтобы указать критические наборы требований или элементы API, которые должна использовать надстройка.</span><span class="sxs-lookup"><span data-stu-id="15c60-117">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="15c60-118">Если ведущее приложение или платформа Office не поддерживает наборы требований или элементы API, указанные в `Requirements` элементе, надстройка не будет запускаться на этом узле или платформе и не будет отображаться в папке "Мои надстройки".</span><span class="sxs-lookup"><span data-stu-id="15c60-118">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="15c60-119">Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="15c60-119">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="15c60-120">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="15c60-120">Office Common API requirement sets</span></span>

<span data-ttu-id="15c60-121">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="15c60-121">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="15c60-122">См. также</span><span class="sxs-lookup"><span data-stu-id="15c60-122">See also</span></span>

- [<span data-ttu-id="15c60-123">Справочная документация по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="15c60-123">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="15c60-124">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="15c60-124">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="15c60-125">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="15c60-125">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="15c60-126">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="15c60-126">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
