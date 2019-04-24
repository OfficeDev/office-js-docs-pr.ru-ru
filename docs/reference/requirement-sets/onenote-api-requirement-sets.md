---
title: Наборы обязательных элементов API JavaScript для OneNote
description: ''
ms.date: 03/19/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 287e405955477a98854b1df4a81fe90ec16e5bbc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450102"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="e2768-102">Наборы обязательных элементов API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="e2768-102">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="e2768-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="e2768-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="e2768-106">В приведенной ниже таблице перечислены наборы обязательных элементов для OneNote, ведущие приложения Office, которые их поддерживают, а также версии сборок или даты выхода.</span><span class="sxs-lookup"><span data-stu-id="e2768-106">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="e2768-107">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="e2768-107">Requirement set</span></span>  |  <span data-ttu-id="e2768-108">Office Online</span><span class="sxs-lookup"><span data-stu-id="e2768-108">Office Online</span></span> | 
|:-----|:-----|
| <span data-ttu-id="e2768-109">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="e2768-109">OneNoteApi 1.1</span></span>  | <span data-ttu-id="e2768-110">Сентябрь 2016 г.</span><span class="sxs-lookup"><span data-stu-id="e2768-110">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="e2768-111">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="e2768-111">Office Common API requirement sets</span></span>

<span data-ttu-id="e2768-112">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e2768-112">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="e2768-113">API JavaScript для OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="e2768-113">OneNote JavaScript API 1.1</span></span> 

<span data-ttu-id="e2768-114">API JavaScript для OneNote 1.1 — первая версия этого API.</span><span class="sxs-lookup"><span data-stu-id="e2768-114">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="e2768-115">Дополнительные сведения об этом API см. в статье [Обзор создания кода с помощью API JavaScript для OneNote](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span><span class="sxs-lookup"><span data-stu-id="e2768-115">For details about the API, see the [OneNote JavaScript API programming overview](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="e2768-116">Проверка поддержки требований в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="e2768-116">Runtime requirement support check</span></span>

<span data-ttu-id="e2768-117">Во время выполнения кода надстройки могут проверять, поддерживает ли ведущее приложение набор обязательных элементов API, выполняя следующую проверку:</span><span class="sxs-lookup"><span data-stu-id="e2768-117">During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check:</span></span> 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="e2768-118">Проверка поддержки обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="e2768-118">Manifest-based requirement support check</span></span>

<span data-ttu-id="e2768-p103">Используйте элемент Requirements в манифесте надстройки, чтобы указать ключевые наборы требований или элементы API, которые должна использовать надстройка. Если платформа или ведущее приложение Office не поддерживает наборы требований или элементы API, указанные в элементе Requirements, надстройка не будет работать в этом ведущем приложении или на этой платформе, а также не будет отображаться в разделе "Мои надстройки".</span><span class="sxs-lookup"><span data-stu-id="e2768-p103">Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="e2768-121">Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="e2768-121">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a><span data-ttu-id="e2768-122">См. также</span><span class="sxs-lookup"><span data-stu-id="e2768-122">See also</span></span>

- [<span data-ttu-id="e2768-123">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="e2768-123">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="e2768-124">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="e2768-124">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="e2768-125">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="e2768-125">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
