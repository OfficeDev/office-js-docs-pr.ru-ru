---
title: Наборы обязательных элементов для приведения изображений
description: Поддержка наборов требований для приведения изображений с надстройками Office в Excel, PowerPoint и Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 59f6891182f47bed1b7e3b6aa69a30e941bce7cb
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094357"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="d4ff4-103">Наборы обязательных элементов для приведения изображений</span><span class="sxs-lookup"><span data-stu-id="d4ff4-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="d4ff4-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="d4ff4-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="d4ff4-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="d4ff4-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="d4ff4-108">Использовать imagecoercion 1,1 обеспечивает преобразование в Image ( `Office.CoercionType.Image` ) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода.</span><span class="sxs-lookup"><span data-stu-id="d4ff4-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="d4ff4-109">Поддерживаются следующие узлы:</span><span class="sxs-lookup"><span data-stu-id="d4ff4-109">The following hosts are supported:</span></span>

- <span data-ttu-id="d4ff4-110">Excel 2013 и более поздних версий в Windows</span><span class="sxs-lookup"><span data-stu-id="d4ff4-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="d4ff4-111">Excel 2016 и более поздних версий на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="d4ff4-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="d4ff4-112">Excel на iPad</span><span class="sxs-lookup"><span data-stu-id="d4ff4-112">Excel on iPad</span></span>
- <span data-ttu-id="d4ff4-113">OneNote в Интернете</span><span class="sxs-lookup"><span data-stu-id="d4ff4-113">OneNote on the web</span></span>
- <span data-ttu-id="d4ff4-114">PowerPoint 2013 и более поздних версий в Windows</span><span class="sxs-lookup"><span data-stu-id="d4ff4-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="d4ff4-115">PowerPoint 2016 и более поздних версий на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="d4ff4-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="d4ff4-116">PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="d4ff4-116">PowerPoint on the web</span></span>
- <span data-ttu-id="d4ff4-117">PowerPoint на iPad</span><span class="sxs-lookup"><span data-stu-id="d4ff4-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="d4ff4-118">Word 2013 и более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="d4ff4-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="d4ff4-119">Word 2016 и более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="d4ff4-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="d4ff4-120">Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="d4ff4-120">Word on the web</span></span>
- <span data-ttu-id="d4ff4-121">Word для iPad</span><span class="sxs-lookup"><span data-stu-id="d4ff4-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="d4ff4-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="d4ff4-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="d4ff4-123">Использовать imagecoercion 1,2 обеспечивает преобразование в формат SVG ( `Office.CoercionType.XmlSvg` ) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода.</span><span class="sxs-lookup"><span data-stu-id="d4ff4-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="d4ff4-124">Поддерживаются следующие узлы:</span><span class="sxs-lookup"><span data-stu-id="d4ff4-124">The following hosts are supported:</span></span>

- <span data-ttu-id="d4ff4-125">Excel в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="d4ff4-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="d4ff4-126">Excel на Mac (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="d4ff4-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="d4ff4-127">PowerPoint в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="d4ff4-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="d4ff4-128">PowerPoint на Mac (с подключением к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="d4ff4-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="d4ff4-129">PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="d4ff4-129">PowerPoint on the web</span></span>
- <span data-ttu-id="d4ff4-130">Word в Windows (подключены к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="d4ff4-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="d4ff4-131">Word на Mac (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="d4ff4-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="d4ff4-132">Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="d4ff4-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="d4ff4-133">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="d4ff4-133">Office Common API requirement sets</span></span>

<span data-ttu-id="d4ff4-134">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="d4ff4-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d4ff4-135">См. также</span><span class="sxs-lookup"><span data-stu-id="d4ff4-135">See also</span></span>

- [<span data-ttu-id="d4ff4-136">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="d4ff4-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="d4ff4-137">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d4ff4-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="d4ff4-138">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d4ff4-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
