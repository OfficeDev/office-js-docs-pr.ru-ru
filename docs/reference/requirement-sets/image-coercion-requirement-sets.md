---
title: Наборы обязательных элементов для приведения изображений
description: Поддержка наборов требований для приведения изображений с надстройками Office в Excel, PowerPoint и Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: ccc65f3c38e8ddc4bea88d897e6abda73aa61e64
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717476"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="0cf95-103">Наборы обязательных элементов для приведения изображений</span><span class="sxs-lookup"><span data-stu-id="0cf95-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="0cf95-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="0cf95-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="0cf95-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="0cf95-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="0cf95-108">Использовать imagecoercion 1,1 обеспечивает преобразование в Image (`Office.CoercionType.Image`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода.</span><span class="sxs-lookup"><span data-stu-id="0cf95-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="0cf95-109">Поддерживаются следующие узлы:</span><span class="sxs-lookup"><span data-stu-id="0cf95-109">The following hosts are supported:</span></span>

- <span data-ttu-id="0cf95-110">Excel 2013 и более поздних версий в Windows</span><span class="sxs-lookup"><span data-stu-id="0cf95-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="0cf95-111">Excel 2016 и более поздних версий на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="0cf95-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="0cf95-112">Excel на iPad</span><span class="sxs-lookup"><span data-stu-id="0cf95-112">Excel on iPad</span></span>
- <span data-ttu-id="0cf95-113">OneNote в Интернете</span><span class="sxs-lookup"><span data-stu-id="0cf95-113">OneNote on the web</span></span>
- <span data-ttu-id="0cf95-114">PowerPoint 2013 и более поздних версий в Windows</span><span class="sxs-lookup"><span data-stu-id="0cf95-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="0cf95-115">PowerPoint 2016 и более поздних версий на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="0cf95-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="0cf95-116">PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="0cf95-116">PowerPoint on the web</span></span>
- <span data-ttu-id="0cf95-117">PowerPoint на iPad</span><span class="sxs-lookup"><span data-stu-id="0cf95-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="0cf95-118">Word 2013 и более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="0cf95-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="0cf95-119">Word 2016 и более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="0cf95-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="0cf95-120">Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="0cf95-120">Word on the web</span></span>
- <span data-ttu-id="0cf95-121">Word для iPad</span><span class="sxs-lookup"><span data-stu-id="0cf95-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="0cf95-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="0cf95-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="0cf95-123">Использовать imagecoercion 1,2 обеспечивает преобразование в формат SVG (`Office.CoercionType.XmlSvg`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода.</span><span class="sxs-lookup"><span data-stu-id="0cf95-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="0cf95-124">Поддерживаются следующие узлы:</span><span class="sxs-lookup"><span data-stu-id="0cf95-124">The following hosts are supported:</span></span>

- <span data-ttu-id="0cf95-125">Excel в Windows (подключен к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="0cf95-125">Excel on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0cf95-126">Excel на Mac (подключен к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="0cf95-126">Excel on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0cf95-127">PowerPoint в Windows (подключено к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="0cf95-127">PowerPoint on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0cf95-128">PowerPoint на Mac (с подключением к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="0cf95-128">PowerPoint on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0cf95-129">PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="0cf95-129">PowerPoint on the web</span></span>
- <span data-ttu-id="0cf95-130">Word в Windows (подключен к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="0cf95-130">Word on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0cf95-131">Word на Mac (подключен к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="0cf95-131">Word on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0cf95-132">Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="0cf95-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="0cf95-133">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="0cf95-133">Office Common API requirement sets</span></span>

<span data-ttu-id="0cf95-134">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="0cf95-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0cf95-135">См. также</span><span class="sxs-lookup"><span data-stu-id="0cf95-135">See also</span></span>

- [<span data-ttu-id="0cf95-136">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="0cf95-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="0cf95-137">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0cf95-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="0cf95-138">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0cf95-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
