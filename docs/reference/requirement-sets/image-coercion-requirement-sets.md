---
title: Наборы требований к приведению изображений
description: Поддержка наборов требований для приведения изображений с надстройками Office в Excel, PowerPoint и Word.
ms.date: 07/11/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: bffe6c074d9e0734299d0087f2488524875931ed
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940855"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="374bf-103">Наборы требований к приведению изображений</span><span class="sxs-lookup"><span data-stu-id="374bf-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="374bf-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="374bf-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="374bf-107">Надстройки Office работают в нескольких версиях Office.</span><span class="sxs-lookup"><span data-stu-id="374bf-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="374bf-108">В приведенной ниже таблице перечислены наборы требований к приведению изображений, ведущие приложения Office, которые поддерживают этот набор требований, а также номера сборок или версий приложений Office.</span><span class="sxs-lookup"><span data-stu-id="374bf-108">The following table lists the Image Coercion requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="374bf-109">Использовать imagecoercion 1,1</span><span class="sxs-lookup"><span data-stu-id="374bf-109">ImageCoercion 1.1</span></span>

<span data-ttu-id="374bf-110">Использовать imagecoercion 1,1 обеспечивает преобразование в Image (`Office.CoercionType.Image`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода.</span><span class="sxs-lookup"><span data-stu-id="374bf-110">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="374bf-111">Поддерживаются следующие узлы:</span><span class="sxs-lookup"><span data-stu-id="374bf-111">The following hosts are supported:</span></span>

- <span data-ttu-id="374bf-112">Excel 2013 и более поздних версий в Windows</span><span class="sxs-lookup"><span data-stu-id="374bf-112">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="374bf-113">Excel 2016 и более поздних версий на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="374bf-113">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="374bf-114">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="374bf-114">Excel on the web</span></span>
- <span data-ttu-id="374bf-115">Excel на iPad</span><span class="sxs-lookup"><span data-stu-id="374bf-115">Excel on iPad</span></span>
- <span data-ttu-id="374bf-116">OneNote в Интернете</span><span class="sxs-lookup"><span data-stu-id="374bf-116">OneNote on the web</span></span>
- <span data-ttu-id="374bf-117">PowerPoint 2013 и более поздних версий в Windows</span><span class="sxs-lookup"><span data-stu-id="374bf-117">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="374bf-118">PowerPoint 2016 и более поздних версий на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="374bf-118">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="374bf-119">PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="374bf-119">PowerPoint on the web</span></span>
- <span data-ttu-id="374bf-120">PowerPoint на iPad</span><span class="sxs-lookup"><span data-stu-id="374bf-120">PowerPoint on iPad</span></span>
- <span data-ttu-id="374bf-121">Word 2013 и более поздние версии для Windows</span><span class="sxs-lookup"><span data-stu-id="374bf-121">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="374bf-122">Word 2016 и более поздние версии на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="374bf-122">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="374bf-123">Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="374bf-123">Word on the web</span></span>
- <span data-ttu-id="374bf-124">Word на iPad</span><span class="sxs-lookup"><span data-stu-id="374bf-124">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="374bf-125">Использовать imagecoercion 1,2</span><span class="sxs-lookup"><span data-stu-id="374bf-125">ImageCoercion 1.2</span></span>

<span data-ttu-id="374bf-126">Использовать imagecoercion 1,2 обеспечивает преобразование в формат SVG (`Office.CoercionType.XmlSvg`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода.</span><span class="sxs-lookup"><span data-stu-id="374bf-126">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="374bf-127">Поддерживаются следующие узлы:</span><span class="sxs-lookup"><span data-stu-id="374bf-127">The following hosts are supported:</span></span>

- <span data-ttu-id="374bf-128">Excel в Windows (подключен к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="374bf-128">Excel on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="374bf-129">Excel на Mac (подключен к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="374bf-129">Excel on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="374bf-130">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="374bf-130">Excel on the web</span></span>
- <span data-ttu-id="374bf-131">PowerPoint в Windows (подключено к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="374bf-131">PowerPoint on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="374bf-132">PowerPoint на Mac (с подключением к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="374bf-132">PowerPoint on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="374bf-133">PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="374bf-133">PowerPoint on the web</span></span>
- <span data-ttu-id="374bf-134">Word в Windows (подключен к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="374bf-134">Word on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="374bf-135">Word на Mac (подключен к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="374bf-135">Word on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="374bf-136">Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="374bf-136">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="374bf-137">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="374bf-137">Office Common API requirement sets</span></span>

<span data-ttu-id="374bf-138">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="374bf-138">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="374bf-139">См. также</span><span class="sxs-lookup"><span data-stu-id="374bf-139">See also</span></span>

- [<span data-ttu-id="374bf-140">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="374bf-140">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="374bf-141">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="374bf-141">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="374bf-142">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="374bf-142">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
