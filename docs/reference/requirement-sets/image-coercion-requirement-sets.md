---
title: Наборы обязательных элементов для приведения изображений
description: Поддержка наборов требований к принуждению к изображениям с помощью надстройок Office в Excel, PowerPoint и Word.
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 52ce46a46580500f5a292bf898674d4798378319
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505530"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="ead18-103">Наборы обязательных элементов для приведения изображений</span><span class="sxs-lookup"><span data-stu-id="ead18-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="ead18-p101">Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="ead18-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="ead18-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="ead18-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="ead18-108">ImageCoercion 1.1 позволяет преобразования в изображение () при записи `Office.CoercionType.Image` данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода.</span><span class="sxs-lookup"><span data-stu-id="ead18-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="ead18-109">Поддерживаются следующие приложения:</span><span class="sxs-lookup"><span data-stu-id="ead18-109">The following applications are supported:</span></span>

- <span data-ttu-id="ead18-110">Excel 2013 и более поздние версии Windows</span><span class="sxs-lookup"><span data-stu-id="ead18-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="ead18-111">Excel 2016 и более поздний mac</span><span class="sxs-lookup"><span data-stu-id="ead18-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="ead18-112">Excel на iPad</span><span class="sxs-lookup"><span data-stu-id="ead18-112">Excel on iPad</span></span>
- <span data-ttu-id="ead18-113">OneNote в Интернете</span><span class="sxs-lookup"><span data-stu-id="ead18-113">OneNote on the web</span></span>
- <span data-ttu-id="ead18-114">PowerPoint 2013 и более поздние версии Windows</span><span class="sxs-lookup"><span data-stu-id="ead18-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="ead18-115">PowerPoint 2016 и более поздний mac</span><span class="sxs-lookup"><span data-stu-id="ead18-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="ead18-116">PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="ead18-116">PowerPoint on the web</span></span>
- <span data-ttu-id="ead18-117">PowerPoint на iPad</span><span class="sxs-lookup"><span data-stu-id="ead18-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="ead18-118">Word 2013 и более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="ead18-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="ead18-119">Word 2016 и более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="ead18-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="ead18-120">Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="ead18-120">Word on the web</span></span>
- <span data-ttu-id="ead18-121">Word для iPad</span><span class="sxs-lookup"><span data-stu-id="ead18-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="ead18-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="ead18-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="ead18-123">ImageCoercion 1.2 позволяет преобразования в формат SVG () при записи данных `Office.CoercionType.XmlSvg` с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода.</span><span class="sxs-lookup"><span data-stu-id="ead18-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="ead18-124">Поддерживаются следующие приложения:</span><span class="sxs-lookup"><span data-stu-id="ead18-124">The following applications are supported:</span></span>

- <span data-ttu-id="ead18-125">Excel на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ead18-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="ead18-126">Excel на Mac (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ead18-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="ead18-127">PowerPoint на Windows (подключена к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ead18-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="ead18-128">PowerPoint на Mac (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ead18-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="ead18-129">PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="ead18-129">PowerPoint on the web</span></span>
- <span data-ttu-id="ead18-130">Word on Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ead18-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="ead18-131">Word на Mac (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ead18-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="ead18-132">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="ead18-132">Office Common API requirement sets</span></span>

<span data-ttu-id="ead18-133">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="ead18-133">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="ead18-134">См. также</span><span class="sxs-lookup"><span data-stu-id="ead18-134">See also</span></span>

- [<span data-ttu-id="ead18-135">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="ead18-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ead18-136">Указание приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="ead18-136">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="ead18-137">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="ead18-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
