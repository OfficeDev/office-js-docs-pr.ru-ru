---
title: Наборы обязательных элементов для приведения изображений
description: Поддержка наборов требований для приведения изображений с надстройками Office в Excel, PowerPoint и Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 7140099757c6e4b5ad405723d5fed95fded6d919
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293550"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="bd44a-103">Наборы обязательных элементов для приведения изображений</span><span class="sxs-lookup"><span data-stu-id="bd44a-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="bd44a-104">Наборы требований — это именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="bd44a-104">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="bd44a-105">Надстройки Office используют наборы требований, указанные в манифесте, или используют проверку среды выполнения, чтобы определить, поддерживает ли приложение Office API, необходимые надстройке.</span><span class="sxs-lookup"><span data-stu-id="bd44a-105">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="bd44a-106">Более подробную информацию можно узнать в статье [версии Office и наборах требований](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="bd44a-106">For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="bd44a-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="bd44a-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="bd44a-108">Использовать imagecoercion 1,1 обеспечивает преобразование в Image ( `Office.CoercionType.Image` ) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода.</span><span class="sxs-lookup"><span data-stu-id="bd44a-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="bd44a-109">Поддерживаются следующие приложения:</span><span class="sxs-lookup"><span data-stu-id="bd44a-109">The following applications are supported:</span></span>

- <span data-ttu-id="bd44a-110">Excel 2013 и более поздних версий в Windows</span><span class="sxs-lookup"><span data-stu-id="bd44a-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="bd44a-111">Excel 2016 и более поздних версий на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="bd44a-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="bd44a-112">Excel на iPad</span><span class="sxs-lookup"><span data-stu-id="bd44a-112">Excel on iPad</span></span>
- <span data-ttu-id="bd44a-113">OneNote в Интернете</span><span class="sxs-lookup"><span data-stu-id="bd44a-113">OneNote on the web</span></span>
- <span data-ttu-id="bd44a-114">PowerPoint 2013 и более поздних версий в Windows</span><span class="sxs-lookup"><span data-stu-id="bd44a-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="bd44a-115">PowerPoint 2016 и более поздних версий на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="bd44a-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="bd44a-116">PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="bd44a-116">PowerPoint on the web</span></span>
- <span data-ttu-id="bd44a-117">PowerPoint на iPad</span><span class="sxs-lookup"><span data-stu-id="bd44a-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="bd44a-118">Word 2013 и более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="bd44a-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="bd44a-119">Word 2016 и более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="bd44a-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="bd44a-120">Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="bd44a-120">Word on the web</span></span>
- <span data-ttu-id="bd44a-121">Word для iPad</span><span class="sxs-lookup"><span data-stu-id="bd44a-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="bd44a-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="bd44a-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="bd44a-123">Использовать imagecoercion 1,2 обеспечивает преобразование в формат SVG ( `Office.CoercionType.XmlSvg` ) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода.</span><span class="sxs-lookup"><span data-stu-id="bd44a-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="bd44a-124">Поддерживаются следующие приложения:</span><span class="sxs-lookup"><span data-stu-id="bd44a-124">The following applications are supported:</span></span>

- <span data-ttu-id="bd44a-125">Excel в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bd44a-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="bd44a-126">Excel на Mac (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bd44a-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="bd44a-127">PowerPoint в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bd44a-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="bd44a-128">PowerPoint на Mac (с подключением к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bd44a-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="bd44a-129">PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="bd44a-129">PowerPoint on the web</span></span>
- <span data-ttu-id="bd44a-130">Word в Windows (подключены к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bd44a-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="bd44a-131">Word на Mac (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bd44a-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="bd44a-132">Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="bd44a-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="bd44a-133">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="bd44a-133">Office Common API requirement sets</span></span>

<span data-ttu-id="bd44a-134">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="bd44a-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="bd44a-135">См. также</span><span class="sxs-lookup"><span data-stu-id="bd44a-135">See also</span></span>

- [<span data-ttu-id="bd44a-136">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="bd44a-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="bd44a-137">Указание приложений Office и требований к API</span><span class="sxs-lookup"><span data-stu-id="bd44a-137">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="bd44a-138">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bd44a-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
