---
title: Наборы обязательных элементов API удостоверений
description: Сведения о наборе требований API удостоверений для надстроек Office.
ms.date: 07/30/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c2c6ea449cef08248a9ba79051b7c0c5f9baa600
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293543"
---
# <a name="identity-api-requirement-sets"></a><span data-ttu-id="26a37-103">Наборы обязательных элементов API удостоверений</span><span class="sxs-lookup"><span data-stu-id="26a37-103">Identity API requirement sets</span></span>

<span data-ttu-id="26a37-104">Наборы требований — это именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="26a37-104">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="26a37-105">Надстройки Office используют наборы требований, указанные в манифесте, или используют проверку среды выполнения, чтобы определить, поддерживает ли приложение Office API, необходимые надстройке.</span><span class="sxs-lookup"><span data-stu-id="26a37-105">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="26a37-106">Более подробную информацию можно узнать в статье [версии Office и наборах требований](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="26a37-106">For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="26a37-107">Надстройки Office работают в нескольких версиях Office.</span><span class="sxs-lookup"><span data-stu-id="26a37-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="26a37-108">В следующей таблице перечислены наборы обязательных элементов API удостоверений, клиентские приложения Office, которые поддерживают этот набор требований, а также номера сборок или версий приложений Office.</span><span class="sxs-lookup"><span data-stu-id="26a37-108">The following table lists the Identity API requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="26a37-109">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="26a37-109">Requirement set</span></span>  | <span data-ttu-id="26a37-110">Office 2013 или более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="26a37-110">Office 2013 or later on Windows</span></span><br><span data-ttu-id="26a37-111">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="26a37-111">(one-time purchase)</span></span> | <span data-ttu-id="26a37-112">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="26a37-112">Office on Windows</span></span><br><span data-ttu-id="26a37-113">(подключено к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="26a37-113">(connected to a Microsoft 365 subscription)</span></span> |  <span data-ttu-id="26a37-114">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="26a37-114">Office on iPad</span></span><br><span data-ttu-id="26a37-115">(подключено к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="26a37-115">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="26a37-116">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="26a37-116">Office on Mac</span></span><br><span data-ttu-id="26a37-117">(подключено к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="26a37-117">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="26a37-118">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="26a37-118">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="26a37-119">IdentityAPI 1,3</span><span class="sxs-lookup"><span data-stu-id="26a37-119">IdentityAPI 1.3</span></span>  | <span data-ttu-id="26a37-120">Недоступно</span><span class="sxs-lookup"><span data-stu-id="26a37-120">N/A</span></span> | <span data-ttu-id="26a37-121">2008 (сборка 13127,20000) или более поздняя</span><span class="sxs-lookup"><span data-stu-id="26a37-121">2008 (build 13127.20000) or later</span></span> | <span data-ttu-id="26a37-122">Скоро</span><span class="sxs-lookup"><span data-stu-id="26a37-122">Coming soon</span></span> | <span data-ttu-id="26a37-123">16,40 или более поздняя версия</span><span class="sxs-lookup"><span data-stu-id="26a37-123">16.40 or later</span></span> | <span data-ttu-id="26a37-124">Август, 2020 \*</span><span class="sxs-lookup"><span data-stu-id="26a37-124">August, 2020\*</span></span> |

> <span data-ttu-id="26a37-125">\* Изначально набор требований поддерживается в Office в Интернете только для документов, открытых из SharePoint Online и OneDrive.com.</span><span class="sxs-lookup"><span data-stu-id="26a37-125">\* Initially, the requirement set is supported in Office on the web only for documents that are opened from SharePoint Online and OneDrive.com.</span></span> <span data-ttu-id="26a37-126">Поддержка других документов будет поступать в Office в Интернете позже в 2020.</span><span class="sxs-lookup"><span data-stu-id="26a37-126">Support for other documents will come to Office on the web later in 2020.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="26a37-127">Номера версий и сборок Office</span><span class="sxs-lookup"><span data-stu-id="26a37-127">Office versions and build numbers</span></span>

<span data-ttu-id="26a37-128">Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:</span><span class="sxs-lookup"><span data-stu-id="26a37-128">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="26a37-129">Обзор Office Online Server</span><span class="sxs-lookup"><span data-stu-id="26a37-129">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="26a37-130">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="26a37-130">Office Common API requirement sets</span></span>

<span data-ttu-id="26a37-131">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="26a37-131">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="identityapi-preview"></a><span data-ttu-id="26a37-132">Предварительный просмотр IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="26a37-132">IdentityAPI Preview</span></span>

<span data-ttu-id="26a37-133">Подробнее об этом API можно узнать в версии, использующей обещания в [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) , или в версии, использующей функции обратного вызова по адресу [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).</span><span class="sxs-lookup"><span data-stu-id="26a37-133">For details about this API, see either the version that uses Promises at [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) or the version that uses callbacks at [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).</span></span>

## <a name="see-also"></a><span data-ttu-id="26a37-134">См. также</span><span class="sxs-lookup"><span data-stu-id="26a37-134">See also</span></span>

- [<span data-ttu-id="26a37-135">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="26a37-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="26a37-136">Указание приложений Office и требований к API</span><span class="sxs-lookup"><span data-stu-id="26a37-136">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="26a37-137">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="26a37-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
