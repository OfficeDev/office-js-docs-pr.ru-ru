---
title: Обзор API JavaScript для OneNote
description: Узнайте больше об API OneNote JavaScript
ms.date: 07/28/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: d917d71cd9d3f4fadbab91a434a177c45b54c6f2
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349114"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="2532c-103">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="2532c-103">OneNote JavaScript API overview</span></span>

<span data-ttu-id="2532c-104">Надстройка OneNote взаимодействует с объектами в OneNote в Интернете с помощью API JavaScript для Office, включающего две объектных модели JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2532c-104">A OneNote add-in interacts with objects in OneNote on the web by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="2532c-p101">**API JavaScript для OneNote** — это [API конкретных приложений](../../develop/application-specific-api-model.md) для OneNote. Впервые представленный в Office 2016, [API JavaScript для OneNote](/javascript/api/onenote) предоставляет строго типизированные объекты, которые можно использовать для доступа к объектам в OneNote для Интернета.</span><span class="sxs-lookup"><span data-stu-id="2532c-p101">**OneNote JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for OneNote. Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span>

* <span data-ttu-id="2532c-107">**Общие API**. Появившиеся в Office 2013 [общие API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="2532c-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="2532c-p102">В этом разделе документации основное внимание уделяется интерфейсу API JavaScript для OneNote, который используется для разработки большинства функций надстройки, ориентированных на OneNote в Интернете. Сведения об общем API см. в статье об [общей объектной модели API JavaScript](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="2532c-p102">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web. For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-programming-concepts"></a><span data-ttu-id="2532c-110">Сведения о понятиях, связанных с программированием</span><span class="sxs-lookup"><span data-stu-id="2532c-110">Learn programming concepts</span></span>

<span data-ttu-id="2532c-111">Сведения о важных понятиях, связанных с программированием, см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="2532c-111">See the following articles for information about important programming concepts.</span></span>

* [<span data-ttu-id="2532c-112">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="2532c-112">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="2532c-113">Работа с содержимым страницы в OneNote</span><span class="sxs-lookup"><span data-stu-id="2532c-113">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="2532c-114">Сведения о возможностях API</span><span class="sxs-lookup"><span data-stu-id="2532c-114">Learn about API capabilities</span></span>

<span data-ttu-id="2532c-115">Чтобы непосредственно использовать API JavaScript для OneNote с целью взаимодействия с содержимым в OneNote в Интернете, выполните [краткие инструкции по началу работы с надстройкой OneNote](../../quickstarts/onenote-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="2532c-115">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span>

<span data-ttu-id="2532c-116">Дополнительные сведения об объектной модели API JavaScript для OneNote см. в [справочной документации по API JavaScript для OneNote](/javascript/api/onenote).</span><span class="sxs-lookup"><span data-stu-id="2532c-116">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="2532c-117">См. также</span><span class="sxs-lookup"><span data-stu-id="2532c-117">See also</span></span>

* [<span data-ttu-id="2532c-118">Документация по надстройкам OneNote</span><span class="sxs-lookup"><span data-stu-id="2532c-118">OneNote add-ins documentation</span></span>](../../onenote/index.yml)
* [<span data-ttu-id="2532c-119">Обзор надстроек OneNote</span><span class="sxs-lookup"><span data-stu-id="2532c-119">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="2532c-120">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="2532c-120">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
* [<span data-ttu-id="2532c-121">Доступность клиентских приложений и платформ Office для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="2532c-121">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
