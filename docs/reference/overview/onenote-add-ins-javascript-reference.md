---
title: Обзор API JavaScript для OneNote
description: ''
ms.date: 02/19/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 8e97b0ac34e02ea64a1cb944be9c113bd37a9717
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325187"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="7ca6c-102">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="7ca6c-102">OneNote JavaScript API overview</span></span>

<span data-ttu-id="7ca6c-103">Надстройка OneNote взаимодействует с объектами в OneNote в Интернете с помощью API JavaScript для Office, включающего две объектных модели JavaScript.</span><span class="sxs-lookup"><span data-stu-id="7ca6c-103">A OneNote add-in interacts with objects in OneNote on the web by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="7ca6c-104">**API JavaScript для OneNote**. Появившийся в Office 2016 [API JavaScript для OneNote](/javascript/api/onenote) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к объектам OneNote в Интернете.</span><span class="sxs-lookup"><span data-stu-id="7ca6c-104">**OneNote JavaScript API**: Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span> 

* <span data-ttu-id="7ca6c-105">**Общие API**. Появившиеся в Office 2013 [общие API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="7ca6c-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="7ca6c-106">В этом разделе рассматривается API JavaScript для OneNote, используемый для разработки большинства функций в надстройках и предназначенный для OneNote в Интернете.</span><span class="sxs-lookup"><span data-stu-id="7ca6c-106">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web.</span></span> <span data-ttu-id="7ca6c-107">Сведения об общем API см. в статье [Общая объектная модель API JavaScript](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="7ca6c-107">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="7ca6c-108">Сведения о концепциях, связанных с программированием</span><span class="sxs-lookup"><span data-stu-id="7ca6c-108">Learn programming concepts</span></span>

<span data-ttu-id="7ca6c-109">Сведения о важных концепциях программирования см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="7ca6c-109">See the following articles for information about important programming concepts:</span></span>

- [<span data-ttu-id="7ca6c-110">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="7ca6c-110">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)

- [<span data-ttu-id="7ca6c-111">Работа с содержимым страницы в OneNote</span><span class="sxs-lookup"><span data-stu-id="7ca6c-111">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="7ca6c-112">Сведения о возможностях API</span><span class="sxs-lookup"><span data-stu-id="7ca6c-112">Learn about API capabilities</span></span>

<span data-ttu-id="7ca6c-113">Чтобы непосредственно использовать API JavaScript для OneNote с целью взаимодействия с содержимым в OneNote в Интернете, выполните [краткие инструкции по началу работы с надстройкой OneNote](../../quickstarts/onenote-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="7ca6c-113">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span> 

<span data-ttu-id="7ca6c-114">Дополнительные сведения об объектной модели API JavaScript для OneNote см. в [справочной документации по API JavaScript для OneNote](/javascript/api/onenote).</span><span class="sxs-lookup"><span data-stu-id="7ca6c-114">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="7ca6c-115">См. также</span><span class="sxs-lookup"><span data-stu-id="7ca6c-115">See also</span></span>

- [<span data-ttu-id="7ca6c-116">Документация по надстройкам OneNote</span><span class="sxs-lookup"><span data-stu-id="7ca6c-116">OneNote add-ins documentation</span></span>](../../onenote/index.md)
- [<span data-ttu-id="7ca6c-117">Обзор надстроек OneNote</span><span class="sxs-lookup"><span data-stu-id="7ca6c-117">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="7ca6c-118">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="7ca6c-118">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="7ca6c-119">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="7ca6c-119">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)

