---
title: Обзор API JavaScript для OneNote
description: ''
ms.date: 02/19/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 27d6770ae64a6f2259e7dbbf38b756d54fe7cec2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165322"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="8e8b5-102">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="8e8b5-102">OneNote JavaScript API overview</span></span>

<span data-ttu-id="8e8b5-103">Надстройка OneNote взаимодействует с объектами в OneNote в Интернете с помощью API JavaScript для Office, включающего две объектных модели JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8e8b5-103">A OneNote add-in interacts with objects in OneNote on the web by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="8e8b5-104">**API JavaScript для OneNote**. Появившийся в Office 2016 [API JavaScript для OneNote](/javascript/api/onenote) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к объектам OneNote в Интернете.</span><span class="sxs-lookup"><span data-stu-id="8e8b5-104">**OneNote JavaScript API**: Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span> 

* <span data-ttu-id="8e8b5-105">**Общие API**. Появившиеся в Office 2013 [общие API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="8e8b5-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="8e8b5-106">В этом разделе рассматривается API JavaScript для OneNote, используемый для разработки большинства функций в надстройках и предназначенный для OneNote в Интернете.</span><span class="sxs-lookup"><span data-stu-id="8e8b5-106">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web.</span></span> <span data-ttu-id="8e8b5-107">Сведения об общем API см. в статье [Общая объектная модель API JavaScript](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="8e8b5-107">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="8e8b5-108">Сведения о концепциях, связанных с программированием</span><span class="sxs-lookup"><span data-stu-id="8e8b5-108">Learn programming concepts</span></span>

<span data-ttu-id="8e8b5-109">Сведения о важных концепциях программирования см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="8e8b5-109">See the following articles for information about important programming concepts:</span></span>

- [<span data-ttu-id="8e8b5-110">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="8e8b5-110">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)

- [<span data-ttu-id="8e8b5-111">Работа с содержимым страницы в OneNote</span><span class="sxs-lookup"><span data-stu-id="8e8b5-111">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="8e8b5-112">Сведения о возможностях API</span><span class="sxs-lookup"><span data-stu-id="8e8b5-112">Learn about API capabilities</span></span>

<span data-ttu-id="8e8b5-113">Чтобы непосредственно использовать API JavaScript для OneNote с целью взаимодействия с содержимым в OneNote в Интернете, выполните [краткие инструкции по началу работы с надстройкой OneNote](../../quickstarts/onenote-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="8e8b5-113">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span> 

<span data-ttu-id="8e8b5-114">Дополнительные сведения об объектной модели API JavaScript для OneNote см. в [справочной документации по API JavaScript для OneNote](/javascript/api/onenote).</span><span class="sxs-lookup"><span data-stu-id="8e8b5-114">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="8e8b5-115">См. также</span><span class="sxs-lookup"><span data-stu-id="8e8b5-115">See also</span></span>

- [<span data-ttu-id="8e8b5-116">Документация по надстройкам OneNote</span><span class="sxs-lookup"><span data-stu-id="8e8b5-116">OneNote add-ins documentation</span></span>](../../onenote/index.md)
- [<span data-ttu-id="8e8b5-117">Обзор надстроек OneNote</span><span class="sxs-lookup"><span data-stu-id="8e8b5-117">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="8e8b5-118">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="8e8b5-118">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="8e8b5-119">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="8e8b5-119">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)

