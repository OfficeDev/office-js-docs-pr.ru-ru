---
title: Обзор API JavaScript для Word
description: ''
ms.date: 07/05/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: fbc9e8293642d1ab8edf32d568a5dab7ef77a8f0
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575627"
---
# <a name="word-javascript-api-overview"></a><span data-ttu-id="c59f2-102">Обзор API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="c59f2-102">Word JavaScript API overview</span></span>

<span data-ttu-id="c59f2-103">Надстройка Word взаимодействует с объектами в Word с помощью API JavaScript для Office, включающего две объектных модели JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c59f2-103">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="c59f2-104">**API JavaScript для Word**. Появившийся в Office 2016 [API JavaScript для Word](/javascript/api/word) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к объектам и метаданным в документе Word.</span><span class="sxs-lookup"><span data-stu-id="c59f2-104">**Word JavaScript API**: Introduced with Office 2016, the [Word JavaScript API](/javascript/api/word) provides strongly-typed objects that you can use to access objects and metadata in a Word document.</span></span> 

* <span data-ttu-id="c59f2-105">**Общие API**. Появившиеся в Office 2013 [общие API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="c59f2-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="c59f2-106">В этом разделе рассматривается API JavaScript для Word, используемый для разработки большинства функций в надстройках и предназначенный для Word в Интернете, Word 2016 или более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="c59f2-106">This section of the documentation focuses on the Word JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Word on the web or Word 2016 or later.</span></span> <span data-ttu-id="c59f2-107">Сведения об общем API см. в статье [API JavaScript для Office](../javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="c59f2-107">For more information about the distinction between host-specific APIs and Common APIs, see [JavaScript API for Office](../javascript-api-for-office.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="c59f2-108">Сведения о концепциях, связанных с программированием</span><span class="sxs-lookup"><span data-stu-id="c59f2-108">Learn programming concepts</span></span>

<span data-ttu-id="c59f2-109">Сведения о важных концепциях программирования см. в статье [Основные концепции программирования с помощью API JavaScript для Word](../../word/word-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="c59f2-109">See [Fundamental programming concepts with the Word JavaScript API](../../word/word-add-ins-core-concepts.md) for information about important programming concepts.</span></span>
 
## <a name="learn-about-api-capabilities"></a><span data-ttu-id="c59f2-110">Сведения о возможностях API</span><span class="sxs-lookup"><span data-stu-id="c59f2-110">Learn about API capabilities</span></span>

<span data-ttu-id="c59f2-111">Используйте другие статьи в этом разделе, чтобы узнать, как [получить весь документ из надстройки](../../word/get-the-whole-document-from-an-add-in-for-word.md), [воспользоваться параметрами поиска, чтобы найти текст в надстройке Word,](../../word/search-option-guidance.md) и т. д.</span><span class="sxs-lookup"><span data-stu-id="c59f2-111">Use other articles in this section of the documentation to learn how to [get the whole document from an add-in](../../word/get-the-whole-document-from-an-add-in-for-word.md), [use search options to find text in your Word add-in](../../word/search-option-guidance.md), and more.</span></span> <span data-ttu-id="c59f2-112">Полный список доступных статей см. в оглавлении.</span><span class="sxs-lookup"><span data-stu-id="c59f2-112">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="c59f2-113">Чтобы получить практический опыт доступа к объектам в Word с помощью API JavaScript для Word, выполните инструкции из [руководства по надстройкам Word](../../tutorials/word-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="c59f2-113">For hands-on experience using the Word JavaScript API to access objects in Word, complete the [Word add-in tutorial](../../tutorials/word-tutorial.md).</span></span> 

<span data-ttu-id="c59f2-114">Дополнительные сведения об объектной модели API JavaScript для Word см. в [справочной документации по API JavaScript для Word](/javascript/api/word).</span><span class="sxs-lookup"><span data-stu-id="c59f2-114">For detailed information about the Word JavaScript API, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="c59f2-115">Опробуйте примеры кода в Script Lab</span><span class="sxs-lookup"><span data-stu-id="c59f2-115">Try out code samples in Script Lab</span></span>

<span data-ttu-id="c59f2-116">Используйте [Script Lab](../../overview/explore-with-script-lab.md), чтобы быстро начать работу с коллекцией встроенных примеров, демонстрирующих выполнение задач с помощью API.</span><span class="sxs-lookup"><span data-stu-id="c59f2-116">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="c59f2-117">Вы можете выполнять примеры в Script Lab, чтобы сразу увидеть результат в области задач или документе, изучать примеры, чтобы понять принципы действия API, и даже использовать примеры для создания собственных надстроек.</span><span class="sxs-lookup"><span data-stu-id="c59f2-117">You can run the samples in Script Lab to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="c59f2-118">См. также</span><span class="sxs-lookup"><span data-stu-id="c59f2-118">See also</span></span>

- [<span data-ttu-id="c59f2-119">Документация по надстройкам Word</span><span class="sxs-lookup"><span data-stu-id="c59f2-119">Word add-ins documentation</span></span>](../../word/index.md)
- [<span data-ttu-id="c59f2-120">Обзор надстроек Word</span><span class="sxs-lookup"><span data-stu-id="c59f2-120">Word add-ins overview</span></span>](../../word/word-add-ins-programming-overview.md)
- [<span data-ttu-id="c59f2-121">Справочник по API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="c59f2-121">Word JavaScript API reference</span></span>](/javascript/api/word)
- [<span data-ttu-id="c59f2-122">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="c59f2-122">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="c59f2-123">Открытые спецификации API</span><span class="sxs-lookup"><span data-stu-id="c59f2-123">API open specifications</span></span>](../openspec/openspec.md)
