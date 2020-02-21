---
title: Обзор API JavaScript для Word
description: ''
ms.date: 02/19/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 90dd7c787086a67dd8607479bbc46c957192d5c3
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163971"
---
# <a name="word-javascript-api-overview"></a><span data-ttu-id="1e091-102">Обзор API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="1e091-102">Word JavaScript API overview</span></span>

<span data-ttu-id="1e091-103">Надстройка Word взаимодействует с объектами в Word с помощью API JavaScript для Office, включающего две объектных модели JavaScript.</span><span class="sxs-lookup"><span data-stu-id="1e091-103">An Word add-in interacts with objects in Word by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="1e091-104">**API JavaScript для Word**. Появившийся в Office 2016 [API JavaScript для Word](/javascript/api/word) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к объектам и метаданным в документе Word.</span><span class="sxs-lookup"><span data-stu-id="1e091-104">**Word JavaScript API**: Introduced with Office 2016, the [Word JavaScript API](/javascript/api/word) provides strongly-typed objects that you can use to access objects and metadata in a Word document.</span></span> 

* <span data-ttu-id="1e091-105">**Общие API**. Появившиеся в Office 2013 [общие API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="1e091-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="1e091-106">В этом разделе рассматривается API JavaScript для Word, используемый для разработки большинства функций в надстройках и предназначенный для Word в Интернете, Word 2016 или более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="1e091-106">This section of the documentation focuses on the Word JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Word on the web or Word 2016 or later.</span></span> <span data-ttu-id="1e091-107">Сведения об общем API см. в статье [Общая объектная модель API JavaScript](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="1e091-107">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="1e091-108">Сведения о концепциях, связанных с программированием</span><span class="sxs-lookup"><span data-stu-id="1e091-108">Learn programming concepts</span></span>

<span data-ttu-id="1e091-109">Сведения о важных концепциях программирования см. в статье [Основные концепции программирования с помощью API JavaScript для Word](../../word/word-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="1e091-109">See [Fundamental programming concepts with the Word JavaScript API](../../word/word-add-ins-core-concepts.md) for information about important programming concepts.</span></span>
 
## <a name="learn-about-api-capabilities"></a><span data-ttu-id="1e091-110">Сведения о возможностях API</span><span class="sxs-lookup"><span data-stu-id="1e091-110">Learn about API capabilities</span></span>

<span data-ttu-id="1e091-111">Используйте другие статьи в этом разделе, чтобы узнать, как [получить весь документ из надстройки](../../word/get-the-whole-document-from-an-add-in-for-word.md), [воспользоваться параметрами поиска, чтобы найти текст в надстройке Word,](../../word/search-option-guidance.md) и т. д.</span><span class="sxs-lookup"><span data-stu-id="1e091-111">Use other articles in this section of the documentation to learn how to [get the whole document from an add-in](../../word/get-the-whole-document-from-an-add-in-for-word.md), [use search options to find text in your Word add-in](../../word/search-option-guidance.md), and more.</span></span> <span data-ttu-id="1e091-112">Полный список доступных статей см. в оглавлении.</span><span class="sxs-lookup"><span data-stu-id="1e091-112">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="1e091-113">Чтобы получить практический опыт доступа к объектам в Word с помощью API JavaScript для Word, выполните инструкции из [руководства по надстройкам Word](../../tutorials/word-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="1e091-113">For hands-on experience using the Word JavaScript API to access objects in Word, complete the [Word add-in tutorial](../../tutorials/word-tutorial.md).</span></span> 

<span data-ttu-id="1e091-114">Дополнительные сведения об объектной модели API JavaScript для Word см. в [справочной документации по API JavaScript для Word](/javascript/api/word).</span><span class="sxs-lookup"><span data-stu-id="1e091-114">For detailed information about the Word JavaScript API object model, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="1e091-115">Опробуйте примеры кода в Script Lab</span><span class="sxs-lookup"><span data-stu-id="1e091-115">Try out code samples in Script Lab</span></span>

<span data-ttu-id="1e091-116">Используйте [Script Lab](../../overview/explore-with-script-lab.md), чтобы быстро начать работу с коллекцией встроенных примеров, демонстрирующих выполнение задач с помощью API.</span><span class="sxs-lookup"><span data-stu-id="1e091-116">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="1e091-117">Вы можете выполнять примеры в Script Lab, чтобы сразу увидеть результат в области задач или документе, изучать примеры, чтобы понять принципы действия API, и даже использовать примеры для создания собственных надстроек.</span><span class="sxs-lookup"><span data-stu-id="1e091-117">You can run the samples in Script Lab to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="1e091-118">См. также</span><span class="sxs-lookup"><span data-stu-id="1e091-118">See also</span></span>

- [<span data-ttu-id="1e091-119">Документация по надстройкам Word</span><span class="sxs-lookup"><span data-stu-id="1e091-119">Word add-ins documentation</span></span>](../../word/index.md)
- [<span data-ttu-id="1e091-120">Обзор надстроек Word</span><span class="sxs-lookup"><span data-stu-id="1e091-120">Word add-ins overview</span></span>](../../word/word-add-ins-programming-overview.md)
- [<span data-ttu-id="1e091-121">Справочник по API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="1e091-121">Word JavaScript API reference</span></span>](/javascript/api/word)
- [<span data-ttu-id="1e091-122">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="1e091-122">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="1e091-123">Открытые спецификации API</span><span class="sxs-lookup"><span data-stu-id="1e091-123">API open specifications</span></span>](../openspec/openspec.md)
