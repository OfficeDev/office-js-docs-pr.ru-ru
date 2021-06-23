---
title: Работа с содержимым страницы в OneNote
description: Узнайте, как работать с OneNote контентом страницы с помощью API JavaScript.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 9c4744f1121bbc5e28783940a946727275b806f2
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076821"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="089e7-103">Работа с содержимым страницы в OneNote</span><span class="sxs-lookup"><span data-stu-id="089e7-103">Work with OneNote page content</span></span>

<span data-ttu-id="089e7-104">В API JavaScript для надстроек OneNote содержимое страницы представлено указанной ниже объектной моделью.</span><span class="sxs-lookup"><span data-stu-id="089e7-104">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![OneNote схема объектной модели страницы.](../images/one-note-om-page.png)

- <span data-ttu-id="089e7-106">Объект Page содержит коллекцию объектов PageContent.</span><span class="sxs-lookup"><span data-stu-id="089e7-106">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="089e7-107">Объект PageContent содержит контент типов Outline, Image или Other.</span><span class="sxs-lookup"><span data-stu-id="089e7-107">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="089e7-108">Объект Outline содержит коллекцию объектов Paragraph.</span><span class="sxs-lookup"><span data-stu-id="089e7-108">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="089e7-109">Объект Paragraph содержит контент типов RichText, Image, Table или Other.</span><span class="sxs-lookup"><span data-stu-id="089e7-109">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="089e7-110">Чтобы создать пустую страницу OneNote, воспользуйтесь одним из указанных ниже методов.</span><span class="sxs-lookup"><span data-stu-id="089e7-110">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="089e7-111">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="089e7-111">Section.addPage</span></span>](/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="089e7-112">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="089e7-112">Page.insertPageAsSibling</span></span>](/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="089e7-113">Затем используйте методы в указанных ниже объектах для работы с содержимым страницы, например `Page.addOutline` и `Outline.appendHtml`.</span><span class="sxs-lookup"><span data-stu-id="089e7-113">Then use methods in the following objects to work with the page content, such as `Page.addOutline` and `Outline.appendHtml`.</span></span>

- [<span data-ttu-id="089e7-114">Страница</span><span class="sxs-lookup"><span data-stu-id="089e7-114">Page</span></span>](/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="089e7-115">Outline</span><span class="sxs-lookup"><span data-stu-id="089e7-115">Outline</span></span>](/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="089e7-116">Paragraph</span><span class="sxs-lookup"><span data-stu-id="089e7-116">Paragraph</span></span>](/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="089e7-p101">Для представления содержимого и структуры страницы OneNote используется HTML. Для создания или обновления содержимого страницы поддерживается только подмножество HTML, как описано ниже.</span><span class="sxs-lookup"><span data-stu-id="089e7-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="089e7-119">Поддерживаемые элементы HTML</span><span class="sxs-lookup"><span data-stu-id="089e7-119">Supported HTML</span></span>

<span data-ttu-id="089e7-120">Для создания и обновления содержимого страницы в API JavaScript для надстроек OneNote используются указанные ниже элементы HTML.</span><span class="sxs-lookup"><span data-stu-id="089e7-120">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="089e7-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="089e7-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span>
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="089e7-122">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="089e7-122">`<ul>`, `<ol>`, `<li>`</span></span>
- <span data-ttu-id="089e7-123">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="089e7-123">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="089e7-124">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="089e7-124">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="089e7-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="089e7-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

> [!NOTE]
> <span data-ttu-id="089e7-126">Импорт HTML в OneNote консолидирует пробелы.</span><span class="sxs-lookup"><span data-stu-id="089e7-126">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="089e7-127">Полученное в результате содержимое вставляется в одну структуру.</span><span class="sxs-lookup"><span data-stu-id="089e7-127">The resulting content is pasted into one outline.</span></span>

<span data-ttu-id="089e7-128">Приложение OneNote пытается наилучшим образом преобразовать HTML в содержимое страницы, обеспечивая безопасность для пользователей.</span><span class="sxs-lookup"><span data-stu-id="089e7-128">OneNote does its best to translate HTML into page content while ensuring security for users.</span></span> <span data-ttu-id="089e7-129">Так как стандарты HTML и CSS не полностью соответствуют модели содержимого OneNote, будут иметься различия во внешнем виде, особенно при использовании стилей CSS.</span><span class="sxs-lookup"><span data-stu-id="089e7-129">HTML and CSS standards do not exactly match OneNote's content model, so there will be differences in appearances, particularly with CSS stylings.</span></span> <span data-ttu-id="089e7-130">Рекомендуется использовать объекты JavaScript, если требуется определенное форматирование.</span><span class="sxs-lookup"><span data-stu-id="089e7-130">We recommend using the JavaScript objects if specific formatting is needed.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="089e7-131">Доступ к содержимому страницы</span><span class="sxs-lookup"><span data-stu-id="089e7-131">Accessing page contents</span></span>

<span data-ttu-id="089e7-p104">Через  доступ можно получить только к `Page#load`. Чтобы изменить активную страницу, вызовите команду `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="089e7-p104">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="089e7-134">Метаданные, например "Название", можно запросить для любой страницы.</span><span class="sxs-lookup"><span data-stu-id="089e7-134">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="089e7-135">См. также</span><span class="sxs-lookup"><span data-stu-id="089e7-135">See also</span></span>

- [<span data-ttu-id="089e7-136">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="089e7-136">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="089e7-137">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="089e7-137">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="089e7-138">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="089e7-138">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="089e7-139">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="089e7-139">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
