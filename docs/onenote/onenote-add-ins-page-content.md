---
title: Работа с содержимым страницы в OneNote
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: f60cdee7eb549acc0f2c84a1aa9acea7fe77274a
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872188"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="6e4d2-102">Работа с содержимым страницы в OneNote</span><span class="sxs-lookup"><span data-stu-id="6e4d2-102">Work with OneNote page content</span></span>

<span data-ttu-id="6e4d2-103">В API JavaScript для надстроек OneNote содержимое страницы представлено указанной ниже объектной моделью.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Схема объектной модели страницы OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="6e4d2-105">Объект Page содержит коллекцию объектов PageContent.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="6e4d2-106">Объект PageContent содержит контент типов Outline, Image или Other.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="6e4d2-107">Объект Outline содержит коллекцию объектов Paragraph.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="6e4d2-108">Объект Paragraph содержит контент типов RichText, Image, Table или Other.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="6e4d2-109">Чтобы создать пустую страницу OneNote, воспользуйтесь одним из указанных ниже методов.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="6e4d2-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="6e4d2-110">Section.addPage</span></span>](/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="6e4d2-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="6e4d2-111">Page.insertPageAsSibling</span></span>](/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="6e4d2-112">Затем используйте методы в указанных ниже объектах для работы с содержимым страницы, например `Page.addOutline` и `Outline.appendHtml`.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-112">Then use methods in the following objects to work with the page content, such as `Page.addOutline` and `Outline.appendHtml`.</span></span>

- [<span data-ttu-id="6e4d2-113">Страница</span><span class="sxs-lookup"><span data-stu-id="6e4d2-113">Page</span></span>](/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="6e4d2-114">Outline</span><span class="sxs-lookup"><span data-stu-id="6e4d2-114">Outline</span></span>](/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="6e4d2-115">Paragraph</span><span class="sxs-lookup"><span data-stu-id="6e4d2-115">Paragraph</span></span>](/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="6e4d2-p101">Для представления содержимого и структуры страницы OneNote используется HTML. Для создания или обновления содержимого страницы поддерживается только подмножество HTML, как описано ниже.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="6e4d2-118">Поддерживаемые элементы HTML</span><span class="sxs-lookup"><span data-stu-id="6e4d2-118">Supported HTML</span></span>

<span data-ttu-id="6e4d2-119">Для создания и обновления содержимого страницы в API JavaScript для надстроек OneNote используются указанные ниже элементы HTML.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="6e4d2-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="6e4d2-120"></span></span>
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="6e4d2-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="6e4d2-121"></span></span>
- <span data-ttu-id="6e4d2-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="6e4d2-122"></span></span>
- <span data-ttu-id="6e4d2-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="6e4d2-123"></span></span>
- <span data-ttu-id="6e4d2-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="6e4d2-124"></span></span>

> [!NOTE]
> <span data-ttu-id="6e4d2-125">Импорт HTML в OneNote консолидирует пробелы.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-125">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="6e4d2-126">Полученное в результате содержимое вставляется в одну структуру.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-126">The resulting content is pasted into one outline.</span></span>

<span data-ttu-id="6e4d2-127">Приложение OneNote пытается наилучшим образом преобразовать HTML в содержимое страницы, обеспечивая безопасность для пользователей.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-127">OneNote does its best to translate HTML into page content while ensuring security for users.</span></span> <span data-ttu-id="6e4d2-128">Так как стандарты HTML и CSS не полностью соответствуют модели содержимого OneNote, будут иметься различия во внешнем виде, особенно при использовании стилей CSS.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-128">HTML and CSS standards do not exactly match OneNote's content model, so there will be differences in appearances, particularly with CSS stylings.</span></span> <span data-ttu-id="6e4d2-129">Рекомендуется использовать объекты JavaScript, если требуется определенное форматирование.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-129">We recommend using the JavaScript objects if specific formatting is needed.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="6e4d2-130">Доступ к содержимому страницы</span><span class="sxs-lookup"><span data-stu-id="6e4d2-130">Accessing page contents</span></span>

<span data-ttu-id="6e4d2-p104">Через \*\* доступ можно получить только к `Page#load`. Чтобы изменить активную страницу, вызовите команду `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-p104">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="6e4d2-133">Метаданные, например "Название", можно запросить для любой страницы.</span><span class="sxs-lookup"><span data-stu-id="6e4d2-133">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="6e4d2-134">См. также</span><span class="sxs-lookup"><span data-stu-id="6e4d2-134">See also</span></span>

- [<span data-ttu-id="6e4d2-135">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="6e4d2-135">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="6e4d2-136">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="6e4d2-136">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="6e4d2-137">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="6e4d2-137">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="6e4d2-138">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="6e4d2-138">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
