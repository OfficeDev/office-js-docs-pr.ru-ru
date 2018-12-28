---
title: Работа с содержимым страницы в OneNote
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: aef9d80ebb37dacd2c3b5f2ec9d33cb0164d8452
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457616"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="e4e0b-102">Работа с содержимым страницы в OneNote</span><span class="sxs-lookup"><span data-stu-id="e4e0b-102">Work with OneNote page content</span></span> 

<span data-ttu-id="e4e0b-103">В API JavaScript для надстроек OneNote содержимое страницы представлено указанной ниже объектной моделью.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Схема объектной модели страницы OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="e4e0b-105">Объект Page содержит коллекцию объектов PageContent.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="e4e0b-106">Объект PageContent содержит контент типов Outline, Image или Other.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="e4e0b-107">Объект Outline содержит коллекцию объектов Paragraph.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="e4e0b-108">Объект Paragraph содержит контент типов RichText, Image, Table или Other.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="e4e0b-109">Чтобы создать пустую страницу OneNote, воспользуйтесь одним из указанных ниже методов.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="e4e0b-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="e4e0b-110">Section.addPage</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="e4e0b-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="e4e0b-111">Page.insertPageAsSibling</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="e4e0b-112">Затем используйте методы в указанных ниже объектах для работы с содержимым страницы, например Page.addOutline и Outline.appendHtml.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="e4e0b-113">Страница</span><span class="sxs-lookup"><span data-stu-id="e4e0b-113">Page</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="e4e0b-114">Структура</span><span class="sxs-lookup"><span data-stu-id="e4e0b-114">Outline</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="e4e0b-115">Абзац</span><span class="sxs-lookup"><span data-stu-id="e4e0b-115">Paragraph</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="e4e0b-p101">Для представления содержимого и структуры страницы OneNote используется HTML. Для создания или обновления содержимого страницы поддерживается только подмножество HTML, как описано ниже.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="e4e0b-118">Поддерживаемые элементы HTML</span><span class="sxs-lookup"><span data-stu-id="e4e0b-118">Supported HTML</span></span>

<span data-ttu-id="e4e0b-119">Для создания и обновления содержимого страницы в API JavaScript для надстроек OneNote используются указанные ниже элементы HTML.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="e4e0b-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="e4e0b-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="e4e0b-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="e4e0b-121">`<ul>`, `<ol>`, `<li>`</span></span> 
- <span data-ttu-id="e4e0b-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="e4e0b-122">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="e4e0b-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="e4e0b-123">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="e4e0b-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="e4e0b-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

> [!NOTE]
> <span data-ttu-id="e4e0b-125">Импорт HTML в OneNote консолидирует пробелы.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-125">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="e4e0b-126">Полученное в результате содержимое вставляется в одну структуру.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-126">The resulting content is pasted into one outline.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="e4e0b-127">Доступ к содержимому страницы</span><span class="sxs-lookup"><span data-stu-id="e4e0b-127">Accessing page contents</span></span>

<span data-ttu-id="e4e0b-p103">Через `Page#load` доступ можно получить только к *содержимому активной страницы*. Чтобы изменить активную страницу, вызовите команду `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-p103">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="e4e0b-130">Метаданные, например "Название", можно запросить для любой страницы.</span><span class="sxs-lookup"><span data-stu-id="e4e0b-130">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="e4e0b-131">См. также</span><span class="sxs-lookup"><span data-stu-id="e4e0b-131">See also</span></span>

- [<span data-ttu-id="e4e0b-132">Обзор создания кода с помощью API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="e4e0b-132">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="e4e0b-133">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="e4e0b-133">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="e4e0b-134">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="e4e0b-134">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="e4e0b-135">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="e4e0b-135">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
